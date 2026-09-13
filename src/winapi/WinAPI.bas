Attribute VB_Name = "WinAPI"
Option Explicit

' High-level WinAPI helpers. Requires WinAPI_Core.bas.
' Import both .bas files into your VBA project.
'
' Usage:
'   WinAPI.WinSleep 1000
'   Debug.Print WinAPI.WinUserName
'   Debug.Print WinAPI.RegReadValue(HKCU, "Environment", "TEMP")
'   WinAPI.ClipSetText "copied"
'   Debug.Print WinAPI.RunAndCapture("ipconfig /all")

' --- EnumWindows callback storage ---
Private m_EnumWin() As Variant
Private m_EnumWinCount As Long

' ===================== PRIVATE HELPERS =====================

Private Function UToDouble(ByVal v As Long) As Double
    If v < 0 Then UToDouble = v + 4294967296# Else UToDouble = v
End Function

Private Function LargeIntToDouble(ByRef li As LARGE_INTEGER) As Double
    LargeIntToDouble = CDbl(li.HighPart) * 4294967296# + UToDouble(li.LowPart)
End Function

Private Function TrimNull(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, vbNullChar)
    If p > 0 Then TrimNull = Left$(s, p - 1) Else TrimNull = s
End Function

Private Function ByteBufToStr(ByRef buf() As Byte, ByVal length As Long) As String
    If length <= 0 Then Exit Function
    Dim tmp() As Byte
    ReDim tmp(0 To length - 1)
    CopyMemory VarPtr(tmp(0)), VarPtr(buf(0)), length
    ByteBufToStr = StrConv(tmp, vbUnicode)
End Function

Private Function IsArrayEmpty(ByRef arr() As Byte) As Boolean
    On Error Resume Next
    Dim lb As Long
    lb = LBound(arr)
    IsArrayEmpty = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function ApiStrBuf(ByVal charCount As Long) As String
    ApiStrBuf = Space$(charCount)
End Function

' ===================== ERROR HANDLING =====================

Public Function GetLastWinError() As String
    Dim errCode As Long
    errCode = GetLastError()
    If errCode = 0 Then Exit Function

    Dim buf As String
    buf = Space$(1024)
    Dim length As Long
    length = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                            0, errCode, 0, StrPtr(buf), 1024, 0)
    If length > 0 Then
        GetLastWinError = "Error " & errCode & ": " & Left$(buf, length)
    Else
        GetLastWinError = "Error " & errCode
    End If
End Function

Public Sub ThrowLastWinError(Optional ByVal context As String = "")
    Dim msg As String
    msg = GetLastWinError()
    If Len(msg) = 0 Then Exit Sub
    If Len(context) > 0 Then msg = context & " - " & msg
    Err.Raise vbObjectError + 1000, "WinAPI", msg
End Sub

' ===================== FILE OPERATIONS =====================

Public Function FileReadBytes(ByVal path As String) As Byte()
    Dim hFile As LongPtr
    hFile = CreateFileW(StrPtr(path), GENERIC_READ, FILE_SHARE_READ, 0, _
                        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then Exit Function

    Dim li As LARGE_INTEGER
    If GetFileSizeEx(hFile, li) = 0 Then
        CloseHandle hFile
        Exit Function
    End If

    Dim fileSize As Long
    fileSize = li.LowPart
    If li.HighPart <> 0 Or fileSize < 0 Then
        CloseHandle hFile
        Exit Function
    End If

    If fileSize = 0 Then
        CloseHandle hFile
        Dim empty(0 To 0) As Byte
        FileReadBytes = empty
        Exit Function
    End If

    Dim buf() As Byte
    ReDim buf(0 To fileSize - 1)
    Dim bytesRead As Long
    ReadFile hFile, VarPtr(buf(0)), fileSize, bytesRead, 0
    CloseHandle hFile

    If bytesRead < fileSize Then ReDim Preserve buf(0 To bytesRead - 1)
    FileReadBytes = buf
End Function

Public Function FileWriteBytes(ByVal path As String, ByRef data() As Byte) As Boolean
    Dim hFile As LongPtr
    hFile = CreateFileW(StrPtr(path), GENERIC_WRITE, 0, 0, _
                        CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then Exit Function

    Dim bytesWritten As Long
    Dim dataLen As Long
    dataLen = UBound(data) - LBound(data) + 1

    FileWriteBytes = (WriteFile(hFile, VarPtr(data(LBound(data))), dataLen, bytesWritten, 0) <> 0)
    FlushFileBuffers hFile
    CloseHandle hFile
End Function

Public Function FileReadText(ByVal path As String) As String
    Dim b() As Byte
    b = FileReadBytes(path)
    If IsArrayEmpty(b) Then Exit Function

    Dim dataLen As Long
    dataLen = UBound(b) - LBound(b) + 1
    If dataLen = 0 Then Exit Function

    Dim start As Long: start = 0

    ' UTF-16 LE BOM
    If dataLen >= 2 And b(0) = &HFF And b(1) = &HFE Then
        Dim uLen As Long
        uLen = dataLen - 2
        If uLen > 0 Then
            FileReadText = Space$(uLen \ 2)
            CopyMemory StrPtr(FileReadText), VarPtr(b(2)), CLngPtr(uLen)
        End If
        Exit Function
    End If

    ' UTF-8 BOM
    If dataLen >= 3 And b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then start = 3

    Dim length As Long
    length = dataLen - start
    If length <= 0 Then Exit Function

    Dim wLen As Long
    wLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(start)), length, 0, 0)
    If wLen > 0 Then
        FileReadText = Space$(wLen)
        MultiByteToWideChar CP_UTF8, 0, VarPtr(b(start)), length, StrPtr(FileReadText), wLen
    Else
        FileReadText = StrConv(b, vbUnicode)
    End If
End Function

Public Function FileWriteText(ByVal path As String, ByVal text As String, _
                              Optional ByVal utf8 As Boolean = True) As Boolean
    Dim buf() As Byte

    If utf8 Then
        Dim length As Long
        length = WideCharToMultiByte(CP_UTF8, 0, StrPtr(text), Len(text), 0, 0, 0, 0)
        If length = 0 And Len(text) > 0 Then Exit Function
        If length = 0 Then
            ReDim buf(0 To 0)
        Else
            ReDim buf(0 To length - 1)
            WideCharToMultiByte CP_UTF8, 0, StrPtr(text), Len(text), VarPtr(buf(0)), length, 0, 0
        End If
    Else
        buf = StrConv(text, vbFromUnicode)
    End If

    FileWriteText = FileWriteBytes(path, buf)
End Function

Public Function WinCopyFile(ByVal src As String, ByVal dst As String, _
                            Optional ByVal overwrite As Boolean = True) As Boolean
    WinCopyFile = (CopyFileW(StrPtr(src), StrPtr(dst), IIf(overwrite, 0, 1)) <> 0)
End Function

Public Function WinMoveFile(ByVal src As String, ByVal dst As String) As Boolean
    WinMoveFile = (MoveFileExW(StrPtr(src), StrPtr(dst), _
                   MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED) <> 0)
End Function

Public Function WinDeleteFile(ByVal path As String) As Boolean
    WinDeleteFile = (DeleteFileW(StrPtr(path)) <> 0)
End Function

Public Function WinFileExists(ByVal path As String) As Boolean
    Dim attr As Long
    attr = GetFileAttributesW(StrPtr(path))
    WinFileExists = (attr <> INVALID_FILE_ATTRIBUTES) And ((attr And FILE_ATTRIBUTE_DIRECTORY) = 0)
End Function

Public Function WinFolderExists(ByVal path As String) As Boolean
    Dim attr As Long
    attr = GetFileAttributesW(StrPtr(path))
    WinFolderExists = (attr <> INVALID_FILE_ATTRIBUTES) And ((attr And FILE_ATTRIBUTE_DIRECTORY) <> 0)
End Function

Public Function WinCreateFolder(ByVal path As String, _
                                Optional ByVal recursive As Boolean = True) As Boolean
    If WinFolderExists(path) Then
        WinCreateFolder = True
        Exit Function
    End If

    If recursive Then
        Dim parent As String
        Dim sep As Long
        sep = InStrRev(path, "\")
        If sep > 0 Then
            parent = Left$(path, sep - 1)
            If Len(parent) > 0 And Not WinFolderExists(parent) Then
                If Not WinCreateFolder(parent, True) Then Exit Function
            End If
        End If
    End If

    WinCreateFolder = (CreateDirectoryW(StrPtr(path), 0) <> 0)
End Function

Public Function WinFileSize(ByVal path As String) As Double
    WinFileSize = -1
    Dim hFile As LongPtr
    hFile = CreateFileW(StrPtr(path), 0, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, _
                        0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then Exit Function

    Dim li As LARGE_INTEGER
    If GetFileSizeEx(hFile, li) <> 0 Then
        WinFileSize = CDbl(li.HighPart) * 4294967296# + UToDouble(li.LowPart)
    End If
    CloseHandle hFile
End Function

Public Function WinTempPath() As String
    Dim buf As String: buf = Space$(MAX_PATH + 1)
    Dim ret As Long
    ret = GetTempPathW(Len(buf), StrPtr(buf))
    If ret > 0 Then WinTempPath = Left$(buf, ret)
End Function

Public Function WinTempFile(Optional ByVal prefix As String = "vba") As String
    Dim tmpDir As String: tmpDir = WinTempPath()
    If Len(tmpDir) = 0 Then Exit Function

    Dim buf As String: buf = Space$(MAX_PATH + 1)
    If GetTempFileNameW(StrPtr(tmpDir), StrPtr(prefix), 0, StrPtr(buf)) <> 0 Then
        WinTempFile = TrimNull(buf)
    End If
End Function

Public Function WinListFiles(ByVal folder As String, _
                             Optional ByVal pattern As String = "*") As Collection
    Set WinListFiles = New Collection

    Dim searchPath As String
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    searchPath = folder & pattern

    Dim fd As WIN32_FIND_DATA
    Dim hFind As LongPtr
    hFind = FindFirstFileW(StrPtr(searchPath), fd)
    If hFind = INVALID_HANDLE_VALUE Then Exit Function

    Do
        Dim fname As String
        fname = TrimNull(fd.cFileName)
        If fname <> "." And fname <> ".." Then
            WinListFiles.Add folder & fname
        End If
    Loop While FindNextFileW(hFind, fd) <> 0
    FindClose hFind
End Function

' ===================== REGISTRY =====================

Public Function RegReadValue(ByVal hive As LongPtr, ByVal subKey As String, _
                             ByVal valueName As String, _
                             Optional ByVal defaultValue As Variant = Empty) As Variant
    Dim hKey As LongPtr
    If RegOpenKeyExW(hive, StrPtr(subKey), 0, KEY_READ, hKey) <> ERROR_SUCCESS Then
        RegReadValue = defaultValue
        Exit Function
    End If

    Dim valueType As Long, dataSize As Long
    If RegQueryValueExW(hKey, StrPtr(valueName), 0, valueType, 0, dataSize) <> ERROR_SUCCESS Then
        RegCloseKey hKey
        RegReadValue = defaultValue
        Exit Function
    End If

    If dataSize = 0 Then
        RegCloseKey hKey
        RegReadValue = defaultValue
        Exit Function
    End If

    Dim buf() As Byte
    ReDim buf(0 To dataSize - 1)
    If RegQueryValueExW(hKey, StrPtr(valueName), 0, valueType, VarPtr(buf(0)), dataSize) <> ERROR_SUCCESS Then
        RegCloseKey hKey
        RegReadValue = defaultValue
        Exit Function
    End If
    RegCloseKey hKey

    Select Case valueType
        Case REG_SZ, REG_EXPAND_SZ
            Dim charCount As Long
            charCount = (dataSize \ 2) - 1
            If charCount > 0 Then
                Dim strResult As String
                strResult = Space$(charCount)
                CopyMemory StrPtr(strResult), VarPtr(buf(0)), CLngPtr(charCount * 2)
                RegReadValue = strResult
            Else
                RegReadValue = ""
            End If
        Case REG_DWORD
            Dim dw As Long
            CopyMemory VarPtr(dw), VarPtr(buf(0)), 4
            RegReadValue = dw
        Case REG_QWORD
            Dim qwLow As Long, qwHigh As Long
            CopyMemory VarPtr(qwLow), VarPtr(buf(0)), 4
            CopyMemory VarPtr(qwHigh), VarPtr(buf(4)), 4
            RegReadValue = CDbl(qwHigh) * 4294967296# + UToDouble(qwLow)
        Case REG_BINARY
            RegReadValue = buf
        Case REG_MULTI_SZ
            Dim msCharCount As Long
            msCharCount = (dataSize \ 2) - 2
            If msCharCount > 0 Then
                Dim msResult As String
                msResult = Space$(msCharCount)
                CopyMemory StrPtr(msResult), VarPtr(buf(0)), CLngPtr(msCharCount * 2)
                RegReadValue = Split(msResult, vbNullChar)
            Else
                RegReadValue = Array()
            End If
        Case Else
            RegReadValue = buf
    End Select
End Function

Public Function RegWriteValue(ByVal hive As LongPtr, ByVal subKey As String, _
                              ByVal valueName As String, ByVal value As Variant, _
                              Optional ByVal valueType As Long = -1) As Boolean
    Dim hKey As LongPtr, disposition As Long
    If RegCreateKeyExW(hive, StrPtr(subKey), 0, 0, 0, KEY_WRITE, 0, hKey, disposition) <> ERROR_SUCCESS Then
        Exit Function
    End If

    If valueType = -1 Then
        Select Case VarType(value)
            Case vbString:    valueType = REG_SZ
            Case vbLong, vbInteger, vbByte, vbBoolean: valueType = REG_DWORD
            Case Else:        valueType = REG_SZ
        End Select
    End If

    Dim ret As Long
    Select Case valueType
        Case REG_SZ, REG_EXPAND_SZ
            Dim s As String: s = CStr(value) & vbNullChar
            ret = RegSetValueExW(hKey, StrPtr(valueName), 0, valueType, StrPtr(s), LenB(s))
        Case REG_DWORD
            Dim dwVal As Long: dwVal = CLng(value)
            ret = RegSetValueExW(hKey, StrPtr(valueName), 0, REG_DWORD, VarPtr(dwVal), 4)
        Case REG_BINARY
            Dim binBuf() As Byte: binBuf = value
            ret = RegSetValueExW(hKey, StrPtr(valueName), 0, REG_BINARY, _
                                 VarPtr(binBuf(LBound(binBuf))), UBound(binBuf) - LBound(binBuf) + 1)
        Case Else
            ret = 1
    End Select

    RegCloseKey hKey
    RegWriteValue = (ret = ERROR_SUCCESS)
End Function

Public Function RegDeleteValue(ByVal hive As LongPtr, ByVal subKey As String, _
                               ByVal valueName As String) As Boolean
    Dim hKey As LongPtr
    If RegOpenKeyExW(hive, StrPtr(subKey), 0, KEY_SET_VALUE, hKey) <> ERROR_SUCCESS Then Exit Function
    RegDeleteValue = (RegDeleteValueW(hKey, StrPtr(valueName)) = ERROR_SUCCESS)
    RegCloseKey hKey
End Function

Public Function RegDeleteSubKey(ByVal hive As LongPtr, ByVal subKey As String) As Boolean
    RegDeleteSubKey = (RegDeleteKeyW(hive, StrPtr(subKey)) = ERROR_SUCCESS)
End Function

Public Function RegKeyExists(ByVal hive As LongPtr, ByVal subKey As String) As Boolean
    Dim hKey As LongPtr
    If RegOpenKeyExW(hive, StrPtr(subKey), 0, KEY_READ, hKey) = ERROR_SUCCESS Then
        RegCloseKey hKey
        RegKeyExists = True
    End If
End Function

Public Function RegValueExists(ByVal hive As LongPtr, ByVal subKey As String, _
                               ByVal valueName As String) As Boolean
    Dim hKey As LongPtr
    If RegOpenKeyExW(hive, StrPtr(subKey), 0, KEY_QUERY_VALUE, hKey) <> ERROR_SUCCESS Then Exit Function
    Dim valueType As Long, dataSize As Long
    RegValueExists = (RegQueryValueExW(hKey, StrPtr(valueName), 0, valueType, 0, dataSize) = ERROR_SUCCESS)
    RegCloseKey hKey
End Function

Public Function RegEnumKeys(ByVal hive As LongPtr, ByVal subKey As String) As Collection
    Set RegEnumKeys = New Collection
    Dim hKey As LongPtr
    If RegOpenKeyExW(hive, StrPtr(subKey), 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> ERROR_SUCCESS Then Exit Function

    Dim idx As Long, nameBuf As String, nameLen As Long
    Do
        nameBuf = Space$(256)
        nameLen = 256
        If RegEnumKeyExW(hKey, idx, StrPtr(nameBuf), nameLen, 0, 0, 0, 0) <> ERROR_SUCCESS Then Exit Do
        RegEnumKeys.Add Left$(nameBuf, nameLen)
        idx = idx + 1
    Loop
    RegCloseKey hKey
End Function

Public Function RegEnumValues(ByVal hive As LongPtr, ByVal subKey As String) As Collection
    Set RegEnumValues = New Collection
    Dim hKey As LongPtr
    If RegOpenKeyExW(hive, StrPtr(subKey), 0, KEY_QUERY_VALUE, hKey) <> ERROR_SUCCESS Then Exit Function

    Dim idx As Long, nameBuf As String, nameLen As Long, valueType As Long
    Do
        nameBuf = Space$(16384)
        nameLen = 16384
        If RegEnumValueW(hKey, idx, StrPtr(nameBuf), nameLen, 0, valueType, 0, 0) <> ERROR_SUCCESS Then Exit Do
        RegEnumValues.Add Left$(nameBuf, nameLen)
        idx = idx + 1
    Loop
    RegCloseKey hKey
End Function

' ===================== PROCESS =====================

Public Function RunCommand(ByVal cmd As String, _
                           Optional ByVal waitForExit As Boolean = True, _
                           Optional ByVal hideWindow As Boolean = True) As Long
    RunCommand = -1
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    si.cb = LenB(si)
    If hideWindow Then
        si.dwFlags = STARTF_USESHOWWINDOW
        si.wShowWindow = SW_HIDE
    End If

    Dim cmdLine As String: cmdLine = "cmd.exe /c " & cmd

    If CreateProcessW(0, StrPtr(cmdLine), 0, 0, 0, CREATE_NO_WINDOW, 0, 0, si, pi) = 0 Then Exit Function

    If waitForExit Then
        WaitForSingleObject pi.hProcess, INFINITE_WAIT
        GetExitCodeProcess pi.hProcess, RunCommand
    Else
        RunCommand = 0
    End If

    CloseHandle pi.hThread
    CloseHandle pi.hProcess
End Function

Public Function RunAndCapture(ByVal cmd As String) As String
    Dim sa As SECURITY_ATTRIBUTES
    sa.nLength = LenB(sa)
    sa.bInheritHandle = 1

    Dim hReadPipe As LongPtr, hWritePipe As LongPtr
    If CreatePipe(hReadPipe, hWritePipe, sa, 0) = 0 Then Exit Function
    SetHandleInformation hReadPipe, HANDLE_FLAG_INHERIT, 0

    Dim si As STARTUPINFO
    si.cb = LenB(si)
    si.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
    si.wShowWindow = SW_HIDE
    si.hStdOutput = hWritePipe
    si.hStdError = hWritePipe

    Dim pi As PROCESS_INFORMATION
    Dim cmdLine As String: cmdLine = "cmd.exe /c " & cmd

    If CreateProcessW(0, StrPtr(cmdLine), 0, 0, 1, CREATE_NO_WINDOW, 0, 0, si, pi) = 0 Then
        CloseHandle hReadPipe
        CloseHandle hWritePipe
        Exit Function
    End If

    CloseHandle hWritePipe
    CloseHandle pi.hThread

    Dim buf(0 To 4095) As Byte
    Dim bytesRead As Long
    Dim result As String

    Do
        If ReadFile(hReadPipe, VarPtr(buf(0)), 4096, bytesRead, 0) = 0 Then Exit Do
        If bytesRead = 0 Then Exit Do
        result = result & ByteBufToStr(buf, bytesRead)
    Loop

    WaitForSingleObject pi.hProcess, INFINITE_WAIT
    CloseHandle pi.hProcess
    CloseHandle hReadPipe

    RunAndCapture = result
End Function

Public Function ShellOpen(ByVal path As String, _
                          Optional ByVal args As String = "", _
                          Optional ByVal verb As String = "open") As Boolean
    Dim ret As LongPtr
    Dim pArgs As LongPtr: If Len(args) > 0 Then pArgs = StrPtr(args)
    ret = ShellExecuteW(0, StrPtr(verb), StrPtr(path), pArgs, 0, SW_SHOWNORMAL)
    ShellOpen = (ret > 32)
End Function

Public Function GetProcessList() As Variant
    Dim wmi As Object, procs As Object, proc As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT ProcessId, Name FROM Win32_Process")

    If procs.Count = 0 Then
        GetProcessList = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To procs.Count, 1 To 2)
    Dim i As Long: i = 1
    For Each proc In procs
        result(i, 1) = proc.ProcessId
        result(i, 2) = proc.Name
        i = i + 1
    Next
    GetProcessList = result
End Function

Public Function ProcessExists(ByVal nameOrPid As Variant) As Boolean
    If IsNumeric(nameOrPid) Then
        Dim h As LongPtr
        h = OpenProcess(PROCESS_QUERY_INFORMATION, 0, CLng(nameOrPid))
        If h <> 0 Then
            CloseHandle h
            ProcessExists = True
        End If
    Else
        Dim wmi As Object
        Set wmi = GetObject("winmgmts:\\.\root\cimv2")
        Dim procs As Object
        Set procs = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='" & _
                                   Replace(CStr(nameOrPid), "'", "''") & "'")
        ProcessExists = (procs.Count > 0)
    End If
End Function

Public Function KillProcess(ByVal nameOrPid As Variant) As Boolean
    If IsNumeric(nameOrPid) Then
        Dim h As LongPtr
        h = OpenProcess(PROCESS_TERMINATE, 0, CLng(nameOrPid))
        If h = 0 Then Exit Function
        KillProcess = (TerminateProcess(h, 1) <> 0)
        CloseHandle h
    Else
        Dim wmi As Object
        Set wmi = GetObject("winmgmts:\\.\root\cimv2")
        Dim procs As Object, proc As Object
        Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & _
                                   Replace(CStr(nameOrPid), "'", "''") & "'")
        For Each proc In procs
            proc.Terminate
            KillProcess = True
        Next
    End If
End Function

' ===================== WINDOW =====================

Public Function FindWindowByTitle(ByVal partialTitle As String, _
                                  Optional ByVal exactMatch As Boolean = False) As LongPtr
    If exactMatch Then
        FindWindowByTitle = FindWindowW(0, StrPtr(partialTitle))
    Else
        Dim wins As Variant
        wins = GetAllWindows()
        If IsEmpty(wins) Then Exit Function
        Dim i As Long
        For i = LBound(wins, 1) To UBound(wins, 1)
            If InStr(1, CStr(wins(i, 2)), partialTitle, vbTextCompare) > 0 Then
                FindWindowByTitle = wins(i, 1)
                Exit Function
            End If
        Next
    End If
End Function

Public Function FindWindowByClass(ByVal className As String) As LongPtr
    FindWindowByClass = FindWindowW(StrPtr(className), 0)
End Function

Public Function GetWinText(ByVal hWnd As LongPtr) As String
    Dim length As Long
    length = GetWindowTextLengthW(hWnd)
    If length = 0 Then Exit Function
    Dim buf As String: buf = Space$(length + 1)
    GetWindowTextW hWnd, StrPtr(buf), length + 1
    GetWinText = Left$(buf, length)
End Function

Public Function SetWinText(ByVal hWnd As LongPtr, ByVal text As String) As Boolean
    SetWinText = (SetWindowTextW(hWnd, StrPtr(text)) <> 0)
End Function

Public Function WinShow(ByVal hWnd As LongPtr) As Boolean
    WinShow = (ShowWindow(hWnd, SW_SHOW) <> 0) Or True
    ShowWindow hWnd, SW_SHOW
End Function

Public Function WinHide(ByVal hWnd As LongPtr) As Boolean
    ShowWindow hWnd, SW_HIDE
    WinHide = True
End Function

Public Function WinMinimize(ByVal hWnd As LongPtr) As Boolean
    ShowWindow hWnd, SW_MINIMIZE
    WinMinimize = True
End Function

Public Function WinMaximize(ByVal hWnd As LongPtr) As Boolean
    ShowWindow hWnd, SW_SHOWMAXIMIZED
    WinMaximize = True
End Function

Public Function WinRestore(ByVal hWnd As LongPtr) As Boolean
    ShowWindow hWnd, SW_RESTORE
    WinRestore = True
End Function

Public Function WinSetTopMost(ByVal hWnd As LongPtr, _
                              Optional ByVal topMost As Boolean = True) As Boolean
    Dim insertAfter As LongPtr
    If topMost Then insertAfter = HWND_TOPMOST Else insertAfter = HWND_NOTOPMOST
    WinSetTopMost = (SetWindowPos(hWnd, insertAfter, 0, 0, 0, 0, _
                     SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW) <> 0)
End Function

Public Function WinActivate(ByVal hWnd As LongPtr) As Boolean
    ShowWindow hWnd, SW_RESTORE
    WinActivate = (SetForegroundWindow(hWnd) <> 0)
End Function

Public Function WinMove(ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, _
                        Optional ByVal w As Long = -1, Optional ByVal h As Long = -1) As Boolean
    If w = -1 Or h = -1 Then
        Dim rc As RECT_API
        GetWindowRect hWnd, rc
        If w = -1 Then w = rc.Right - rc.Left
        If h = -1 Then h = rc.Bottom - rc.Top
    End If
    WinMove = (MoveWindow(hWnd, x, y, w, h, 1) <> 0)
End Function

Public Function WinGetRect(ByVal hWnd As LongPtr) As String
    Dim rc As RECT_API
    If GetWindowRect(hWnd, rc) <> 0 Then
        WinGetRect = rc.Left & "," & rc.Top & "," & (rc.Right - rc.Left) & "," & (rc.Bottom - rc.Top)
    End If
End Function

Public Function GetAllWindows() As Variant
    m_EnumWinCount = 0
    ReDim m_EnumWin(1 To 500, 1 To 4)

    EnumWindows AddressOf EnumWindowsProc, 0

    If m_EnumWinCount = 0 Then
        GetAllWindows = Empty
        Exit Function
    End If

    ReDim Preserve m_EnumWin(1 To m_EnumWinCount, 1 To 4)
    GetAllWindows = m_EnumWin
End Function

Public Function WinSendMsg(ByVal hWnd As LongPtr, ByVal msg As Long, _
                           ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    WinSendMsg = SendMessageW(hWnd, msg, wParam, lParam)
End Function

Public Function WinPostMsg(ByVal hWnd As LongPtr, ByVal msg As Long, _
                           ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Boolean
    WinPostMsg = (PostMessageW(hWnd, msg, wParam, lParam) <> 0)
End Function

Public Function WinCloseWindow(ByVal hWnd As LongPtr) As Boolean
    WinCloseWindow = (PostMessageW(hWnd, WM_CLOSE, 0, 0) <> 0)
End Function

' --- EnumWindows callback (must be Public in a standard module) ---
Public Function EnumWindowsProc(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
    EnumWindowsProc = 1

    If IsWindowVisible(hWnd) = 0 Then Exit Function

    Dim title As String: title = GetWinText(hWnd)
    If Len(title) = 0 Then Exit Function

    Dim className As String: className = Space$(256)
    Dim classLen As Long
    classLen = GetClassNameW(hWnd, StrPtr(className), 256)
    className = Left$(className, classLen)

    Dim pid As Long
    GetWindowThreadProcessId hWnd, pid

    m_EnumWinCount = m_EnumWinCount + 1
    If m_EnumWinCount > UBound(m_EnumWin, 1) Then
        ReDim Preserve m_EnumWin(1 To UBound(m_EnumWin, 1) + 500, 1 To 4)
    End If

    m_EnumWin(m_EnumWinCount, 1) = hWnd
    m_EnumWin(m_EnumWinCount, 2) = title
    m_EnumWin(m_EnumWinCount, 3) = className
    m_EnumWin(m_EnumWinCount, 4) = pid
End Function

' ===================== SYSTEM =====================

Public Function WinComputerName() As String
    Dim buf As String: buf = Space$(256)
    Dim nSize As Long: nSize = 256
    If GetComputerNameW(StrPtr(buf), nSize) <> 0 Then
        WinComputerName = Left$(buf, nSize)
    End If
End Function

Public Function WinUserName() As String
    Dim buf As String: buf = Space$(256)
    Dim nSize As Long: nSize = 256
    If GetUserNameW(StrPtr(buf), nSize) <> 0 Then
        WinUserName = Left$(buf, nSize - 1)
    End If
End Function

Public Function WinSystemDir() As String
    Dim buf As String: buf = Space$(MAX_PATH)
    Dim ret As Long
    ret = GetSystemDirectoryW(StrPtr(buf), MAX_PATH)
    If ret > 0 Then WinSystemDir = Left$(buf, ret)
End Function

Public Function WinWindowsDir() As String
    Dim buf As String: buf = Space$(MAX_PATH)
    Dim ret As Long
    ret = GetWindowsDirectoryW(StrPtr(buf), MAX_PATH)
    If ret > 0 Then WinWindowsDir = Left$(buf, ret)
End Function

Public Function WinGetEnv(ByVal name As String) As String
    Dim buf As String: buf = Space$(32767)
    Dim ret As Long
    ret = GetEnvironmentVariableW(StrPtr(name), StrPtr(buf), 32767)
    If ret > 0 Then WinGetEnv = Left$(buf, ret)
End Function

Public Function WinSetEnv(ByVal name As String, ByVal value As String) As Boolean
    WinSetEnv = (SetEnvironmentVariableW(StrPtr(name), StrPtr(value)) <> 0)
End Function

Public Sub WinSleep(ByVal milliseconds As Long)
    Sleep milliseconds
End Sub

Public Function WinMemoryInfo() As String
    Dim ms As MEMORYSTATUSEX
    ms.dwLength = LenB(ms)
    If GlobalMemoryStatusEx(ms) <> 0 Then
        WinMemoryInfo = "Load: " & ms.dwMemoryLoad & "%" & vbCrLf & _
                        "Physical: " & FormatNumber(ms.ullAvailPhys * 10000 / 1073741824#, 1) & _
                        " / " & FormatNumber(ms.ullTotalPhys * 10000 / 1073741824#, 1) & " GB"
    End If
End Function

' ===================== CLIPBOARD =====================

Public Function ClipGetText() As String
    If OpenClipboard(0) = 0 Then Exit Function

    Dim hMem As LongPtr
    If IsClipboardFormatAvailable(CF_UNICODETEXT) <> 0 Then
        hMem = GetClipboardData(CF_UNICODETEXT)
        If hMem <> 0 Then
            Dim ptr As LongPtr: ptr = GlobalLock(hMem)
            If ptr <> 0 Then
                Dim length As Long: length = lstrlenW(ptr)
                If length > 0 Then
                    ClipGetText = Space$(length)
                    CopyMemory StrPtr(ClipGetText), ptr, CLngPtr(length * 2)
                End If
                GlobalUnlock hMem
            End If
        End If
    ElseIf IsClipboardFormatAvailable(CF_TEXT) <> 0 Then
        hMem = GetClipboardData(CF_TEXT)
        If hMem <> 0 Then
            Dim ptrA As LongPtr: ptrA = GlobalLock(hMem)
            If ptrA <> 0 Then
                Dim sz As LongPtr: sz = GlobalSize(hMem)
                If sz > 0 Then
                    Dim buf() As Byte
                    ReDim buf(0 To CLng(sz) - 1)
                    CopyMemory VarPtr(buf(0)), ptrA, sz
                    ClipGetText = StrConv(buf, vbUnicode)
                    ClipGetText = TrimNull(ClipGetText)
                End If
                GlobalUnlock hMem
            End If
        End If
    End If

    CloseClipboard
End Function

Public Function ClipSetText(ByVal text As String) As Boolean
    If OpenClipboard(0) = 0 Then Exit Function
    EmptyClipboard

    Dim byteLen As Long: byteLen = (Len(text) + 1) * 2
    Dim hMem As LongPtr: hMem = GlobalAlloc(GHND, byteLen)
    If hMem = 0 Then
        CloseClipboard
        Exit Function
    End If

    Dim ptr As LongPtr: ptr = GlobalLock(hMem)
    If ptr <> 0 Then
        CopyMemory ptr, StrPtr(text), CLngPtr(Len(text) * 2)
        GlobalUnlock hMem
        ClipSetText = (SetClipboardData(CF_UNICODETEXT, hMem) <> 0)
    End If

    CloseClipboard
End Function

' ===================== MEMORY (cross-process) =====================

Public Function MemAlloc(ByVal size As LongPtr, _
                         Optional ByVal executable As Boolean = False) As LongPtr
    Dim protect As Long
    If executable Then protect = PAGE_EXECUTE_READWRITE Else protect = PAGE_READWRITE
    MemAlloc = VirtualAlloc(0, size, MEM_COMMIT Or MEM_RESERVE, protect)
End Function

Public Function MemFree(ByVal addr As LongPtr) As Boolean
    MemFree = (VirtualFree(addr, 0, MEM_RELEASE) <> 0)
End Function

Public Function MemProtect(ByVal addr As LongPtr, ByVal size As LongPtr, _
                           ByVal newProtect As Long) As Long
    VirtualProtect addr, size, newProtect, MemProtect
End Function

Public Function MemRead(ByVal hProcess As LongPtr, ByVal addr As LongPtr, _
                        ByVal size As Long) As Byte()
    Dim buf() As Byte
    ReDim buf(0 To size - 1)
    Dim bytesRead As LongPtr
    If ReadProcessMemory(hProcess, addr, VarPtr(buf(0)), CLngPtr(size), bytesRead) = 0 Then
        Erase buf
    ElseIf CLng(bytesRead) < size Then
        ReDim Preserve buf(0 To CLng(bytesRead) - 1)
    End If
    MemRead = buf
End Function

Public Function MemWrite(ByVal hProcess As LongPtr, ByVal addr As LongPtr, _
                         ByRef data() As Byte) As Boolean
    Dim dataLen As Long: dataLen = UBound(data) - LBound(data) + 1
    Dim bytesWritten As LongPtr
    MemWrite = (WriteProcessMemory(hProcess, addr, VarPtr(data(LBound(data))), _
                CLngPtr(dataLen), bytesWritten) <> 0)
End Function

Public Function MemReadLong(ByVal hProcess As LongPtr, ByVal addr As LongPtr) As Long
    Dim bytesRead As LongPtr
    ReadProcessMemory hProcess, addr, VarPtr(MemReadLong), 4, bytesRead
End Function

Public Function MemWriteLong(ByVal hProcess As LongPtr, ByVal addr As LongPtr, _
                             ByVal value As Long) As Boolean
    Dim bytesWritten As LongPtr
    MemWriteLong = (WriteProcessMemory(hProcess, addr, VarPtr(value), 4, bytesWritten) <> 0)
End Function

Public Function MemReadString(ByVal hProcess As LongPtr, ByVal addr As LongPtr, _
                              ByVal maxLen As Long) As String
    Dim buf() As Byte
    buf = MemRead(hProcess, addr, maxLen)
    If IsArrayEmpty(buf) Then Exit Function

    Dim i As Long
    For i = 0 To UBound(buf)
        If buf(i) = 0 Then
            If i = 0 Then Exit Function
            ReDim Preserve buf(0 To i - 1)
            Exit For
        End If
    Next
    MemReadString = StrConv(buf, vbUnicode)
End Function

Public Function OpenProc(ByVal pid As Long, _
                         Optional ByVal access As Long = -1) As LongPtr
    If access = -1 Then access = PROCESS_ALL_ACCESS
    OpenProc = OpenProcess(access, 0, pid)
End Function

Public Sub CloseProc(ByVal hProcess As LongPtr)
    If hProcess <> 0 Then CloseHandle hProcess
End Sub

' ponytail: MemAllocEx/MemFreeEx for remote process alloc, add when needed

' ===================== TIMER =====================

Public Function TimerStart() As Double
    Dim t As LARGE_INTEGER, f As LARGE_INTEGER
    QueryPerformanceCounter t
    QueryPerformanceFrequency f
    TimerStart = LargeIntToDouble(t) / LargeIntToDouble(f)
End Function

Public Function TimerElapsed(ByVal startSeconds As Double) As Double
    TimerElapsed = TimerStart() - startSeconds
End Function

' ===================== UTILITY =====================

Public Function PtrToStr(ByVal ptr As LongPtr) As String
    If ptr = 0 Then Exit Function
    Dim length As Long: length = lstrlenW(ptr)
    If length = 0 Then Exit Function
    PtrToStr = Space$(length)
    CopyMemory StrPtr(PtrToStr), ptr, CLngPtr(length * 2)
End Function

' ===================== DEMO =====================

Public Sub Demo()
    Debug.Print "=== WinAPI Toolkit Demo ==="
    Debug.Print

    Debug.Print "Computer: " & WinComputerName()
    Debug.Print "User: " & WinUserName()
    Debug.Print "SystemDir: " & WinSystemDir()
    Debug.Print "TempPath: " & WinTempPath()
    Debug.Print "PATH (first 80): " & Left$(WinGetEnv("PATH"), 80) & "..."
    Debug.Print WinMemoryInfo()
    Debug.Print

    Debug.Print "--- Command capture ---"
    Debug.Print RunAndCapture("echo Hello from WinAPI && ver")

    Debug.Print "--- File ops ---"
    Dim tmpFile As String: tmpFile = WinTempFile("demo")
    Debug.Print "Temp file: " & tmpFile
    FileWriteText tmpFile, "Hello WinAPI!"
    Debug.Print "Read back: " & FileReadText(tmpFile)
    Debug.Print "Size: " & WinFileSize(tmpFile) & " bytes"
    Debug.Print "Exists: " & WinFileExists(tmpFile)
    WinDeleteFile tmpFile
    Debug.Print "After delete: " & WinFileExists(tmpFile)
    Debug.Print

    Debug.Print "--- Registry ---"
    Dim userEnv As String: userEnv = "Environment"
    Dim envKeys As Collection: Set envKeys = RegEnumValues(HKCU, userEnv)
    Debug.Print "User env vars: " & envKeys.Count
    Debug.Print "TEMP = " & RegReadValue(HKCU, userEnv, "TEMP", "(not set)")
    Debug.Print

    Debug.Print "--- Windows ---"
    Dim wins As Variant: wins = GetAllWindows()
    If Not IsEmpty(wins) Then
        Dim i As Long
        Debug.Print "Visible windows (first 5):"
        For i = 1 To Application.WorksheetFunction.Min(5, UBound(wins, 1))
            Debug.Print "  [" & wins(i, 4) & "] " & wins(i, 2)
        Next
    End If
    Debug.Print

    Debug.Print "--- Clipboard ---"
    Dim oldClip As String: oldClip = ClipGetText()
    ClipSetText "WinAPI test"
    Debug.Print "Clipboard: " & ClipGetText()
    If Len(oldClip) > 0 Then ClipSetText oldClip
    Debug.Print

    Debug.Print "--- Timer ---"
    Dim t As Double: t = TimerStart()
    WinSleep 100
    Debug.Print "100ms sleep took: " & Format$(TimerElapsed(t) * 1000, "0.00") & " ms"

    Debug.Print
    Debug.Print "=== Done ==="
End Sub
