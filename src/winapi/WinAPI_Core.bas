Attribute VB_Name = "WinAPI_Core"
Option Explicit

' Raw WinAPI declarations, types, and constants.
' Use alongside WinAPI.bas for high-level helpers, or call these directly.
' VBA7 (Office 2010+), 32-bit and 64-bit compatible.

' ===================== CONSTANTS =====================

Public Const NULL_PTR As LongPtr = 0
Public Const INVALID_HANDLE_VALUE As LongPtr = -1
Public Const MAX_PATH As Long = 260

Public Const WAIT_OBJECT_0 As Long = &H0
Public Const WAIT_TIMEOUT As Long = &H102
Public Const INFINITE_WAIT As Long = &HFFFFFFFF

' File access
Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const GENERIC_EXECUTE As Long = &H20000000
Public Const GENERIC_ALL As Long = &H10000000
Public Const FILE_SHARE_READ As Long = &H1
Public Const FILE_SHARE_WRITE As Long = &H2
Public Const FILE_SHARE_DELETE As Long = &H4
Public Const CREATE_NEW As Long = 1
Public Const CREATE_ALWAYS As Long = 2
Public Const OPEN_EXISTING As Long = 3
Public Const OPEN_ALWAYS As Long = 4
Public Const TRUNCATE_EXISTING As Long = 5
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const INVALID_FILE_ATTRIBUTES As Long = -1
Public Const MOVEFILE_REPLACE_EXISTING As Long = &H1
Public Const MOVEFILE_COPY_ALLOWED As Long = &H2

' FormatMessage
Public Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

' Process
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Public Const PROCESS_VM_READ As Long = &H10
Public Const PROCESS_VM_WRITE As Long = &H20
Public Const PROCESS_VM_OPERATION As Long = &H8
Public Const PROCESS_QUERY_INFORMATION As Long = &H400
Public Const PROCESS_TERMINATE As Long = &H1

' Memory
Public Const MEM_COMMIT As Long = &H1000
Public Const MEM_RESERVE As Long = &H2000
Public Const MEM_RELEASE As Long = &H8000
Public Const PAGE_READWRITE As Long = &H4
Public Const PAGE_EXECUTE_READWRITE As Long = &H40
Public Const PAGE_READONLY As Long = &H2
Public Const PAGE_EXECUTE_READ As Long = &H20

' Registry
Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003
Public Const HKCR As Long = &H80000000
Public Const HKCU As Long = &H80000001
Public Const HKLM As Long = &H80000002
Public Const HKU As Long = &H80000003
Public Const KEY_READ As Long = &H20019
Public Const KEY_WRITE As Long = &H20006
Public Const KEY_ALL_ACCESS As Long = &HF003F
Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const REG_NONE As Long = 0
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_QWORD As Long = 11
Public Const ERROR_SUCCESS As Long = 0
Public Const ERROR_FILE_NOT_FOUND As Long = 2
Public Const ERROR_MORE_DATA As Long = 234
Public Const ERROR_NO_MORE_ITEMS As Long = 259

' Window
Public Const SW_HIDE As Long = 0
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMINIMIZED As Long = 2
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOW As Long = 5
Public Const SW_MINIMIZE As Long = 6
Public Const SW_RESTORE As Long = 9
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST As Long = -2
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_SHOWWINDOW As Long = &H40
Public Const GWL_STYLE As Long = -16
Public Const GWL_EXSTYLE As Long = -20
Public Const WM_CLOSE As Long = &H10
Public Const WM_SETTEXT As Long = &HC
Public Const WM_GETTEXT As Long = &HD
Public Const WM_GETTEXTLENGTH As Long = &HE

' Clipboard
Public Const CF_TEXT As Long = 1
Public Const CF_UNICODETEXT As Long = 13

' CreateProcess
Public Const STARTF_USESHOWWINDOW As Long = &H1
Public Const STARTF_USESTDHANDLES As Long = &H100
Public Const CREATE_NO_WINDOW As Long = &H8000000
Public Const NORMAL_PRIORITY_CLASS As Long = &H20

' Toolhelp32
Public Const TH32CS_SNAPPROCESS As Long = &H2

' Global memory
Public Const GMEM_MOVEABLE As Long = &H2
Public Const GMEM_ZEROINIT As Long = &H40
Public Const GHND As Long = &H42

' Handle
Public Const HANDLE_FLAG_INHERIT As Long = &H1

' Code pages
Public Const CP_UTF8 As Long = 65001
Public Const CP_ACP As Long = 0

' ===================== TYPES =====================

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Public Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

Public Type RECT_API
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINT_API
    x As Long
    y As Long
End Type

Public Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Public Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As Currency
    ullAvailPhys As Currency
    ullTotalPageFile As Currency
    ullAvailPageFile As Currency
    ullTotalVirtual As Currency
    ullAvailVirtual As Currency
    ullAvailExtendedVirtual As Currency
End Type

' ===================== API DECLARES =====================

' --- kernel32: Error ---
Public Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Public Declare PtrSafe Function FormatMessageW Lib "kernel32" ( _
    ByVal dwFlags As Long, ByVal lpSource As LongPtr, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As LongPtr, ByVal nSize As Long, _
    ByVal Arguments As LongPtr) As Long

' --- kernel32: File ---
Public Declare PtrSafe Function CreateFileW Lib "kernel32" ( _
    ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr
Public Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
Public Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
Public Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Public Declare PtrSafe Function FlushFileBuffers Lib "kernel32" (ByVal hFile As LongPtr) As Long
Public Declare PtrSafe Function GetFileSizeEx Lib "kernel32" ( _
    ByVal hFile As LongPtr, ByRef lpFileSize As LARGE_INTEGER) As Long
Public Declare PtrSafe Function CopyFileW Lib "kernel32" ( _
    ByVal lpExistingFileName As LongPtr, ByVal lpNewFileName As LongPtr, _
    ByVal bFailIfExists As Long) As Long
Public Declare PtrSafe Function MoveFileExW Lib "kernel32" ( _
    ByVal lpExistingFileName As LongPtr, ByVal lpNewFileName As LongPtr, _
    ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function DeleteFileW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
Public Declare PtrSafe Function CreateDirectoryW Lib "kernel32" ( _
    ByVal lpPathName As LongPtr, ByVal lpSecurityAttributes As LongPtr) As Long
Public Declare PtrSafe Function RemoveDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As Long
Public Declare PtrSafe Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
Public Declare PtrSafe Function FindFirstFileW Lib "kernel32" ( _
    ByVal lpFileName As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATA) As LongPtr
Public Declare PtrSafe Function FindNextFileW Lib "kernel32" ( _
    ByVal hFindFile As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long
Public Declare PtrSafe Function GetTempPathW Lib "kernel32" ( _
    ByVal nBufferLength As Long, ByVal lpBuffer As LongPtr) As Long
Public Declare PtrSafe Function GetTempFileNameW Lib "kernel32" ( _
    ByVal lpPathName As LongPtr, ByVal lpPrefixString As LongPtr, _
    ByVal uUnique As Long, ByVal lpTempFileName As LongPtr) As Long

' --- kernel32: Process ---
Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As LongPtr
Public Declare PtrSafe Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
Public Declare PtrSafe Function CreateProcessW Lib "kernel32" ( _
    ByVal lpApplicationName As LongPtr, ByVal lpCommandLine As LongPtr, _
    ByVal lpProcessAttributes As LongPtr, ByVal lpThreadAttributes As LongPtr, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As LongPtr, _
    ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare PtrSafe Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" ( _
    ByVal hProcess As LongPtr, ByRef lpExitCode As Long) As Long
Public Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr

' --- kernel32: Memory ---
Public Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Public Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long
Public Declare PtrSafe Function VirtualProtect Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Public Declare PtrSafe Function VirtualAllocEx Lib "kernel32" ( _
    ByVal hProcess As LongPtr, ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr
Public Declare PtrSafe Function VirtualFreeEx Lib "kernel32" ( _
    ByVal hProcess As LongPtr, ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, ByVal dwFreeType As Long) As Long
Public Declare PtrSafe Function ReadProcessMemory Lib "kernel32" ( _
    ByVal hProcess As LongPtr, ByVal lpBaseAddress As LongPtr, _
    ByVal lpBuffer As LongPtr, ByVal nSize As LongPtr, _
    ByRef lpNumberOfBytesRead As LongPtr) As Long
Public Declare PtrSafe Function WriteProcessMemory Lib "kernel32" ( _
    ByVal hProcess As LongPtr, ByVal lpBaseAddress As LongPtr, _
    ByVal lpBuffer As LongPtr, ByVal nSize As LongPtr, _
    ByRef lpNumberOfBytesWritten As LongPtr) As Long
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)

' --- kernel32: System ---
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
    ByRef lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
    ByRef lpFrequency As LARGE_INTEGER) As Long
Public Declare PtrSafe Function GetComputerNameW Lib "kernel32" ( _
    ByVal lpBuffer As LongPtr, ByRef nSize As Long) As Long
Public Declare PtrSafe Function GetSystemDirectoryW Lib "kernel32" ( _
    ByVal lpBuffer As LongPtr, ByVal uSize As Long) As Long
Public Declare PtrSafe Function GetWindowsDirectoryW Lib "kernel32" ( _
    ByVal lpBuffer As LongPtr, ByVal uSize As Long) As Long
Public Declare PtrSafe Function GetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, ByVal lpBuffer As LongPtr, ByVal nSize As Long) As Long
Public Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long
Public Declare PtrSafe Function GlobalMemoryStatusEx Lib "kernel32" ( _
    ByRef lpBuffer As MEMORYSTATUSEX) As Long
Public Declare PtrSafe Function GetModuleHandleW Lib "kernel32" ( _
    ByVal lpModuleName As LongPtr) As LongPtr
Public Declare PtrSafe Function LoadLibraryW Lib "kernel32" ( _
    ByVal lpLibFileName As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Public Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Public Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

' --- kernel32: Pipe ---
Public Declare PtrSafe Function CreatePipe Lib "kernel32" ( _
    ByRef hReadPipe As LongPtr, ByRef hWritePipe As LongPtr, _
    ByRef lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Public Declare PtrSafe Function SetHandleInformation Lib "kernel32" ( _
    ByVal hObject As LongPtr, ByVal dwMask As Long, ByVal dwFlags As Long) As Long

' --- kernel32: Encoding ---
Public Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Public Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long

' --- kernel32: Global memory (clipboard) ---
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

' --- user32: Window ---
Public Declare PtrSafe Function FindWindowW Lib "user32" ( _
    ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
Public Declare PtrSafe Function FindWindowExW Lib "user32" ( _
    ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As LongPtr, ByVal lpszWindow As LongPtr) As LongPtr
Public Declare PtrSafe Function GetWindowTextW Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal nMaxCount As Long) As Long
Public Declare PtrSafe Function GetWindowTextLengthW Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function SetWindowTextW Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function MoveWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function EnableWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal bEnable As Long) As Long
Public Declare PtrSafe Function GetWindowRect Lib "user32" ( _
    ByVal hWnd As LongPtr, ByRef lpRect As RECT_API) As Long
Public Declare PtrSafe Function GetClientRect Lib "user32" ( _
    ByVal hWnd As LongPtr, ByRef lpRect As RECT_API) As Long
Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Public Declare PtrSafe Function SendMessageW Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal Msg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Public Declare PtrSafe Function PostMessageW Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal Msg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Public Declare PtrSafe Function GetClassNameW Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
Public Declare PtrSafe Function EnumWindows Lib "user32" ( _
    ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Public Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long

' --- user32: Clipboard ---
Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As LongPtr) As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function SetClipboardData Lib "user32" ( _
    ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Public Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

' --- advapi32: Registry ---
Public Declare PtrSafe Function RegOpenKeyExW Lib "advapi32" ( _
    ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As LongPtr) As Long
Public Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
Public Declare PtrSafe Function RegQueryValueExW Lib "advapi32" ( _
    ByVal hKey As LongPtr, ByVal lpValueName As LongPtr, ByVal lpReserved As LongPtr, _
    ByRef lpType As Long, ByVal lpData As LongPtr, ByRef lpcbData As Long) As Long
Public Declare PtrSafe Function RegSetValueExW Lib "advapi32" ( _
    ByVal hKey As LongPtr, ByVal lpValueName As LongPtr, ByVal Reserved As Long, _
    ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
Public Declare PtrSafe Function RegCreateKeyExW Lib "advapi32" ( _
    ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal Reserved As Long, _
    ByVal lpClass As LongPtr, ByVal dwOptions As Long, ByVal samDesired As Long, _
    ByVal lpSecurityAttributes As LongPtr, ByRef phkResult As LongPtr, _
    ByRef lpdwDisposition As Long) As Long
Public Declare PtrSafe Function RegDeleteKeyW Lib "advapi32" ( _
    ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr) As Long
Public Declare PtrSafe Function RegDeleteValueW Lib "advapi32" ( _
    ByVal hKey As LongPtr, ByVal lpValueName As LongPtr) As Long
Public Declare PtrSafe Function RegEnumKeyExW Lib "advapi32" ( _
    ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As LongPtr, _
    ByRef lpcchName As Long, ByVal lpReserved As LongPtr, ByVal lpClass As LongPtr, _
    ByVal lpcchClass As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
Public Declare PtrSafe Function RegEnumValueW Lib "advapi32" ( _
    ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpValueName As LongPtr, _
    ByRef lpcchValueName As Long, ByVal lpReserved As LongPtr, ByRef lpType As Long, _
    ByVal lpData As LongPtr, ByRef lpcbData As Long) As Long

' --- advapi32: Security ---
Public Declare PtrSafe Function GetUserNameW Lib "advapi32" ( _
    ByVal lpBuffer As LongPtr, ByRef pcbBuffer As Long) As Long

' --- shell32 ---
Public Declare PtrSafe Function ShellExecuteW Lib "shell32" ( _
    ByVal hWnd As LongPtr, ByVal lpOperation As LongPtr, _
    ByVal lpFile As LongPtr, ByVal lpParameters As LongPtr, _
    ByVal lpDirectory As LongPtr, ByVal nShowCmd As Long) As LongPtr
