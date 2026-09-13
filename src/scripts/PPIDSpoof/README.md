# PPID Spoofing (VBA)

Creates a process with a fake parent PID. The spawned process appears as a child of a chosen legitimate process (default: `SecurityHealthSystray.exe`) instead of Word.

## How it works

1. Find the target parent process by name via WMI
2. `OpenProcess` with `PROCESS_CREATE_PROCESS`
3. Build a `PROC_THREAD_ATTRIBUTE_LIST` with `PROC_THREAD_ATTRIBUTE_PARENT_PROCESS` pointing to that handle
4. `CreateProcessW` with `EXTENDED_STARTUPINFO_PRESENT | CREATE_SUSPENDED` — the new process inherits the spoofed parent
5. `ResumeThread` to start execution

The spawned `cmd.exe` shows up in Process Explorer / Task Manager under `SecurityHealthSystray.exe`, not under `WINWORD.EXE`.

## Config

Edit the constants in the `Run()` sub:

```vba
Const PARENT_NAME As String = "SecurityHealthSystray.exe"
cmdLine = "C:\Windows\System32\cmd.exe /K echo PPID Spoofed from VBA! && pause"
```

Change `PARENT_NAME` to any running process you want as the apparent parent. Change `cmdLine` to whatever you want to launch.

## Note

PPID spoofing changes only the parent PID metadata. It does not affect the security token — the spawned process still runs under Word's token with the current user's privileges. EDR products that log process creation events by token rather than parent PID will still attribute the child to Word.
