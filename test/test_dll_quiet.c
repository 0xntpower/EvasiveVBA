#include <windows.h>

/* Quiet test DLL for the dropper self-test: on DLL_PROCESS_ATTACH writes a
 * marker file into %TEMP%. No UI (the original test_dll_x64.dll shows a modal
 * MessageBox, which hangs automation). */
BOOL WINAPI DllMain(HINSTANCE h, DWORD reason, LPVOID rsv) {
    if (reason == DLL_PROCESS_ATTACH) {
        char path[MAX_PATH];
        DWORD n = GetEnvironmentVariableA("TEMP", path, MAX_PATH);
        if (n > 0 && n < MAX_PATH - 40) {
            lstrcatA(path, "\\evba_dropper_ok.txt");
            HANDLE hf = CreateFileA(path, GENERIC_WRITE, 0, NULL,
                                    CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
            if (hf != INVALID_HANDLE_VALUE) {
                DWORD w;
                WriteFile(hf, "ok", 2, &w, NULL);
                CloseHandle(hf);
            }
        }
    }
    return TRUE;
}
