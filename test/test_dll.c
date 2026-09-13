#include <windows.h>

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved) {
    if (fdwReason == DLL_PROCESS_ATTACH) {
        MessageBoxA(NULL, "Reflective DLL loaded successfully!", "EvasiveVBA Test", MB_OK | MB_ICONINFORMATION);
    }
    return TRUE;
}
