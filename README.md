# EvasiveVBA

This repo contains a small toolkit I developed during research I did on utilizing Office and PDF documents for initial access.

Included in the kit:

1. **Python tools** that operate on Office documents at the binary level (stomp VBA source, extract and inspect macro strings)
2. **WinAPI helper modules** for VBA — prebuilt declares and wrappers so you spend less time on API signatures and more on the actual payload
3. **Standalone VBA payloads** implementing specific techniques (PPID spoofing, AMSI neutralization, ETW silencing, a full-chain dropper, more to come)

## Quick start

```bash
pip install -r requirements.txt
```

**Stomp a document** (nulls VBA source, preserves compiled P-code):
```bash
python src/tools/vba-stomp/vba_stomp.py payload.doc -o clean.doc
```

More on VBA source stomping for evading detection [VBA-STOMP.md](src/tools/vba-stomp/README.md).

**Extract VBA source from doc**:
```bash
python src/tools/vba-extract/vba_extract.py clean.doc
```

## WinAPI modules

Import `WinAPI_Core.bas` and `WinAPI.bas` into any VBA project (Word, Excel, Access). Covers file I/O, registry, process management, window manipulation, clipboard, memory operations, and high-resolution timing. 64-bit safe.

See the module headers for the full function list.

## VBA payloads

Each technique lives in its own folder under `src/payloads/`. Every folder contains:

- `*_Standalone.bas` — self-contained, paste into a standard module and go
- `*.bas` — same logic, but uses WinAPI_Core.bas for shared types/declares
- `ThisDocument.txt` — the event hook to paste into ThisDocument

| Payload | What it does |
|---|---|
| `AmsiNeutralize` | Neutralizes runtime content scans by relocating the engine's stored scan pointer to a return-success gadget in ntdll. |
| `EtwSilence` | Silences user-mode ETW by clearing the cached enable bytes on every provider registration in the process. |
| `Dropper` | The full chain in one module: both neutralizers, then WinHTTP fetch, manual mapping, and a threadless DllMain call. |
| `ReflectiveLoader_WinHTTP` | Fetches a DLL over HTTP(S) and manually maps it in-process; no file touches disk. |
| `ReflectiveLoader_DNS` | Same mapper, with the payload retrieved through DNS TXT record queries. |
| `ReflectiveLoader_ICMP` | Same mapper, with the payload carried inside ICMP echo packets. |
| `PPIDSpoof` | Spawns a process under a chosen legitimate parent process. |

---

This is meant for authorized pentesting only.
