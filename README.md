# EvasiveVBA

This repo contains a small toolkit I developed during research I did on utilizing Office and PDF documents for initial access.

Included in the kit:

1. **Python tools** that operate on Office documents at the binary level (stomp VBA source, extract and inspect macro strings)
2. **WinAPI helper modules** for VBA — prebuilt declares and wrappers so you spend less time on API signatures and more on the actual payload
3. **Standalone VBA scripts** implementing specific techniques (PPID spoofing, more to come)

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

## VBA scripts

Each technique lives in its own folder under `src/scripts/`. Every folder contains:

- `*_Standalone.bas` — self-contained, paste into a standard module and go
- `*.bas` — same logic, but uses WinAPI_Core.bas for shared types/declares
- `ThisDocument.txt` — the event hook to paste into ThisDocument
