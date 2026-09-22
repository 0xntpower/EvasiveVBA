# AMSI Neutralization (VBA)

Neutralizes the Antimalware Scan Interface inside the running Office process. After `Arm()`, the VBA engine's scan calls return immediately with the verdict the engine set before the call, which is clean. The scan function itself never runs. No memory in any loaded module is written and no page protections change.

Two variants, per repo convention:

- `AmsiNeutralize_Standalone.bas`: self-contained, paste into a standard module and go
- `AmsiNeutralize.bas`: same logic, reuses `WinAPI_Core.bas` for shared declares

Internal VBA module names are neutral (`ScanCtx` / `ScanCtx_Standalone`, same convention as `RefLoader_*`) because the module name is part of the source that gets scanned at runtime.

## Technique: pointer relocation (data-only)

VBE7.dll (the VBA runtime) does not import amsi.dll statically. At startup its AMSI bootstrap calls `LoadLibrary`/`GetProcAddress` and stores the resolved scan function in its own writable data:

```
VBE7 .data (current build):
    HAMSICONTEXT, hModule, pAmsiInitialize, pAmsiScanString, pAmsiUninitialize
```

Every runtime macro-content scan goes through that stored pointer:

```c
result = AMSI_RESULT_CLEAN;                 // pre-initialized to 0 (verified @ 0x11B8AE)
pAmsiScanString(ctx, bstr, contentName, 0, &result);   // HRESULT is ignored (verified)
if (result > 1) block;                      // verdict test (verified @ 0x11B654)
```

This module finds that pointer slot and repoints it at a `xor eax,eax; ret` (`33 C0 C3`) gadget that already exists in ntdll's executable section. The call returns immediately, `result` keeps its pre-initialized 0, and the engine proceeds. `Disarm` puts the original pointer back.

What this leaves in the process: amsi.dll stays byte-identical to its on-disk image, so entry-point memory scans and page-protection telemetry have nothing to find. VBE7's code is untouched; one qword in an already-RW data section now points into ntdll, which looks the same as any other resolved API pointer. The API surface is `GetModuleHandle`/`GetProcAddress`/`RtlMoveMemory` and nothing else. No `VirtualProtect`, no breakpoints, no exception handlers.

### How locations are found (no hardcoded offsets)

1. Resolve `amsi!AmsiScanString` by runtime-decoded names.
2. Parse VBE7's in-memory PE headers and enumerate writable sections.
3. Scan those sections for the 8-byte value of the resolved pointer. The slot locates itself by content, so Office updates that shift the .data layout do not matter.
4. Scan ntdll's executable section for `33 C0 C3` to find the gadget (entry at a byte offset is fine; x86 is variable-length and the entry point is ours to choose).
5. Overwrite the slot, then read it back to verify.

If the engine never bound the scanner (trusted document, or macro runtime scope off), `Arm()` reports `inactive` and succeeds. VBE7's own flag check already fails open in that state (verified @ 0x11B5B9), so there is nothing to neutralize.

## Why the module itself stays off AMSI's radar

VBA/AMSI integration is trigger-based: VBE7 keeps a circular log of COM/Win32 activity and submits the (deobfuscated, in-memory) macro source to the scan function when a high-risk event fires. This module's source will be scanned, so it is built to survive that.

No `amsi`, function-name, or patch-byte literals appear anywhere in the source. Sensitive strings are rebuilt at runtime from offset numeric arrays (`Ds()`), so the source only contains numbers. The declare set avoids the signatured clusters: no `VirtualProtect`, no `WriteProcessMemory`, no `CallWindowProcW` shellcode-runner pattern, no `AddVectoredExceptionHandler`. It uses module/proc resolution and `RtlMoveMemory`, the same primitives the repo's own `WinAPI_Core` ships for benign use. The write itself is one pointer-sized store into RW memory, which produces no protection changes and no kernel-emitted ETW-TI events.

Static (on-disk) detection of the carrier document is a separate problem. Pair with the repo's `vba-stomp` tool, which nulls the VBA source while keeping compiled p-code; the runtime source AMSI does see is what this module neutralizes.

## R&D record: technique selection

Candidates evaluated against 2025-era detection research and the current local build (amsi.dll Dec 2025 / Win11 26200, VBE7 16.0.20326):

| Approach | Writes | ETW-TI surface | Status |
|---|---|---|---|
| `AmsiScanBuffer` entry-point stub patch | amsi.dll .text | VirtualProtect-on-image + kernel memory-scan | Dead. Defender kills the process via kernel ETW-TI, which userland cannot bypass |
| je→jne single-byte flip on `AmsiOpenSession` | 1 byte .text | VirtualProtect-on-image | Works but per-build fragile; `AmsiOpenSession` was reshaped on the current build, and page-protect telemetry remains |
| Deep-offset patch (r-tec 2025 variant) | .text | VirtualProtect-on-image | Evades the current check; in the author's words, don't expect it to hold |
| HW breakpoint + VEH armed via `NtContinue` | none | none (per Praetorian 2025) | Attempted, then rejected for VBA; see below |
| Pointer relocation to existing gadget (this module) | 1 qword in RW data | none | Selected |

### The hardware-breakpoint experiment (negative result)

The HWBP+VEH technique (CCob 2022, Praetorian 2025 `NtContinue` arming) was implemented first and tested in Word on this machine.

Arming worked: VEH registration, `RtlCaptureContext` to Dr0/Dr7 via `NtContinue`, and a `GetThreadContext` readback confirming the breakpoint address. The `NtContinue` arming avoided `EtwTiLogSetContextThread` as documented.

The first scan interception crashed WINWORD (`0xC0000005` in KERNELBASE, in the string-marshalling neighborhood). Step-granular logging localized it: arming returned True, and the crash happened at the first `AmsiScanBuffer` invocation, the moment the kernel exception dispatcher called the VBA `AddressOf` callback.

VBA p-code functions are not viable exception-handler callbacks. `AddressOf` thunks hold up from message-pump contexts (EnumWindows, SetTimer) but not from kernel exception dispatch. That is a structural VBA limitation, independent of handler implementation, and the pointer-relocation design has no exception machinery at all.

### VBE7 ground truth (current build, disassembled locally)

- `AmsiInitialize` / `AmsiScanString` / `AmsiUninitialize` are resolved at runtime via GetProcAddress and stored as qwords in a tight `.data` cluster (no IAT entries).
- The scan wrapper pre-initializes the verdict local to 0 before every call and ignores the HRESULT. The verdict test is `*result > 1` and nothing else.
- If AMSI is not bound (init flag clear), the wrapper returns an error without scanning and the pre-initialized verdict stands. The unbound state is fail-open by design, which is what the `inactive` Arm() state relies on.
- The bootstrap runs once (flag-guarded), so a redirection persists for the process lifetime. A reopened document gets fresh state and must be re-armed.

## Limitations

- 64-bit Office only. The slot is 8 bytes and the gadget ABI is x64; a 32-bit port needs a 4-byte slot and an `xor eax,eax; ret 0x14` (stdcall stack-pop) gadget. Fails closed with a message on 32-bit.
- VBE7-scoped. Only the VBA engine's scans are neutralized, which is the whole macro-relevant path. Other consumers in the process keep a working AMSI.
- Re-arm on reload. Closing and reopening the document resets VBE7 state; the `ThisDocument` hooks (`Document_Open` → `Arm`, `Document_Close` → `Disarm`) cover it.
- Arming before the engine binds the scanner reports `inactive`. On the default macro runtime scope (low-trust documents) the bootstrap runs at VBA init, before any macro executes, so the slot is present by macro time in that case.
- The redirected pointer lives in RW data. An EDR that validates VBE7's stored AMSI pointer targets would notice it; nothing public does this today.

## Usage

```vba
ScanCtx_Standalone.Arm        ' True when redirected (or inactive-nothing-to-do)
ScanCtx_Standalone.GetState   ' redirected | inactive (...) | failed | idle
ScanCtx_Standalone.GetSlot(0) ' redirected slot address (diagnostics)
ScanCtx_Standalone.Run        ' convenience entry point with UI feedback
```

Arm early in the execution chain (e.g. `Document_Open`): scans that happen before arming see the (clean) source normally.

## Verification (dev box, 2026-09-22)

`python test/amsi_selftest.py` builds a Word doc via COM, forces `MacroRuntimeScanScope=2` (HKCU, restored afterwards) so the engine binds its scan pointer even in automation context, then runs the probe. Results:

```
s04_arm=True state=redirected slots=1
s05_slot=7FFC836ADD80 value=7FFD5B06C0F8     <- slot relocated to ntdll gadget
s06_gadget_bytes=33,C0,C3                     <- xor eax,eax; ret verified
s07_redirect_ok=True
s08_engine_hr=-2147024875 verdict=1           <- real engine responds (E_NOT_READY: RTP off on dev box)
s09_disarm=True
s10_restored=7FFD39B28100 match=True          <- original pointer restored byte-exact
s11_done                                      <- redirect live while the rest of the macro ran
```

All checks pass. The slot was located by value against a fresh ASLR base (no hardcoded offsets), repointed at the gadget, and restored to the original pointer after Disarm. The macro ran to completion with the redirect live and scanning scope forced to all documents, so any trigger scan that fired mid-macro went through the gadget.

Known-bad-content validation (a flagged string surviving a trigger scan) belongs in the detonation VM, not the dev box. Two environment notes for reproducing there: with the default scope, trusted or automation-opened documents never bind the scanner (Arm reports `inactive`, which is correct, since VBE7 skips AMSI for those by design), and with real-time protection off the engine answers `0x80070015` instead of `0x0` for direct control scans.

## References

- Microsoft, Office VBA + AMSI: Parting the veil on malicious macros (2018). VBE7 trigger/circular-log design; macro runtime scope.
- Outflank, Bypassing AMSI for VBA (2019). VBA-side AMSI mechanics, trust scope.
- S3cur3Th1sSh1t / r-tec, Bypass AMSI in 2025. Patch-class detection status, technique viability matrix.
- Praetorian, ETW Threat Intelligence and Hardware Breakpoints (2025). The `NtContinue` arming used in the rejected HWBP variant.
- CCob / Ethical Chaos, In-Process Patchless AMSI Bypass (2022). VEH pattern.
- Maldev Academy, AMSI Bypass Via Byte Patching. The je→jne baseline evaluated and rejected; the module notes were the starting point for this work.
- Local disassembly: amsi.dll (Dec 2025, Win11 26200) and VBE7.dll 16.0.20326. All structural claims above were verified there.
