# ETW Silence (VBA)

Kills every user-mode ETW event the Office process emits, process-wide, by clearing
the cached enable state on each provider registration ntdll keeps in-process. After
`Arm()`, `EtwEventWrite`, `EtwEventWriteEx`, `EtwWriteTransfer`, and TraceLogging
writes all return success without writing or submitting anything. `Disarm` restores
the original bytes.

Two variants, per repo convention:

- `EtwSilence_Standalone.bas`: self-contained, paste into a standard module and go
- `EtwSilence.bas`: same logic, reuses `WinAPI_Core.bas` for shared declares
- Internal module names are neutral (`EvtCtx` / `EvtCtx_Standalone`), same convention
  as `ScanCtx`.

This is the second stage of the chain documented in this repo:
`AmsiNeutralize` (content scans) -> `EtwSilence` (user-mode telemetry) -> one of the
`ReflectiveLoader_*` modules (fetch and map the payload). `ThisDocument.txt` wires
the full chain.

## How it works

Manifest-based ETW events all funnel through the provider registration entries
ntdll keeps in per-process slot tables. The registration handle a provider receives
encodes a bucket, an index into that bucket's slot array, and a 16-bit token stored
at `RegEntry+0x54`. Both write paths are gated by cached enable bytes on the entry:

- `RegEntry+0x74`: session-path enable. Zero means the `NtTraceEvent` submission
  path is skipped (verified @ `EtwEventWrite` 0x18006C031 on ntdll 26100.9444).
- `RegEntry+0xEC`: private trace buffer enable. Zero means the private-buffer path
  is skipped (verified @ 0x18006BFCE region).

With both zero, `EtwEventWrite` falls through to its common exit and returns
success. No event is built, no syscall is made.

The module never stores an ntdll offset in its source. Discovery is anchored on
code and validated by roundtrip:

1. Register a probe provider (fixed research GUID, never written to) and decode
   its handle into bucket/index/token.
2. Read the first 0x120 bytes of `EtwEventWrite` (resolved by a runtime-decoded
   name) and scan for the instructions that reference the slot-table and the
   per-bucket counts arrays. Both encodings are covered: the imagebase-relative
   form this build uses (`mov r11,[rbx+rdx*8+rva]`, `cmp eax,[rbx+rdx*4+rva]`)
   and the RIP-relative form older builds use (`4C 8B 1D disp32`, `3B 05 disp32`).
3. A candidate table or counts address is accepted only if the probe's own entry
   validates through it (entry found at bucket/index, even pointer, token word
   matches) and the counts value for the probe bucket bounds the probe index.
   All validation reads are constrained to ntdll's mapped image.
4. The probe provider is unregistered. Zeroed state: nothing remains.
5. Blind: walk all 8 buckets, and for every live entry save and clear the two
   enable bytes. `Disarm` writes the saved bytes back.

Nothing executable is written; the only writes are single bytes into ntdll heap
registration entries that are already RW. The two ntdll registration APIs are
declared by name (`EtwEventRegister`, `EtwEventUnregister`) because VBA cannot
marshal their 4th out-parameter through a function pointer; they are ordinary
telemetry APIs, not evasion indicators.

## Why this exists: the Defender for Endpoint verdict

The starting brief for this module was "neutralize Microsoft Defender for
Endpoint." Research and local component analysis say literal neutralization from an
unprivileged Office macro is not a real option, and trying is itself the alert:

- `MsMpEng.exe` and the MDE sensor (`Sense.exe`) are protected-process-light. A
  user-mode handle with write access cannot be opened, so in-process patching
  (the Maldev-style approach in the notes that started this) is not reachable
  from the macro even before detection.
- Service, driver, and registry tampering (stop WinDefend, unload WdFilter,
  `DisableAntiSpyware`, exclusion writes) is blocked and reported by Tamper
  Protection, which is kernel-enforced and on by default. Public tooling that
  does any of this is signatured.
- The sensors that actually see this chain after AMSI is handled are kernel-side:
  ETW Threat Intelligence (alloc/protect exec memory, remote memory ops, thread
  context changes), process/thread/image notify callbacks, and the minifilter.
  None of these can be blinded from user mode; that requires a driver, which is a
  different project and its own detection story.
- The Maldev notes that seeded this task hook `AmsiScanBuffer` and `NtTraceEvent`
  with hardware breakpoints plus a vectored exception handler. The AMSI half is
  covered better by `AmsiNeutralize`, and the VEH half was tested during that
  module's development: a VBA p-code callback invoked from kernel exception
  dispatch crashes the process. The same notes also record that Microsoft
  Defender did not use user-mode ETW for their scenario.

What user-mode ETW silence is worth against MDE specifically is therefore narrow:
it is insurance, and it covers the in-process providers that do exist. It kills
the `Microsoft-Antimalware-Scan-Interface` provider events for any scan path not
already redirected, the .NET runtime provider events if the chain ends up
hosting the CLR in-process, and the general diagnostic providers Office runs.
Against third-party EDRs that do consume user-mode ETW from Office processes, it
is load-bearing. The kernel-side surfaces it cannot touch are listed above and
belong to the payload and the mapper, not this module.

The loudest kernel-visible artifact this chain still produces is the reflective
loader's single `VirtualAlloc(..., PAGE_EXECUTE_READWRITE)` for the mapped image,
which emits an ETW-TI alloc-exec event on every modern build. Splitting it into
an RW allocation plus one RX transition does not hide it, but it does stop
advertising RWX. That change belongs in the loader modules, not here.

## Limitations

- 64-bit Office only. The handle layout, slot pointers, and gate offsets verified
  here are x64.
- Registrations created after `Arm()` get fresh enable state from the kernel and
  are not silenced. Arm late (right before the loader), or re-Arm, if the chain
  loads components that register providers of their own (the CLR does).
- The gate offsets (+0x74, +0xEC) and the token offset (+0x54) are ntdll-build
  facts, verified on 26100.9444. The token roundtrip validates the entry's
  identity, not these offsets; a future ntdll that moves the enable cluster would
  make the writes ineffective rather than corrupting unrelated memory (the fields
  around both offsets are the rest of the enable state), but the functional test
  should be rerun per build before operational use.
- Silencing is process-wide and indiscriminate: Office diagnostic telemetry stops
  too. A defender correlating "process alive, zero events across every provider"
  could in principle notice; nothing public does this today.
- One event per provider may already be in flight on another thread while the
  byte is cleared; the write itself is atomic.

## Usage

```vba
EvtCtx_Standalone.Arm        ' True when silenced (or inactive-nothing-found)
EvtCtx_Standalone.GetState   ' silenced | inactive (...) | failed | idle
EvtCtx_Standalone.GetCount   ' registrations silenced
EvtCtx_Standalone.GetEntry(0)' entry address (diagnostics)
EvtCtx_Standalone.Disarm     ' restores original enable bytes
EvtCtx_Standalone.Run        ' convenience entry point with UI feedback
```

Run it after `ScanCtx.Arm` in the document-open chain (see `ThisDocument.txt`).

## Verification (dev box, 2026-09-23)

`python test/etw_selftest.py` builds a Word doc, registers a probe provider from
VBA, starts a trace session and enables it from inside the same process, and
emits three batches through a live ETL: 3 events, 7 events while armed, 2 events
after Disarm. Confirmed results:

```
c1_starttrace=0 c2_enable=0      <- session + enable from VBA
s02_enabled=1                    <- enable state reached the provider gates
s04_arm=True state=silenced count=359
s05_silent7_hr=0                 <- 7 events emitted while armed
c3_stop=0                        <- ETL flushed
Provider events in ETL: 5        <- 3 pre + 0 silenced + 2 control
```

Exactly the pre and control batches landed; every armed-phase event was
dropped with a success return, and Disarm restored the flow. The 359 count is
the whole-process registration population, so the silencing is process-wide,
not probe-specific.

Two harness findings:

- Cross-process enable from this dev shell never propagated into WINWORD (same
  controller into a python child works both orderings). The test therefore
  controls the session from inside Word, which also matches the in-process
  reality an operator cares about. The kernel-side delivery is
  `NtTraceControl(27)`-registered event, `EtwpNotificationThread` threadpool
  wait, `NtTraceControl(16)` pull, `EtwpUpdateEnableInfoAndCallback` writes
  the gates; verified in ntdll 26100.9444 disassembly.
- `logman -p {guid}` does not propagate enable state to unmanifested providers
  at all (verified against a known-good provider). Use
  `StartTrace` + `EnableTraceEx2` when tracing this module by hand.

## References

- Local ground truth: ntdll 26100.9444 `EtwEventWrite` / `EtwEventRegister`
  disassembled during this session; slot-table and counts-array references,
  handle encoding, and both gate offsets verified there.
- Maldev Academy, *Evading Microsoft Defender Via Patching* (the HBP/VEH approach
  this module replaces for VBA; also the source of the "Defender does not use
  user-mode ETW" observation for its scenario).
- Praetorian, *ETW Threat Intelligence and Hardware Breakpoints* (2025), kernel
  sensor surface and `NtContinue`.
- r-tec / S3cur3Th1sSh1t, *Bypass AMSI in 2025*, detection stance on
  user-mode telemetry patches.
- geoffchappell.com, ETW registration structure notes (kernel-side layout
  background).
- Microsoft, Tamper Protection and PPL documentation for the neutralization
  verdict above.
