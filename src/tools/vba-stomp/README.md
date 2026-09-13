# VBA Stomp

Strips VBA source code from `.doc` files while preserving compiled P-code.

## Install

```
pip install olefile
```

## Usage

```bash
# Stomp source code from all modules
python vba_stomp.py payload.doc

# Custom output path + hide modules from VBA editor
python vba_stomp.py payload.doc -o clean.doc --stomp-project

# Recon only -- list modules and offsets
python vba_stomp.py payload.doc --list
```

## What It Does

Every VBA module stream in an Office document stores two representations of the macro:

| Region | What | Where in stream |
|--------|------|-----------------|
| **PerformanceCache** (P-code) | Compiled bytecode | Top of stream |
| **CompressedSourceCode** | Textual VBA source | Bottom of stream (from `MODULEOFFSET`) |

When Word opens a document, it checks the `_VBA_PROJECT` header for the Office version that compiled the P-code. If it matches the local install, **Word executes P-code directly and ignores the source**. If it doesn't match, Word falls back to decompiling the source.

`vba-stomp` replaces the CompressedSourceCode region with null bytes. The P-code stays intact. Most AV engines inspect the VBA source text, not the P-code binary -- so they see nothing.

## The Drawback

**P-code is version-locked.** The compiled bytecode only executes on the exact Office version and bitness (32/64-bit) that created it. On any other version, Word falls back to the source code -- which is now nulled -- and **nothing runs**.

This means:

- You must know (or guess) the target's Office version before stomping
- A document stomped for Office 2019 x86 will silently fail on Office 2021 x64
- There is no graceful fallback -- it either works perfectly or fails completely
- Post-execution, Word decompiles P-code back into the editor, so the source reappears in memory (forensic recovery is possible after the fact)

The technique trades **broad compatibility for detection evasion**. It's a sniper round, not a shotgun.
