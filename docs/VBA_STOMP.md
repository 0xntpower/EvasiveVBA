# VBA Stomp

Strips VBA source code from `.doc` files while preserving compiled P-code.

## Install

```
pip install -r requirements.txt
```

## Usage

```bash
# Stomp source code from all modules
python vba_stomp.py verylegit.doc

# Custom output path + hide modules from VBA editor
python vba_stomp.py verylegit.doc -o clean.doc --stomp-project

# Recon only -- list modules and offsets
python vba_stomp.py verylegit.doc --list
```

## What It Does

Every VBA module stream in an Office document stores two representations of the macro:

| Region | What | Where in stream |
|--------|------|-----------------|
| **PerformanceCache** (P-code) | Compiled bytecode | Top of stream |
| **CompressedSourceCode** | Textual VBA source | Bottom of stream (from `MODULEOFFSET`) |

When Word opens a document, it checks the `_VBA_PROJECT` header for the Office version that compiled the P-code. If it matches the local install, **Word executes P-code directly and ignores the source**. If it doesn't match, Word falls back to the textual VBA source.

`vba-stomp` replaces the CompressedSourceCode region with null bytes. The P-code stays intact. Most AV engines inspect the VBA source text, not the P-code binary - so **they see nothing**.

In Fact - after uploading your doc to sites like VirusTotal (don't do that with real operational payloads), comparing the stompted and none stompt versions of your doc you'll
see detection rates fall from 80-90% to 10-20%.

## The Drawback

**P-code is version-locked.** The compiled bytecode only executes on the exact Office version and bitness (32/64-bit) that created it. On any other version, Word falls back to the source code which is now nulled.

This means you would require knowing what version of Office your target is running to be able
to tailore your offensive document to their environment.