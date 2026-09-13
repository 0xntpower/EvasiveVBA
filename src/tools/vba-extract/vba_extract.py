"""VBA Extract -- dump VBA source strings from Office document module streams.

Reads CompressedSourceCode from each VBA module and prints the decompressed
text. Use after stomping to verify no readable source remains.
"""
from __future__ import annotations

import argparse
import struct
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Final

try:
    import olefile
except ImportError:
    sys.stderr.write("[!] Required dependency: pip install olefile\n")
    sys.exit(1)

# Reuse core routines from vba_stomp (sibling tool directory)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "vba-stomp"))
from vba_stomp import (  # noqa: E402
    DecompressionError,
    FileFormatError,
    ModuleInfo,
    StompError,
    _decompress_vba,
    _find_vba_root,
    _parse_dir_modules,
    _read_pcode_target,
)

# -- Constants ----------------------------------------------------------------

_VERSION: Final[str] = "1.0.0"
_OLE_MAGIC: Final[bytes] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

# -- Types --------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class ModuleSource:
    """Extracted source for one VBA module."""

    name: str
    source: str
    is_document: bool
    is_stomped: bool


# -- Terminal Helpers ---------------------------------------------------------

_USE_COLOR: bool = sys.stdout.isatty()


def _enable_win_ansi() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
    except (AttributeError, OSError):
        pass


def _c(text: str, code: str) -> str:
    return f"\033[{code}m{text}\033[0m" if _USE_COLOR else text


def _banner() -> None:
    print()
    print(f"  {_c('vba-extract', '1;96')} {_c(f'v{_VERSION}', '2')}")
    line = "-" * 36
    print(f"  {_c(line, '2')}")
    print()


# -- Source Decompression -----------------------------------------------------


def _decompress_source(raw_stream: bytes, offset: int) -> str:
    """Decompress VBA source from a module stream at the given offset.

    Args:
        raw_stream: Full module stream bytes.
        offset: Byte offset where CompressedSourceCode begins.

    Returns:
        Decompressed source text, or empty string if stomped/invalid.
    """
    if offset >= len(raw_stream):
        return ""

    source_region = raw_stream[offset:]

    if all(b == 0 for b in source_region):
        return ""

    try:
        raw_text = _decompress_vba(source_region)
    except DecompressionError:
        return ""

    try:
        text = raw_text.decode("utf-8", errors="ignore")
    except (UnicodeDecodeError, ValueError):
        text = raw_text.decode("latin-1", errors="ignore")

    return text.replace("\r\n", "\n").replace("\r", "\n")


# -- Core Engine --------------------------------------------------------------


def extract_sources(input_path: Path) -> list[ModuleSource]:
    """Extract VBA source text from all modules in a document.

    Args:
        input_path: Path to the .doc file.

    Returns:
        List of ModuleSource for each module found.

    Raises:
        FileFormatError: If the file isn't a valid OLE document with VBA.
    """
    if not input_path.exists():
        msg = f"File not found: {input_path}"
        raise FileFormatError(msg)

    with open(input_path, "rb") as f:
        magic = f.read(8)
    if magic != _OLE_MAGIC:
        msg = f"Not an OLE compound file: {input_path.name}"
        raise FileFormatError(msg)

    ole = olefile.OleFileIO(str(input_path))
    try:
        vba_root = _find_vba_root(ole)
        dir_path = f"{vba_root}/dir"
        compressed_dir = ole.openstream(dir_path).read()
        decompressed_dir = _decompress_vba(compressed_dir)
        modules = _parse_dir_modules(decompressed_dir, vba_root)

        target_info = _read_pcode_target(ole, vba_root, decompressed_dir)
        print(f"  {_c('[*]', '94')} VBA storage: {_c(vba_root, '96')}")
        print(f"  {_c('[*]', '94')} Target:      {_c(target_info.office_range, '92')}")
        print()

        results: list[ModuleSource] = []
        for mod in modules:
            if not ole.exists(mod.stream_path):
                results.append(ModuleSource(mod.name, "", mod.is_document, True))
                continue

            raw = ole.openstream(mod.stream_path).read()
            source = _decompress_source(raw, mod.source_offset)
            is_stomped = len(source.strip()) == 0

            results.append(ModuleSource(mod.name, source, mod.is_document, is_stomped))

        return results
    finally:
        ole.close()


# -- CLI ----------------------------------------------------------------------


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="vba-extract",
        description="Extract and print VBA source strings from Office documents.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "examples:\n"
            "  vba-extract payload.doc\n"
            "  vba-extract payload_stomped.doc --raw\n"
        ),
    )
    p.add_argument("input", type=Path, help="Target Office file (.doc, .xls, .ppt, etc.)")
    p.add_argument(
        "--raw", action="store_true",
        help="Print raw source without decoration (for piping)",
    )
    p.add_argument(
        "--hex", action="store_true",
        help="Also show hex dump of the source region (first 256 bytes)",
    )
    return p


def _hex_dump(data: bytes, max_bytes: int = 256) -> str:
    """Format bytes as a hex dump with ASCII sidebar."""
    lines: list[str] = []
    chunk = data[:max_bytes]
    for i in range(0, len(chunk), 16):
        row = chunk[i : i + 16]
        hex_part = " ".join(f"{b:02X}" for b in row)
        ascii_part = "".join(chr(b) if 32 <= b < 127 else "." for b in row)
        lines.append(f"      {i:04X}  {hex_part:<48}  {ascii_part}")
    if len(data) > max_bytes:
        lines.append(f"      ... ({len(data) - max_bytes} more bytes)")
    return "\n".join(lines)


def main() -> int:
    """CLI entry point."""
    _enable_win_ansi()
    _banner()

    parser = _build_parser()
    args = parser.parse_args()
    input_path: Path = args.input

    print(f"  {_c('[*]', '94')} File: {_c(str(input_path), '1')}")
    print()

    try:
        sources = extract_sources(input_path)
    except StompError as exc:
        print(f"  {_c('[-]', '91')} {exc}")
        return 1

    if not sources:
        print(f"  {_c('[!]', '93')} No VBA modules found")
        return 0

    stomped_count = sum(1 for s in sources if s.is_stomped)
    intact_count = len(sources) - stomped_count

    for mod in sources:
        mod_type = "document" if mod.is_document else "standard"
        if mod.is_stomped:
            status = _c("STOMPED", "92")
        else:
            status = _c("INTACT", "93")

        if args.raw:
            print(f"--- {mod.name} ({mod_type}) [{('STOMPED' if mod.is_stomped else 'INTACT')}] ---")
            if mod.source:
                print(mod.source)
            print()
        else:
            print(f"  {_c('=' * 60, '2')}")
            print(f"  {_c(mod.name, '1')}  ({mod_type})  {status}")
            print(f"  {_c('=' * 60, '2')}")

            if mod.is_stomped:
                print(f"  {_c('(no source -- region is null bytes)', '2')}")
            else:
                for line in mod.source.splitlines():
                    print(f"  {_c('|', '2')} {line}")

            if args.hex and mod.source:
                print()
                print(f"  {_c('Hex (source region):', '2')}")
                print(_hex_dump(mod.source.encode("latin-1", errors="replace")))

            print()

    # Summary
    print(f"  {_c('-' * 60, '2')}")
    print(
        f"  {_c('[*]', '94')} {len(sources)} module(s): "
        f"{_c(f'{stomped_count} stomped', '92')}, "
        f"{_c(f'{intact_count} intact', '93' if intact_count else '92')}"
    )

    if stomped_count == len(sources):
        print(f"  {_c('[+]', '92')} All source fully stomped")
    elif intact_count > 0:
        names = [s.name for s in sources if not s.is_stomped]
        print(f"  {_c('[!]', '93')} Readable source remains in: {', '.join(names)}")

    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
