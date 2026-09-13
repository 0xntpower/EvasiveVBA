"""VBA Stomp -- strip VBA source from Office documents, preserving compiled P-code.

Replaces CompressedSourceCode in VBA module streams with null bytes while
leaving PerformanceCache (P-code) intact. On a matching Office version, the
document executes from P-code alone -- source-level inspection sees nothing.
"""
from __future__ import annotations

import argparse
import logging
import shutil
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

logger = logging.getLogger(__name__)

# -- Constants ----------------------------------------------------------------

_VERSION: Final[str] = "1.0.0"
_VBA_COMPRESS_SIG: Final[int] = 0x01
_REC_MODULE_NAME: Final[int] = 0x0019
_REC_MODULE_STREAM: Final[int] = 0x001A
_REC_MODULE_OFFSET: Final[int] = 0x0031
_REC_MODULE_END: Final[int] = 0x002B
_REC_MODULES_HEADER: Final[int] = 0x000F
_REC_PROJECT_COOKIE: Final[int] = 0x0013
_STOMP_BYTE: Final[int] = 0x00
_OLE_MAGIC: Final[bytes] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

# -- Exceptions ---------------------------------------------------------------


class StompError(Exception):
    """Base for all VBA stomping errors."""


class FileFormatError(StompError):
    """Input is not a valid OLE document or lacks a VBA project."""


class DecompressionError(StompError):
    """VBA compressed stream is malformed."""


# -- Types --------------------------------------------------------------------


MODULE_PROCEDURAL: Final[int] = 0x0021
MODULE_DOCUMENT: Final[int] = 0x0022


@dataclass(frozen=True, slots=True)
class ModuleInfo:
    """Metadata for a VBA module within the project."""

    name: str
    stream_path: str
    source_offset: int
    is_document: bool = False


@dataclass(frozen=True, slots=True)
class StompResult:
    """Outcome of stomping one module."""

    module_name: str
    bytes_stomped: int
    stream_size: int


# -- P-code Version Identification --------------------------------------------

_SYSKIND_NAMES: Final[dict[int, str]] = {
    0: "16-bit Windows",
    1: "32-bit Windows",
    2: "Macintosh",
    3: "64-bit Windows",
}

_SYSKIND_SHORT: Final[dict[int, str]] = {
    0: "16-bit",
    1: "32-bit",
    2: "Mac",
    3: "64-bit",
}

_VERSION_MAP: Final[dict[int, tuple[str, str]]] = {
    0x000001CA: ("VBA 5/6", "Office 97-2003"),
    0x000001CB: ("VBA 6",   "Office 2000-2003"),
    0x000001CC: ("VBA 6",   "Office 2003-2007"),
    0x000001CD: ("VBA 7.0", "Office 2010"),
    0x000001CE: ("VBA 7.1", "Office 2013-365"),
}


@dataclass(frozen=True, slots=True)
class PCodeTarget:
    """P-code compatibility fingerprint."""

    platform: str
    vba_version: str
    office_range: str
    raw_header: str


def _parse_project_version(dir_data: bytes) -> tuple[int, int, int]:
    """Extract SysKind and VBA version from early dir stream records.

    Parses sequentially from the start of the decompressed dir stream
    through PROJECTVERSION (0x0009), which has a non-standard layout.

    Returns:
        (sys_kind, major_version, minor_version)
    """
    sys_kind = -1
    major = 0
    minor = 0
    pos = 0

    for _ in range(30):
        if pos + 6 > len(dir_data):
            break

        rec_id = struct.unpack_from("<H", dir_data, pos)[0]

        if rec_id == 0x0009:
            if pos + 12 <= len(dir_data):
                major = struct.unpack_from("<I", dir_data, pos + 6)[0]
                minor = struct.unpack_from("<H", dir_data, pos + 10)[0]
            break

        rec_size = struct.unpack_from("<I", dir_data, pos + 2)[0]

        if rec_id == 0x0001 and rec_size == 4 and pos + 10 <= len(dir_data):
            sys_kind = struct.unpack_from("<I", dir_data, pos + 6)[0]

        pos += 6 + rec_size
        if rec_size > 100_000:
            break

    return sys_kind, major, minor


def _read_pcode_target(
    ole: olefile.OleFileIO,  # type: ignore[no-any-unimported]
    vba_root: str,
    dir_data: bytes,
) -> PCodeTarget:
    """Build a P-code compatibility fingerprint from the VBA project.

    Combines SysKind (platform) from the dir stream with the _VBA_PROJECT
    header bytes and PROJECTVERSION to identify the target Office range.

    Args:
        ole: Open OLE file handle.
        vba_root: VBA storage path.
        dir_data: Decompressed dir stream.

    Returns:
        PCodeTarget with platform, VBA version, and Office range.
    """
    sys_kind, major, _minor = _parse_project_version(dir_data)

    platform = _SYSKIND_NAMES.get(sys_kind, f"Unknown (0x{sys_kind:02X})")
    bitness = _SYSKIND_SHORT.get(sys_kind, "?-bit")

    # Spec-defined values first, then infer from SysKind for modern Office
    if major in _VERSION_MAP:
        vba_ver, office = _VERSION_MAP[major]
    elif sys_kind == 3:
        vba_ver, office = "VBA 7.x", "Office 2010-365"
    elif sys_kind == 1 and major > 0x000001CE:
        vba_ver, office = "VBA 7.x", "Office 2010-365"
    elif sys_kind == 1:
        vba_ver, office = "VBA 6/7", "Office 2003-365"
    else:
        vba_ver, office = "Unknown", "Unknown"

    # Narrow the range using _VBA_PROJECT header byte 2 (build fingerprint)
    vba_proj_path = f"{vba_root}/_VBA_PROJECT"
    raw_hex = ""
    header = b""
    if ole.exists(vba_proj_path):
        header = ole.openstream(vba_proj_path).read()[:16]
        raw_hex = " ".join(f"{b:02X}" for b in header)

    if len(header) >= 3 and office.startswith("Office 2010-365"):
        ver_byte = header[2]
        if ver_byte >= 0xA0:
            office = "Office 2019-365"
        elif ver_byte >= 0x80:
            office = "Office 2016-365"
        elif ver_byte >= 0x60:
            office = "Office 2013-365"

    office_range = f"{office} ({bitness})"

    return PCodeTarget(
        platform=platform,
        vba_version=vba_ver,
        office_range=office_range,
        raw_header=raw_hex,
    )


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
    print(f"  {_c('vba-stomp', '1;96')} {_c(f'v{_VERSION}', '2')}")
    line = "-" * 36
    print(f"  {_c(line, '2')}")
    print()


def _info(msg: str) -> None:
    print(f"  {_c('[*]', '94')} {msg}")


def _ok(msg: str) -> None:
    print(f"  {_c('[+]', '92')} {msg}")


def _warn(msg: str) -> None:
    print(f"  {_c('[!]', '93')} {msg}")


def _err(msg: str) -> None:
    print(f"  {_c('[-]', '91')} {msg}")


def _detail(msg: str) -> None:
    print(f"      {msg}")


# -- VBA Decompression (MS-OVBA 2.4.1) ---------------------------------------


def _decompress_vba(compressed: bytes) -> bytes:
    """Decompress a VBA compressed container.

    Args:
        compressed: Raw bytes starting with signature 0x01.

    Returns:
        Decompressed byte content.

    Raises:
        DecompressionError: If the stream is malformed.
    """
    if not compressed or compressed[0] != _VBA_COMPRESS_SIG:
        sig = f"0x{compressed[0]:02X}" if compressed else "empty"
        msg = f"Bad VBA compression signature: {sig}"
        raise DecompressionError(msg)

    out = bytearray()
    pos = 1

    while pos < len(compressed):
        if pos + 1 >= len(compressed):
            break

        header = struct.unpack_from("<H", compressed, pos)[0]
        pos += 2

        chunk_data_len = (header & 0x0FFF) + 1
        is_compressed = bool(header & 0x8000)
        chunk_end = min(pos + chunk_data_len, len(compressed))
        chunk_start = len(out)

        if not is_compressed:
            out.extend(compressed[pos:chunk_end])
            pos = chunk_end
            continue

        while pos < chunk_end:
            if pos >= len(compressed):
                break
            flag = compressed[pos]
            pos += 1

            for bit in range(8):
                if pos >= chunk_end:
                    break

                if not (flag & (1 << bit)):
                    out.append(compressed[pos])
                    pos += 1
                else:
                    if pos + 1 >= len(compressed):
                        break
                    token = struct.unpack_from("<H", compressed, pos)[0]
                    pos += 2

                    diff = len(out) - chunk_start
                    bit_count = max((diff - 1).bit_length(), 4) if diff > 0 else 4
                    length_bits = 16 - bit_count
                    length = (token & ((1 << length_bits) - 1)) + 3
                    offset = (token >> length_bits) + 1

                    src = len(out) - offset
                    for _ in range(length):
                        out.append(out[src])
                        src += 1

    return bytes(out)


# -- Dir Stream Parsing -------------------------------------------------------


def _parse_dir_modules(data: bytes, vba_root: str) -> list[ModuleInfo]:
    """Extract module metadata from the decompressed dir stream.

    Locates the PROJECTMODULES section and parses each module's name,
    stream path, and CompressedSourceCode offset.

    Args:
        data: Decompressed dir stream bytes.
        vba_root: OLE path prefix for the VBA storage (e.g. "Macros/VBA").

    Returns:
        List of ModuleInfo for each module found.

    Raises:
        StompError: If the PROJECTMODULES header is missing.
    """
    marker = struct.pack("<HI", _REC_MODULES_HEADER, 2)
    idx = data.find(marker)
    if idx == -1:
        msg = "PROJECTMODULES section not found in dir stream"
        raise StompError(msg)

    module_count = struct.unpack_from("<H", data, idx + 6)[0]
    pos = idx + 8

    if pos + 5 < len(data) and struct.unpack_from("<H", data, pos)[0] == _REC_PROJECT_COOKIE:
        cookie_size = struct.unpack_from("<I", data, pos + 2)[0]
        pos += 6 + cookie_size

    modules: list[ModuleInfo] = []
    cur_name = ""
    cur_stream = ""
    cur_offset = -1
    cur_is_doc = False

    while pos < len(data) - 5 and len(modules) < module_count:
        rec_id = struct.unpack_from("<H", data, pos)[0]
        rec_size = struct.unpack_from("<I", data, pos + 2)[0]
        rec_data = data[pos + 6 : pos + 6 + rec_size]
        pos += 6 + rec_size

        if rec_id == _REC_MODULE_NAME:
            cur_name = rec_data.decode("ascii", errors="replace")
        elif rec_id == _REC_MODULE_STREAM:
            cur_stream = rec_data.decode("ascii", errors="replace")
        elif rec_id == MODULE_DOCUMENT:
            cur_is_doc = True
        elif rec_id == MODULE_PROCEDURAL:
            cur_is_doc = False
        elif rec_id == _REC_MODULE_OFFSET and rec_size == 4:
            cur_offset = struct.unpack_from("<I", rec_data, 0)[0]
        elif rec_id == _REC_MODULE_END:
            if cur_offset >= 0:
                stream = cur_stream or cur_name
                modules.append(
                    ModuleInfo(
                        name=cur_name or f"Module{len(modules)}",
                        stream_path=f"{vba_root}/{stream}",
                        source_offset=cur_offset,
                        is_document=cur_is_doc,
                    )
                )
            cur_name = ""
            cur_stream = ""
            cur_offset = -1
            cur_is_doc = False

    return modules


# -- VBA Root Discovery -------------------------------------------------------


def _find_vba_root(ole: olefile.OleFileIO) -> str:  # type: ignore[no-any-unimported]
    """Locate the VBA project storage within the OLE document.

    Args:
        ole: Open OLE file handle.

    Returns:
        OLE path to the VBA storage (e.g. "Macros/VBA").

    Raises:
        FileFormatError: If no VBA project is found.
    """
    for entry in ole.listdir(streams=True, storages=False):
        if entry[-1].lower() == "dir":
            parent = "/".join(entry[:-1])
            vba_proj = parent + "/_VBA_PROJECT"
            if ole.exists(vba_proj):
                return parent

    msg = "No VBA project found -- is this a macro-enabled document?"
    raise FileFormatError(msg)


# -- PROJECT Stream Stomping --------------------------------------------------


def _stomp_project_stream(
    ole: olefile.OleFileIO,  # type: ignore[no-any-unimported]
    vba_root: str,
    module_names: list[str],
) -> int:
    """Remove module references from the PROJECT text stream.

    Strips lines matching 'Module=<name>' so the VBA editor doesn't
    list the stomped modules.

    Args:
        ole: Open OLE file in write mode.
        vba_root: VBA storage path.
        module_names: Names of modules to hide.

    Returns:
        Number of references removed.
    """
    parts = vba_root.split("/")
    if len(parts) >= 2:
        project_path = "/".join(parts[:-1]) + "/PROJECT"
    else:
        project_path = "PROJECT"

    if not ole.exists(project_path):
        return 0

    raw = ole.openstream(project_path).read()
    text = raw.decode("latin-1")
    lines = text.splitlines(keepends=True)
    removed = 0
    kept: list[str] = []

    name_set = {n.lower() for n in module_names}
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("Module="):
            mod_name = stripped[7:]
            if mod_name.lower() in name_set:
                removed += 1
                continue
        kept.append(line)

    if removed > 0:
        new_data = "".join(kept).encode("latin-1")
        original_len = len(raw)
        if len(new_data) < original_len:
            new_data += b"\x00" * (original_len - len(new_data))
        ole.write_stream(project_path, new_data)

    return removed


# -- Core Engine --------------------------------------------------------------


def stomp_document(
    input_path: Path,
    output_path: Path,
    *,
    stomp_project: bool = False,
    stomp_events: bool = False,
    list_only: bool = False,
) -> list[StompResult]:
    """Stomp VBA source code in an Office document.

    Reads the VBA project, identifies all modules, and replaces each module's
    CompressedSourceCode region with null bytes while preserving P-code.
    Document modules (ThisDocument etc.) are skipped by default because
    stomping them breaks event handlers like Document_Open.

    Args:
        input_path: Path to the source .doc file.
        output_path: Where to write the stomped copy.
        stomp_project: Also clear module references from the PROJECT stream.
        list_only: Only parse and display -- don't modify anything.

    Returns:
        List of StompResult for each processed module.

    Raises:
        FileFormatError: If the file isn't a valid OLE document with VBA.
        StompError: If stomping fails.
    """
    if not input_path.exists():
        msg = f"File not found: {input_path}"
        raise FileFormatError(msg)

    with open(input_path, "rb") as f:
        magic = f.read(8)
    if magic != _OLE_MAGIC:
        msg = f"Not an OLE compound file: {input_path.name}"
        raise FileFormatError(msg)

    if not list_only:
        shutil.copy2(input_path, output_path)
        target = output_path
    else:
        target = input_path

    ole = olefile.OleFileIO(str(target), write_mode=not list_only)
    try:
        vba_root = _find_vba_root(ole)
        _info(f"VBA storage: {_c(vba_root, '96')}")

        dir_path = f"{vba_root}/dir"
        compressed_dir = ole.openstream(dir_path).read()
        decompressed_dir = _decompress_vba(compressed_dir)
        modules = _parse_dir_modules(decompressed_dir, vba_root)

        target_info = _read_pcode_target(ole, vba_root, decompressed_dir)
        print()
        _info("P-code target:")
        _detail(f"Platform:     {_c(target_info.platform, '1')}")
        _detail(f"VBA version:  {_c(target_info.vba_version, '1')}")
        _detail(f"Compatible:   {_c(target_info.office_range, '92')}")
        if target_info.raw_header:
            _detail(f"_VBA_PROJECT: {_c(target_info.raw_header, '2')}")
        _detail("")
        _detail(
            f"{_c('Note:', '93')} P-code runs only on matching version + bitness."
        )
        _detail("      Broad compat within same VBA major; exact match is safest.")

        if not modules:
            _warn("No VBA modules found in project")
            return []

        print()
        _info(f"Modules found: {_c(str(len(modules)), '1')}")

        col_w = max(len(m.name) for m in modules) + 2
        print()
        _detail(f"{'Module':<{col_w}} {'Type':<10} {'Offset':>10}   {'Stream Size':>12}")
        _detail(f"{'-' * col_w} {'-' * 10} {'-' * 10}   {'-' * 12}")

        for mod in modules:
            mod_type = _c("document", "93") if mod.is_document else _c("standard", "96")
            if ole.exists(mod.stream_path):
                raw = ole.openstream(mod.stream_path).read()
                _detail(
                    f"{mod.name:<{col_w}} {mod_type:<19} {f'0x{mod.source_offset:04X}':>10}"
                    f"   {f'{len(raw):,} B':>12}"
                )
            else:
                _detail(f"{mod.name:<{col_w}} {mod_type:<19} {'(missing)':>10}   {'N/A':>12}")
        print()

        if list_only:
            return []

        results: list[StompResult] = []

        _info("Stomping...")
        for mod in modules:
            if mod.is_document and not stomp_events:
                _warn(f"{mod.name:<{col_w}} skipped (document module, has event handlers)")
                continue

            if not ole.exists(mod.stream_path):
                _warn(f"  Stream missing: {mod.stream_path}")
                continue

            raw = ole.openstream(mod.stream_path).read()
            stream_len = len(raw)

            if mod.source_offset >= stream_len:
                _warn(
                    f"  {mod.name}: offset 0x{mod.source_offset:04X}"
                    f" beyond stream ({stream_len} B)"
                )
                continue

            source_len = stream_len - mod.source_offset
            stomped = bytearray(raw)
            stomped[mod.source_offset :] = bytes([_STOMP_BYTE] * source_len)

            ole.write_stream(mod.stream_path, bytes(stomped))

            pct = (source_len / stream_len * 100) if stream_len > 0 else 0
            _ok(f"{mod.name:<{col_w}} {source_len:>6,} bytes nulled ({pct:.1f}%)")
            results.append(StompResult(mod.name, source_len, stream_len))

        if stomp_project:
            print()
            removed = _stomp_project_stream(ole, vba_root, [m.name for m in modules])
            if removed:
                _info(f"PROJECT stream: {removed} module reference(s) cleared")
            else:
                _info("PROJECT stream: no references to clear")

        return results

    finally:
        ole.close()


# -- CLI ----------------------------------------------------------------------


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="vba-stomp",
        description="Strip VBA source code from Office documents, preserving P-code.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "examples:\n"
            "  vba-stomp payload.doc\n"
            "  vba-stomp payload.doc -o clean.doc --stomp-project\n"
            "  vba-stomp payload.doc --list\n"
        ),
    )
    p.add_argument("input", type=Path, help="Source Office file with VBA macros (.doc, .xls, .ppt, etc.)")
    p.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Output path (default: <input>_stomped.doc)",
    )
    p.add_argument(
        "--stomp-project", action="store_true",
        help="Also strip module refs from PROJECT stream",
    )
    p.add_argument(
        "--stomp-events", action="store_true",
        help="Also stomp document modules (ThisDocument etc.) -- breaks event triggers",
    )
    p.add_argument(
        "--list", action="store_true", dest="list_only",
        help="List modules without modifying",
    )
    p.add_argument("-v", "--verbose", action="store_true", help="Debug logging")
    return p


def main() -> int:
    """CLI entry point."""
    _enable_win_ansi()
    _banner()

    parser = _build_parser()
    args = parser.parse_args()

    if args.verbose:
        logging.basicConfig(level=logging.DEBUG, format="%(name)s: %(message)s")

    input_path: Path = args.input
    if args.output is not None:
        output_path: Path = args.output
    else:
        output_path = input_path.with_stem(input_path.stem + "_stomped")

    _info(f"Input:  {_c(str(input_path), '1')}")
    if not args.list_only:
        _info(f"Output: {_c(str(output_path), '1')}")
    print()

    try:
        results = stomp_document(
            input_path,
            output_path,
            stomp_project=args.stomp_project,
            stomp_events=args.stomp_events,
            list_only=args.list_only,
        )
    except StompError as exc:
        _err(str(exc))
        return 1

    if args.list_only:
        return 0

    print()
    total_stomped = sum(r.bytes_stomped for r in results)
    _ok(
        f"Done -> {_c(str(output_path), '1')}  "
        f"({len(results)} module(s), {total_stomped:,} bytes nulled)"
    )
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
