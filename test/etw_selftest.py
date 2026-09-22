"""Verify the EvtCtx ETW silencer against real Word with a real trace session.

Usage: python test/etw_selftest.py

Controller: StartTraceW + EnableTraceEx2 (a proper manifest-provider enable; the
logman GUID path does not propagate enable state to unmanifested providers).

Two-phase flow: the Word probe registers its benign provider first (phase 1);
the harness then starts the trace session and enables the provider (enable
state is pushed to already-registered providers); phase 2 emits events in
three batches:

  pre     3 events   -> must appear in the ETL
  silent  7 events   -> emitted while EvtCtx is armed; must NOT appear
  control 2 events   -> after Disarm; must appear again

Expected provider events in the ETL: 5 (3 + 0 + 2), counted by provider GUID.
"""
import ctypes
import subprocess
import sys
import time
import uuid
from ctypes import wintypes
from pathlib import Path

try:
    import pythoncom
    import win32com.client
except ImportError:
    sys.stderr.write("[!] Required: pip install pywin32\n")
    sys.exit(1)

ROOT = Path(__file__).resolve().parent.parent
BAS = ROOT / "src" / "payloads" / "EtwSilence" / "EtwSilence_Standalone.bas"
DOC_PATH = ROOT / "test" / "etw_selftest.doc"
LOG_PATH = ROOT / "test" / "etw_selftest.log"
ETL_PATH = ROOT / "test" / "etw_selftest.etl"
CSV_PATH = ROOT / "test" / "etw_selftest.csv"

PROBE_GUID = "{9F3C05D2-4A77-4E88-8B6E-1D2F60C94A55}"
PROBE_TAIL = "60c94a55"   # order-invariant GUID tail; tracerpt renders data4 bytes both ways
SESSION = "EvSilenceSelfTest"

U64 = ctypes.c_ulonglong
advapi = ctypes.WinDLL("advapi32")

PROBE_MODULE = """Option Explicit

' TEST-ONLY probe: a benign custom provider emitted into a live trace session.
Private Declare PtrSafe Function EtwEventRegister Lib "ntdll" (ByVal ProviderId As LongPtr, ByVal EnableCallback As LongPtr, ByVal CallbackContext As LongPtr, ByRef RegHandle As LongPtr) As Long
Private Declare PtrSafe Function EtwEventUnregister Lib "ntdll" (ByVal RegHandle As LongPtr) As Long
Private Declare PtrSafe Function EtwEventWritePtr Lib "ntdll" Alias "EtwEventWrite" (ByVal RegHandle As LongPtr, ByVal EventDescriptor As LongPtr, ByVal UserDataCount As Long, ByVal UserData As LongPtr) As Long
Private Declare PtrSafe Function EtwEventWrite Lib "ntdll" (ByVal RegHandle As LongPtr, ByVal EventDescriptor As LongPtr, ByVal UserDataCount As Long, ByVal UserData As LongPtr) As Long
Private Declare PtrSafe Function EtwEventEnabled Lib "ntdll" (ByVal RegHandle As LongPtr, ByVal EventDescriptor As LongPtr) As Long
Private Declare PtrSafe Function StartTraceW Lib "advapi32" (ByRef hSession As LongPtr, ByVal name As LongPtr, ByVal props As LongPtr) As Long
Private Declare PtrSafe Function EnableTraceEx2 Lib "advapi32" (ByVal hSession As LongPtr, ByVal guid As LongPtr, ByVal control As Long, ByVal level As Byte, ByVal anyKw As LongPtr, ByVal allKw As LongPtr, ByVal timeout As Long, ByVal enableInfo As LongPtr) As Long
Private Declare PtrSafe Function ControlTraceByHandle Lib "advapi32" Alias "ControlTraceW" (ByVal hSession As LongPtr, ByVal name As LongPtr, ByVal props As LongPtr, ByVal control As Long) As Long
Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (ByVal d As LongPtr, ByVal s As LongPtr, ByVal l As LongPtr)

Private Const LOG_PATH As String = "{log}"

Private mReg As LongPtr
Private mDesc(0 To 15) As Byte

Private Sub L(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LOG_PATH For Append As #f
    Print #f, s & " err=" & Err.Number
    Close #f
End Sub

Public Sub Phase1()
    On Error Resume Next
    Dim g(0 To 15) As Byte
    g(0) = &HD2: g(1) = &H5: g(2) = &H3C: g(3) = &H9F
    g(4) = &H77: g(5) = &H4A: g(6) = &H88: g(7) = &H4E
    g(8) = &H6E: g(9) = &H8B: g(10) = &H2F: g(11) = &H1D
    g(12) = &H60: g(13) = &HC9: g(14) = &H4A: g(15) = &H55

    mDesc(0) = 1: mDesc(4) = 4    ' EVENT_DESCRIPTOR: Id=1, Level=4

    Dim hr As Long
    hr = EtwEventRegister(VarPtr(g(0)), 0, 0, mReg)
    L "p1_register_hr=" & hr & " h=" & Hex(mReg)
End Sub

Public Sub Phase2()
    On Error Resume Next
    Dim hr As Long, i As Long
    Dim en As Long

    ' ---- in-process trace controller (test-only) ----
    Const PROP_SIZE As Long = 120
    Const BUFSIZE As Long = 120 + 2 * 520
    Dim props(0 To BUFSIZE - 1) As Byte
    Dim hSess As LongPtr
    Dim g2(0 To 15) As Byte
    g2(0) = &HD2: g2(1) = &H5: g2(2) = &H3C: g2(3) = &H9F
    g2(4) = &H77: g2(5) = &H4A: g2(6) = &H88: g2(7) = &H4E
    g2(8) = &H6E: g2(9) = &H8B: g2(10) = &H2F: g2(11) = &H1D
    g2(12) = &H60: g2(13) = &HC9: g2(14) = &H4A: g2(15) = &H55

    Dim sessName As String: sessName = "EvInProc"
    Dim etlPath As String: etlPath = "C:\\Users\\itag7\\ll_projects\\EvasiveVBA\\test\\inproc.etl"
    Dim u32 As Long

    u32 = BUFSIZE: CopyMem VarPtr(props(0)), VarPtr(u32), 4          ' Wnode.BufferSize
    u32 = 1: CopyMem VarPtr(props(0)) + 64, VarPtr(u32), 4           ' LogFileMode = SEQ
    u32 = PROP_SIZE: CopyMem VarPtr(props(0)) + 116, VarPtr(u32), 4  ' LoggerNameOffset
    u32 = PROP_SIZE + 520: CopyMem VarPtr(props(0)) + 112, VarPtr(u32), 4 ' LogFileNameOffset
    CopyMem VarPtr(props(0)) + PROP_SIZE, StrPtr(sessName), Len(sessName) * 2 + 2
    CopyMem VarPtr(props(0)) + PROP_SIZE + 520, StrPtr(etlPath), Len(etlPath) * 2 + 2

    hr = StartTraceW(hSess, StrPtr(sessName), VarPtr(props(0)))
    L "c1_starttrace=" & hr & " h=" & Hex(hSess)
    hr = EnableTraceEx2(hSess, VarPtr(g2(0)), 1, 5, -1, 0, 0, 0)
    L "c2_enable=" & hr

    ' let the enable notification drain on the threadpool worker
    Dim t0 As Single: t0 = Timer
    Do While Timer < t0 + 1!
        DoEvents
    Loop

    en = EtwEventEnabled(mReg, VarPtr(mDesc(0))) And &HFF
    L "s02_enabled=" & en

    L "s01_start"
    For i = 1 To 3
        hr = EtwEventWritePtr(mReg, VarPtr(mDesc(0)), 0, 0)
    Next i
    L "s03_pre3_hr=" & hr

    Dim ok As Boolean
    ok = EvtCtx_Standalone.Arm()
    L "s04_arm=" & ok & " state=" & EvtCtx_Standalone.GetState() & " count=" & EvtCtx_Standalone.GetCount()

    For i = 1 To 7
        hr = EtwEventWritePtr(mReg, VarPtr(mDesc(0)), 0, 0)
    Next i
    L "s05_silent7_hr=" & hr

    Dim dok As Boolean
    dok = EvtCtx_Standalone.Disarm()
    L "s06_disarm=" & dok

    For i = 1 To 2
        hr = EtwEventWritePtr(mReg, VarPtr(mDesc(0)), 0, 0)
    Next i
    L "s07_ctrl2_hr=" & hr

    EtwEventUnregister mReg

    ' stop session, flush ETL
    u32 = BUFSIZE: CopyMem VarPtr(props(0)), VarPtr(u32), 4
    hr = ControlTraceByHandle(0, StrPtr(sessName), VarPtr(props(0)), 1)
    L "c3_stop=" & hr
    L "s08_done"
End Sub
"""


def stop_existing():
    PROP_SIZE = 120
    buf = ctypes.create_string_buffer(PROP_SIZE + 260 * 2)
    ctypes.memmove(ctypes.byref(buf), ctypes.byref(wintypes.ULONG(PROP_SIZE + 520)), 4)
    advapi.ControlTraceW(0, SESSION, buf, 1)  # EVENT_TRACE_CONTROL_STOP


def start_session() -> bool:
    PROP_SIZE = 120
    BUFSIZE = PROP_SIZE + 2 * 260 * 2
    buf = ctypes.create_string_buffer(BUFSIZE)

    def wr_u32(off, v):
        ctypes.memmove(ctypes.byref(buf, off), ctypes.byref(wintypes.ULONG(v)), 4)

    ctypes.memmove(ctypes.byref(buf), ctypes.byref(wintypes.ULONG(BUFSIZE)), 4)
    wr_u32(64, 1)                    # LogFileMode = EVENT_TRACE_FILE_MODE_SEQ
    wr_u32(116, PROP_SIZE)           # LoggerNameOffset
    logname_off = PROP_SIZE + 260 * 2
    wr_u32(112, logname_off)         # LogFileNameOffset
    ctypes.memmove(ctypes.byref(buf, PROP_SIZE), SESSION.encode("utf-16-le") + b"\x00\x00", len(SESSION) * 2 + 2)
    ctypes.memmove(ctypes.byref(buf, logname_off), str(ETL_PATH).encode("utf-16-le") + b"\x00\x00", len(str(ETL_PATH)) * 2 + 2)

    hSession = U64(0)
    advapi.StartTraceW.restype = ctypes.c_ulong
    advapi.StartTraceW.argtypes = [ctypes.POINTER(U64), wintypes.LPCWSTR, ctypes.c_void_p]
    st = advapi.StartTraceW(ctypes.byref(hSession), SESSION, buf)
    if st != 0:
        print(f"[!] StartTrace failed: {st} (trace sessions may need an elevated shell)")
        return False

    advapi.EnableTraceEx2.restype = ctypes.c_ulong
    advapi.EnableTraceEx2.argtypes = [U64, ctypes.c_void_p, wintypes.ULONG, ctypes.c_ubyte, U64, U64, wintypes.ULONG, wintypes.ULONG, ctypes.c_void_p]
    pguid = (wintypes.BYTE * 16).from_buffer_copy(uuid.UUID(PROBE_GUID).bytes)
    st = advapi.EnableTraceEx2(hSession, pguid, 1, 5, 0xFFFFFFFFFFFFFFFF, 0, 0, 0, None)
    if st != 0:
        print(f"[!] EnableTraceEx2 failed: {st}")
        return False
    return True


def stop_session_and_count():
    PROP_SIZE = 120
    buf = ctypes.create_string_buffer(PROP_SIZE + 260 * 2)
    ctypes.memmove(ctypes.byref(buf), ctypes.byref(wintypes.ULONG(PROP_SIZE + 520)), 4)
    advapi.ControlTraceW.restype = ctypes.c_ulong
    advapi.ControlTraceW.argtypes = [U64, wintypes.LPCWSTR, ctypes.c_void_p, wintypes.ULONG]
    advapi.ControlTraceW(0, SESSION, buf, 1)
    time.sleep(0.5)

    if CSV_PATH.exists():
        CSV_PATH.unlink()
    r = subprocess.run(["tracerpt", str(ETL_PATH), "-o", str(CSV_PATH), "-of", "csv", "-y"], capture_output=True, text=True)
    if not CSV_PATH.exists():
        print(f"[!] tracerpt failed: {r.stderr[:200]}")
        return None
    n = 0
    for line in CSV_PATH.read_text(encoding="utf-8", errors="replace").splitlines():
        if PROBE_TAIL in line:
            n += 1
    return n


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
    mod.Name = "EvtCtx_Standalone"
    mod.CodeModule.AddFromString(bas_code)

    probe = vb_proj.VBComponents.Add(1)
    probe.Name = "EtwProbe"
    probe.CodeModule.AddFromString(PROBE_MODULE.format(log=LOG_PATH))

    com_call_with_retry(lambda: doc.Save())
    time.sleep(0.5)
    com_call_with_retry(lambda: doc.Close(SaveChanges=0))
    print(f"[+] Built {DOC_PATH}")


def main():
    if ETL_PATH.exists():
        ETL_PATH.unlink()

    stop_existing()
    print("[*] Starting Word (phase 1: register provider)...")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    time.sleep(2)

    try:
        build_doc(word)
        word.AutomationSecurity = 1
        doc = com_call_with_retry(lambda: word.Documents.Open(str(DOC_PATH)))
        time.sleep(1)

        com_call_with_retry(lambda: word.Run("EtwProbe.Phase1"))
        time.sleep(1)

        com_call_with_retry(lambda: word.Run("EtwProbe.Phase2"))
        time.sleep(1)
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
        try:
            word.Quit()
        except Exception:
            pass

    INPROC = ROOT / "test" / "inproc.etl"
    count = None
    if INPROC.exists():
        csvp = INPROC.with_suffix(".csv")
        if csvp.exists():
            csvp.unlink()
        r = subprocess.run(["tracerpt", str(INPROC), "-o", str(csvp), "-of", "csv", "-y"], capture_output=True, text=True)
        if csvp.exists():
            count = sum(1 for line in csvp.read_text(encoding="utf-8", errors="replace").splitlines()
                        if PROBE_TAIL in line)
    else:
        print("[!] No in-process ETL produced")
    print()
    if not LOG_PATH.exists():
        print("[!] No probe log - macro did not run")
        return 1

    results = {}
    print("[*] Probe log:")
    for line in LOG_PATH.read_text().splitlines():
        print("    " + line)
        if "=" in line:
            k, v = line.split("=", 1)
            results[k] = v

    if count is None:
        return 1
    print(f"[*] Provider events in ETL: {count} (expected 5 = 3 pre + 0 silenced + 2 control)")

    try:
        cnt_silenced = int(results.get("s04_arm", "count=0").split("count=")[-1].split(" ")[0])
    except ValueError:
        cnt_silenced = 0
    checks = [
        ("provider registered", results.get("p1_register_hr", "").startswith("0 ")),
        ("pre-phase writes ok", results.get("s03_pre3_hr", "").startswith("0")),
        ("arm succeeded", results.get("s04_arm", "").startswith("True")),
        ("entries silenced", cnt_silenced >= 1),
        ("disarm ok", results.get("s06_disarm", "").startswith("True")),
        ("event count == 5 (silenced batch dropped)", count == 5),
    ]
    ok = True
    print("[*] Checks:")
    for name, passed in checks:
        print(f"    [{'+' if passed else '!'}] {name}")
        ok = ok and passed

    print()
    if ok:
        print("[+] PASS - events dropped while armed, flow restored after Disarm")
        return 0
    print("[!] FAIL")
    return 1


if __name__ == "__main__":
    sys.exit(main())
