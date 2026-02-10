"""
32-bit COM surrogate server for the PlateLoc ActiveX control.

This script is designed to be run by a 32-bit Python interpreter.
It hosts the PlateLoc ActiveX control inside a hidden AtlAxWin
window (required for visual ActiveX controls) and exposes a simple
JSON-over-stdin/stdout protocol so the main (64-bit) process can
send commands and receive results.

The main thread runs a Windows message pump (required by the
ActiveX control), while a background thread reads requests from
stdin and posts them to the main thread via a thread-safe queue.

Protocol (one JSON object per line):
  Request:  {"cmd": "method_name", "args": [...]}
  Response: {"ok": true, "result": ...}  or  {"ok": false, "error": "..."}

This file is an internal implementation detail and should not be
imported directly.
"""

import ctypes
import ctypes.wintypes
import json
import queue
import sys
import threading
import time
import traceback


def main():
    import pythoncom
    import win32api
    import win32com.client
    import win32con
    import win32gui

    # Initialise COM as STA — required for ActiveX controls
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

    # --- Load ATL and register the AtlAxWin window class ---------------
    atl = None
    for lib_name in ("atl", "atl110", "atl100", "atl90", "atl80", "atl71"):
        try:
            atl = ctypes.windll.LoadLibrary(lib_name)
            atl.AtlAxWinInit()
            break
        except (OSError, AttributeError):
            atl = None

    if atl is None:
        _fatal("Could not load ATL library — AtlAxWin hosting unavailable")

    # --- Create a hidden host window -----------------------------------
    wc = win32gui.WNDCLASS()
    wc.lpszClassName = "PlateLocSurrogateHost"
    wc.lpfnWndProc = {
        win32con.WM_DESTROY: lambda hwnd, msg, wp, lp: win32gui.PostQuitMessage(0),
    }
    try:
        win32gui.RegisterClass(wc)
    except Exception:
        pass  # class may already be registered from a previous run

    hwnd_parent = win32gui.CreateWindow(
        "PlateLocSurrogateHost", "PlateLoc COM Host",
        0, 0, 0, 100, 100,
        0, 0, win32api.GetModuleHandle(None), None,
    )

    # Shared state
    plateloc = None
    hwnd_ax = None
    progid = "PLATELOC.PlateLocCtrl.2"

    # Request queue: background stdin reader → main thread
    req_queue: queue.Queue = queue.Queue()
    shutdown = threading.Event()

    # ------------------------------------------------------------------
    # Request handler (runs on the main/STA thread)
    # ------------------------------------------------------------------
    def handle(request: dict) -> dict:
        nonlocal plateloc, hwnd_ax, progid
        cmd = request.get("cmd", "")
        args = request.get("args", [])

        try:
            # -- lifecycle ---
            if cmd == "create":
                if args:
                    progid = args[0]
                hwnd_ax = win32gui.CreateWindow(
                    "AtlAxWin", progid,
                    win32con.WS_CHILD,
                    0, 0, 100, 100,
                    hwnd_parent, 0,
                    win32api.GetModuleHandle(None), None,
                )
                # Retrieve the hosted IDispatch interface
                punk = ctypes.c_void_p()
                hr = atl.AtlAxGetControl(hwnd_ax, ctypes.byref(punk))
                if hr != 0 or not punk.value:
                    return {"ok": False, "error": f"AtlAxGetControl failed (hr={hr})"}
                idisp = pythoncom.ObjectFromAddress(
                    punk.value, pythoncom.IID_IDispatch,
                )
                plateloc = win32com.client.Dispatch(idisp)
                return {"ok": True, "result": f"Created {progid}"}

            if plateloc is None:
                return {"ok": False, "error": "COM object not created yet"}

            # -- properties ---
            if cmd == "set_blocking":
                plateloc.Blocking = bool(args[0])
                return {"ok": True, "result": None}

            # -- init / close ---
            if cmd == "initialize":
                profile = args[0] if args else "default"
                res = plateloc.Initialize(profile)
                return {"ok": True, "result": res}

            if cmd == "close":
                res = plateloc.Close()
                return {"ok": True, "result": res}

            # -- profiles ---
            if cmd == "enumerate_profiles":
                profiles = plateloc.EnumerateProfiles()
                return {"ok": True, "result": list(profiles) if profiles else []}

            # -- temperature ---
            if cmd == "get_actual_temperature":
                return {"ok": True, "result": plateloc.GetActualTemperature()}

            if cmd == "get_sealing_temperature":
                return {"ok": True, "result": plateloc.GetSealingTemperature()}

            if cmd == "set_sealing_temperature":
                return {"ok": True, "result": plateloc.SetSealingTemperature(int(args[0]))}

            # -- sealing time ---
            if cmd == "get_sealing_time":
                return {"ok": True, "result": plateloc.GetSealingTime()}

            if cmd == "set_sealing_time":
                return {"ok": True, "result": plateloc.SetSealingTime(float(args[0]))}

            # -- cycle control ---
            if cmd == "start_cycle":
                return {"ok": True, "result": plateloc.StartCycle()}

            if cmd == "stop_cycle":
                return {"ok": True, "result": plateloc.StopCycle()}

            if cmd == "apply_seal":
                return {"ok": True, "result": plateloc.ApplySeal()}

            # -- stage ---
            if cmd == "move_stage_in":
                return {"ok": True, "result": plateloc.MoveStageIn()}

            if cmd == "move_stage_out":
                return {"ok": True, "result": plateloc.MoveStageOut()}

            # -- error handling ---
            if cmd == "abort":
                return {"ok": True, "result": plateloc.Abort()}

            if cmd == "retry":
                return {"ok": True, "result": plateloc.Retry()}

            if cmd == "ignore":
                return {"ok": True, "result": plateloc.Ignore()}

            # -- info ---
            if cmd == "get_firmware_version":
                return {"ok": True, "result": plateloc.GetFirmwareVersion()}

            if cmd == "get_version":
                return {"ok": True, "result": plateloc.GetVersion()}

            if cmd == "get_last_error":
                return {"ok": True, "result": plateloc.GetLastError()}

            if cmd == "get_cycle_count":
                return {"ok": True, "result": plateloc.GetCycleCount()}

            # -- UI ---
            if cmd == "show_diags_dialog":
                modal = bool(args[0]) if args else True
                sec = int(args[1]) if len(args) > 1 else 0
                return {"ok": True, "result": plateloc.ShowDiagsDialog(modal, sec)}

            # -- housekeeping ---
            if cmd == "ping":
                return {"ok": True, "result": "pong"}

            if cmd == "quit":
                return {"ok": True, "result": "bye"}

            return {"ok": False, "error": f"Unknown command: {cmd}"}

        except Exception as exc:
            return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}

    # ------------------------------------------------------------------
    # Background stdin reader
    # ------------------------------------------------------------------
    def stdin_reader():
        try:
            for line in sys.stdin:
                line = line.strip()
                if not line:
                    continue
                try:
                    request = json.loads(line)
                except json.JSONDecodeError as exc:
                    request = {"_parse_error": str(exc)}
                req_queue.put(request)
                if request.get("cmd") == "quit":
                    break
        except Exception:
            pass
        finally:
            shutdown.set()
            # Unblock the message loop
            ctypes.windll.user32.PostQuitMessage(0)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def write_response(resp: dict):
        sys.stdout.write(json.dumps(resp) + "\n")
        sys.stdout.flush()

    # Signal ready
    write_response({"ok": True, "result": "ready"})

    # Start stdin reader thread
    reader = threading.Thread(target=stdin_reader, daemon=True)
    reader.start()

    # ------------------------------------------------------------------
    # Main message loop
    # ------------------------------------------------------------------
    QS_ALLINPUT = 0x04FF
    PM_REMOVE = 0x0001
    msg = ctypes.wintypes.MSG()

    while not shutdown.is_set():
        # Drain request queue
        while True:
            try:
                request = req_queue.get_nowait()
            except queue.Empty:
                break

            if "_parse_error" in request:
                write_response({"ok": False, "error": f"Invalid JSON: {request['_parse_error']}"})
                continue

            response = handle(request)
            write_response(response)

            if request.get("cmd") == "quit":
                shutdown.set()
                break

        if shutdown.is_set():
            break

        # Wait for messages or a 100 ms timeout (to re-check the queue)
        ctypes.windll.user32.MsgWaitForMultipleObjects(
            0, None, False, 100, QS_ALLINPUT,
        )

        # Pump pending Windows messages
        while ctypes.windll.user32.PeekMessageW(
            ctypes.byref(msg), None, 0, 0, PM_REMOVE,
        ):
            if msg.message == 0x0012:  # WM_QUIT
                shutdown.set()
                break
            ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
            ctypes.windll.user32.DispatchMessageW(ctypes.byref(msg))

    # ------------------------------------------------------------------
    # Cleanup
    # ------------------------------------------------------------------
    if plateloc:
        try:
            plateloc.Close()
        except Exception:
            pass

    if hwnd_ax:
        try:
            win32gui.DestroyWindow(hwnd_ax)
        except Exception:
            pass
    try:
        win32gui.DestroyWindow(hwnd_parent)
    except Exception:
        pass

    pythoncom.CoUninitialize()


def _fatal(msg: str):
    """Write an error and exit."""
    sys.stdout.write(json.dumps({"ok": False, "error": msg}) + "\n")
    sys.stdout.flush()
    sys.exit(1)


if __name__ == "__main__":
    main()
