import os
from typing import List
import win32gui
import time
import psutil
import win32process
import pypresence

rpc = pypresence.Presence("1171889899679531129")
data = None

rpc.connect()


def find_active_processes(name: str) -> List[int]:
    processes = []
    for pid in psutil.pids():
        try:
            if psutil.Process(pid) and psutil.Process(pid).name() == name:
                processes.append(pid)
        except psutil.NoSuchProcess:
            pass
    return processes


LTSPICE_NAME = "LTSPICE.exe"

last_state = {}
last_program = {"pid": None, "path": None}


def handle_ltspice(title: str) -> None:
    global last_state
    filename = title.replace("LTspice XVII - ", "")
    file, ext = os.path.splitext(filename)

    # The window title is of the format "LTspice XVII - [filename.extension]"
    # if the view is maximized
    if ext and ext[-1] == "]":
        ext = ext[:-1]
        file = file[1:]
        # filename = filename[1:-1]

    state = {}
    if ext in (".asc",):  # SCHEMATICS
        state = {
            "large_image": "schematic",
            "large_text": "Schematic",
            "details": "Editing a schematic",
            "state": filename,
        }
    elif ext in (
        ".cir",
        ".net",
        ".sp",
    ):  # NETLISTS
        state = {
            "large_image": "symbol",  # TODO: Add images for these
            "large_text": "Symbol",
            "details": "Editing a netlist",
            "state": filename,
        }
    elif ext in (
        ".raw",
        ".fra",
    ):  # WAVEFORMS
        state = {
            "large_image": "graph",
            "large_text": "Waveform",
            "details": "Viewing a graph",
            "state": filename,
        }
    elif ext in (
        ".bjt",
        ".cap",
        ".dio",
        ".ind",
        ".jft",
        ".mos",
        ".res",
    ):  # DISCRETES
        state = {
            "large_image": "symbol",  # TODO: Add images for each one of these
            "large_text": "Symbol",
            "details": "Editing a discrete component",
            "state": filename,
        }
    elif ext in (".asy",):  # SYMBOLS
        state = {
            "large_image": "symbol",
            "large_text": "Symbol",
            "details": "Editing a symbol",
            "state": filename,
        }
    else:
        rpc.clear()
        state = {
            "large_image": "logo",
            "large_text": "LTspice",
            "details": "Idling",
        }

    if state != last_state:
        print("UPDATED!")
        print(filename, file, ext)
        rpc.update(
            **state,
            start=int(time.time()),
            small_image="logo",
            small_text="LTspice",
        )
        last_state = state
        return True
    return False


def main() -> None:
    global last_program
    WINDOW_HANDLERS = {
        LTSPICE_NAME: handle_ltspice,
    }
    while True:
        time.sleep(3)
        foreground_hwnd = win32gui.GetForegroundWindow()
        pid = win32process.GetWindowThreadProcessId(foreground_hwnd)
        if pid[-1] >= 0:
            try:
                active_process = psutil.Process(pid[-1])
                if active_process.name() in WINDOW_HANDLERS.keys():
                    handler = WINDOW_HANDLERS[active_process.name()]
                    if handler(win32gui.GetWindowText(foreground_hwnd)):
                        last_program = {
                            "pid": active_process.pid,
                            "path": active_process.name(),
                        }
            except psutil.NoSuchProcess:
                pass
        if last_program["pid"] is not None and last_program[
            "pid"
        ] not in find_active_processes(last_program["path"]):
            last_program = {"pid": None, "path": None}
            print("Clear RPC - program closed")
            rpc.clear()


if __name__ == "__main__":
    main()
