"""
lineage_app.py
==============
Single entry point for a packaged "Data Lineage Builder" executable
(PyInstaller). Running the modules directly still works too:
    python lineage_ui.py          # GUI
    python data_lineage.py ...    # CLI

Dispatch:
  * no arguments (e.g. the .exe is double-clicked) -> launch the GUI, print a
    short loading line, and hide the console window once the window is up
    (unless the .exe was started from a real terminal, which is left alone);
  * any arguments -> run the command-line interface (data_lineage.main), so
    the same .exe works from a terminal:
        "Data Lineage Builder.exe" -i report.csv -o out.xlsx
        "Data Lineage Builder.exe" --gui
        "Data Lineage Builder.exe" --help
"""

import os
import sys


def _pid_to_name():
    """Map every running PID -> executable name (Windows, via Toolhelp)."""
    import ctypes
    from ctypes import wintypes

    k32 = ctypes.windll.kernel32

    class PROCESSENTRY32(ctypes.Structure):
        _fields_ = [('dwSize', wintypes.DWORD), ('cntUsage', wintypes.DWORD),
                    ('th32ProcessID', wintypes.DWORD),
                    ('th32DefaultHeapID', ctypes.c_void_p),
                    ('th32ModuleID', wintypes.DWORD), ('cntThreads', wintypes.DWORD),
                    ('th32ParentProcessID', wintypes.DWORD),
                    ('pcPriClassBase', ctypes.c_long), ('dwFlags', wintypes.DWORD),
                    ('szExeFile', ctypes.c_char * 260)]

    k32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    k32.Process32First.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32)]
    k32.Process32Next.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32)]
    k32.CloseHandle.argtypes = [wintypes.HANDLE]

    names = {}
    snap = k32.CreateToolhelp32Snapshot(0x00000002, 0)  # TH32CS_SNAPPROCESS
    if snap and snap != wintypes.HANDLE(-1).value:
        entry = PROCESSENTRY32()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
        ok = k32.Process32First(snap, ctypes.byref(entry))
        while ok:
            names[entry.th32ProcessID] = entry.szExeFile.decode('ascii', 'ignore')
            ok = k32.Process32Next(snap, ctypes.byref(entry))
        k32.CloseHandle(snap)
    return names


def _hide_own_console():
    """Hide the console window if it belongs to this app (e.g. the .exe was
    double-clicked). Never touch it when a real shell is sharing the console,
    so running from a terminal doesn't hide the user's window."""
    if sys.platform != 'win32':
        return
    try:
        import ctypes
        from ctypes import wintypes

        k32 = ctypes.windll.kernel32
        k32.GetConsoleWindow.restype = wintypes.HWND
        k32.GetConsoleProcessList.argtypes = [ctypes.POINTER(wintypes.DWORD), wintypes.DWORD]
        k32.GetConsoleProcessList.restype = wintypes.DWORD

        hwnd = k32.GetConsoleWindow()
        if not hwnd:
            return
        count = k32.GetConsoleProcessList((wintypes.DWORD * 1)(), 1)
        if count <= 0:
            return
        buf = (wintypes.DWORD * count)()
        got = k32.GetConsoleProcessList(buf, count)
        pids = set(buf[:got])

        names = _pid_to_name()
        shells = {'cmd.exe', 'powershell.exe', 'pwsh.exe', 'wt.exe',
                  'windowsterminal.exe', 'bash.exe', 'sh.exe', 'zsh.exe',
                  'fish.exe', 'conemu64.exe', 'conemuc64.exe', 'tcc.exe'}
        if any(names.get(p, '').lower() in shells for p in pids):
            return  # a real terminal is attached -> leave it alone

        ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception:
        pass


def _run_gui():
    print('Opening Data Lineage Builder - loading, please wait...', flush=True)
    import tkinter as tk
    import lineage_ui
    root = tk.Tk()
    lineage_ui.LineageApp(root)
    root.update_idletasks()
    root.update()           # draw the window before we hide the console
    _hide_own_console()
    root.mainloop()


def main():
    # MUST run first: in a frozen exe, the parallel render's worker processes
    # re-launch this exe; freeze_support intercepts them (else they'd each open
    # the GUI / re-run the build — a fork bomb).
    import multiprocessing
    multiprocessing.freeze_support()
    if len(sys.argv) > 1:
        import data_lineage
        data_lineage.main()
    else:
        _run_gui()


if __name__ == '__main__':
    main()
