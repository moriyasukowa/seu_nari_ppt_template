"""Shared PowerPoint COM plumbing for the inspection helpers.

IMPORTANT — do not reintroduce a global taskkill here.

An earlier version ran `taskkill /F /IM POWERPNT.EXE` before each Open(). That
force-kills EVERY PowerPoint process on the machine, including windows the user
has open, and because a force-kill leaves no crash report or event-log entry it
looks to the user like PowerPoint "随机闪退". Windows logs showed no POWERPNT
crash events at all, which is how that was diagnosed.

The problem taskkill was working around is real but narrow: right after another
script calls Quit(), the process lingers in a shutting-down state, and a new
Dispatch() attaches to that dying instance so Presentations.Open() fails with a
generic 0x80070030 that looks like a corrupt file. That is handled here by
waiting for the process to settle and retrying — destructively killing is only
available through force_kill(), which callers must opt into explicitly.
"""
import os
import subprocess
import time

import win32com.client

OPEN_ATTEMPTS = 4
OPEN_BACKOFF_S = 1.5
SETTLE_TIMEOUT_S = 8.0


def _powerpnt_count():
    """How many POWERPNT.EXE processes exist, or -1 if it cannot be determined."""
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                             capture_output=True, text=True, timeout=15).stdout
    except Exception:  # noqa: BLE001
        return -1
    return sum(1 for line in out.splitlines() if "POWERPNT.EXE" in line.upper())


def force_kill():
    """Kill ALL PowerPoint processes. Closes the user's open windows too — only
    call this from an explicit user-facing escape hatch, never automatically."""
    subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(2.0)


def _wait_for_settle():
    """Wait until the PowerPoint process count stops changing.

    A mid-Quit instance makes the next Open() fail spuriously; letting it finish
    exiting fixes that without touching anyone's windows.
    """
    deadline = time.time() + SETTLE_TIMEOUT_S
    prev = _powerpnt_count()
    stable = 0
    while time.time() < deadline:
        time.sleep(0.6)
        now = _powerpnt_count()
        if now == prev:
            stable += 1
            if stable >= 2:
                return
        else:
            stable = 0
            prev = now


def connect():
    """Return a healthy PowerPoint Application, waiting out a dying instance."""
    _wait_for_settle()
    return win32com.client.Dispatch("PowerPoint.Application")


def open_pres(app, path):
    """Open read-only, windowless, retrying while a stale instance clears."""
    path = os.path.abspath(path)
    last = None
    for attempt in range(OPEN_ATTEMPTS):
        try:
            return app.Presentations.Open(path, ReadOnly=True, Untitled=False,
                                          WithWindow=False)
        except Exception as exc:  # noqa: BLE001 - retry the COM hiccup
            last = exc
            time.sleep(OPEN_BACKOFF_S * (attempt + 1))
    raise RuntimeError("could not open %s after %d attempts: %s"
                       % (path, OPEN_ATTEMPTS, last))


def shutdown(app):
    """Quit and wait for the process to actually go, so the next script is clean."""
    try:
        app.Quit()
    except Exception:  # noqa: BLE001
        pass
    deadline = time.time() + SETTLE_TIMEOUT_S
    while time.time() < deadline:
        if _powerpnt_count() <= 0:
            return
        time.sleep(0.5)
