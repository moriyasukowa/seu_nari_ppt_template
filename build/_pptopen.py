"""Open a pptx in PowerPoint and report whether it succeeds.

Use this to answer "is this package valid?" without rendering.

Usage:  python _pptopen.py <file.pptx> [...more.pptx]
        python _pptopen.py --force-kill <file.pptx>   # escape hatch, see below

Never kills PowerPoint automatically: a force-kill closes the user's windows and
leaves no crash trail, which reads as a mystery crash. See _com.py.
"""
import os
import sys

import _com


def probe(path):
    app = _com.connect()
    try:
        pr = _com.open_pres(app, path)
        n = pr.Slides.Count
        pr.Close()
        return True, "slides=%d" % n
    except Exception as exc:  # noqa: BLE001 - report the raw COM failure
        return False, str(exc)[:120]
    finally:
        _com.shutdown(app)


if __name__ == "__main__":
    args = sys.argv[1:]
    if args[:1] == ["--force-kill"]:
        _com.force_kill()
        args = args[1:]
    if not args:
        raise SystemExit(__doc__)
    rc = 0
    for p in args:
        ok, msg = probe(p)
        print("%-4s %-34s %s" % ("OK" if ok else "FAIL", os.path.basename(p), msg))
        rc |= 0 if ok else 1
    sys.exit(rc)
