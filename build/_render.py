"""Render .pptx files to PNG pages via PowerPoint COM (read-only inspection helper).

Usage:  python _render.py <in.pptx> <out_dir> [width] [height]

Never kills PowerPoint (see _com.py for why that mattered).
"""
import os
import sys

import _com


def render(pptx_path, out_dir, width=1920, height=1080):
    pptx_path = os.path.abspath(pptx_path)
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    app = _com.connect()
    pres = None
    try:
        pres = _com.open_pres(app, pptx_path)
        n = pres.Slides.Count
        for i in range(1, n + 1):
            pres.Slides(i).Export(os.path.join(out_dir, "%02d.png" % i), "PNG",
                                  width, height)
        return n
    finally:
        if pres is not None:
            pres.Close()
        _com.shutdown(app)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    print("%d slides" % render(sys.argv[1], sys.argv[2]))
