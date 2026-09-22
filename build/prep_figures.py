"""Collect the figures a deck needs into assets/figures/ under stable ASCII keys.

Why this exists: source illustration folders tend to have names that are awkward
to reference from code — non-ASCII filenames, doubled extensions like
"foo.png.jpeg", mixed case, duplicates. This script copies (or rasterises) every
image into one flat directory with a predictable key and writes a manifest of
pixel sizes so slides.py can fit each image to the canvas without distortion.

Point SOURCE_DIR at your own image folder. Everything is discovered
automatically; RENAMES is optional and only exists so you can give a few keys
readable names instead of ones derived from the filename.

    python prep_figures.py

Optional: the template's demo deck does not need this to run. If SOURCE_DIR does
not exist the script exits with a clear message rather than silently doing
nothing.
"""
import json
import os
import re
import shutil
import subprocess
from collections import OrderedDict

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))

# ---------------------------------------------------------------------------
# 改这里：你的插图目录。位图直接复制，矢量 PDF 会被光栅化。
# ---------------------------------------------------------------------------
SOURCE_DIR = os.path.join(ROOT, "你的插图目录")     # ← 指向你的图目录
DEST = os.path.join(ROOT, "assets", "figures")
# TeX Live 的 bin 目录：优先环境变量 TEXBIN，其次 PATH 上的 pdflatex，
# 最后回退到本机安装位置。回退为空串时 pdflatex.exe 由 PATH 解析。
TEXBIN = (os.environ.get("TEXBIN")
          or os.path.dirname(shutil.which("pdflatex") or "")
          or r"D:\texlive\2024\bin\windows")
DPI = 400

# 可选：给个别文件指定更好读的 key；未列出的按文件名自动生成。
RENAMES = {
    # "韧性图.pdf": "resilience",
    # "路网图NC.png": "nc_network",
}

BITMAP_EXT = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp")
VECTOR_EXT = (".pdf",)


def _run(cmd, cwd):
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError("%s failed:\n%s\n%s"
                           % (os.path.basename(cmd[0]), p.stdout[-1200:], p.stderr[-1200:]))


def key_for(name, taken):
    """A stable ASCII key for a filename.

    Strips a doubled extension ("foo.png.jpeg" -> "foo"), folds everything
    non-alphanumeric to "_", and falls back to figNN when nothing survives
    (e.g. an all-CJK filename).
    """
    stem = name
    for ext in sorted(BITMAP_EXT + VECTOR_EXT, key=len, reverse=True):
        if stem.lower().endswith(ext):
            stem = stem[: -len(ext)]
            break
    key = re.sub(r"[^0-9A-Za-z]+", "_", stem).strip("_").lower() or "fig"
    base, i = key, 2
    while key in taken:
        key = "%s_%d" % (base, i)
        i += 1
    return key


def rasterise(src_name, key):
    src = os.path.join(SOURCE_DIR, src_name)
    # ⚠️ 输出必须是相对名 + cwd=DEST：pdftocairo 会把含非 ASCII 的**绝对输出
    # 路径**过 ANSI 代码页弄坏（"Error opening output file"），输入侧不受影响。
    # 详见 make_formulas.py 里同样的注释。
    # -singlefile: without it pdftocairo appends "-1" to the page name
    _run([os.path.join(TEXBIN, "pdftocairo.exe"), "-png", "-singlefile",
          "-r", str(DPI), os.path.abspath(src), key], DEST)
    png = os.path.join(DEST, key + ".png")
    if not os.path.exists(png):
        raise RuntimeError("no PNG for %s" % key)
    return png


def copy_bitmap(src_name, key):
    ext = os.path.splitext(src_name)[1].lower()
    dest = os.path.join(DEST, key + ext)
    shutil.copyfile(os.path.join(SOURCE_DIR, src_name), dest)
    return dest


def discover():
    """[(filename, key)] for every supported image under SOURCE_DIR."""
    if not os.path.isdir(SOURCE_DIR):
        raise SystemExit(
            "SOURCE_DIR not found: %s\n"
            "  改 prep_figures.py 顶部的 SOURCE_DIR 指向你的插图目录，或跳过这一步\n"
            "  （示例 deck 不需要它）。" % SOURCE_DIR)
    taken, out = set(), []
    for name in sorted(os.listdir(SOURCE_DIR)):
        ext = os.path.splitext(name)[1].lower()
        if ext not in BITMAP_EXT + VECTOR_EXT or os.path.isdir(os.path.join(SOURCE_DIR, name)):
            continue
        key = RENAMES.get(name) or key_for(name, taken)
        taken.add(key)
        out.append((name, key))
    if not out:
        raise SystemExit("no images found in %s" % SOURCE_DIR)
    return out


def main():
    os.makedirs(DEST, exist_ok=True)
    manifest = OrderedDict()
    for src_name, key in discover():
        if os.path.splitext(src_name)[1].lower() in VECTOR_EXT:
            path = rasterise(src_name, key)
            extra = {"dpi": DPI}
            note = "(from pdf @%d dpi)" % DPI
        else:
            path = copy_bitmap(src_name, key)
            extra = {}
            note = ""
        with Image.open(path) as im:
            w, h = im.size
        manifest[key] = {"file": "assets/figures/" + os.path.basename(path),
                         "w": w, "h": h, "source": src_name}
        manifest[key].update(extra)
        print("%-20s %5dx%-6d %s %s" % (key, w, h, src_name, note))

    out = os.path.join(DEST, "manifest.json")
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
    print("wrote %s (%d figures)" % (out, len(manifest)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
