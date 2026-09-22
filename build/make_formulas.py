"""Render LaTeX formulas / pseudocode to transparent PNGs for the deck.

PowerPoint's own equation editor output looks poor next to the rest of the
design, so the pipeline is: LaTeX -> PDF -> transparent PNG -> placed as a
picture on the slide. Equation NUMBERS stay as slide text so they remain
editable and searchable.

  latex (standalone, tight border)  ->  formula.pdf
  pdftocairo -png -transp -r <dpi>  ->  transparent PNG
  Pillow alpha bbox                 ->  trim transparent margin

Two optional sources feed this, and either can be empty:
  * FORMULAS / ALGOS - written inline below. Edit these for your own deck.
  * _math_spec.json  - blocks pulled verbatim out of a LaTeX source by
                       extract_math.py. Optional: skipped if that file is absent.

Run before build_pptx.py:
    python make_formulas.py

Writes assets/formulas/<key>.png plus manifest.json (pixel sizes, so slides.py
can size each image without guessing an aspect ratio).
"""
import json
import os
import shutil
import subprocess
import sys

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
OUTDIR = os.path.join(ROOT, "assets", "formulas")
WORK = os.path.join(HERE, "_formula_work")
SPEC = os.path.join(HERE, "_math_spec.json")
# TeX Live 的 bin 目录：优先环境变量 TEXBIN，其次 PATH 上的 pdflatex，
# 最后回退到本机安装位置。回退为空串时 pdflatex.exe 由 PATH 解析。
TEXBIN = (os.environ.get("TEXBIN")
          or os.path.dirname(shutil.which("pdflatex") or "")
          or r"D:\texlive\2024\bin\windows")
DPI = 400
BASE_PT = 10          # the standalone class default, so slides.py can scale consistently

# Ink colour must match design/tokens.json -> color.ink
INK = "1A1A1A"

# (key, latex, tag). The tag is drawn as slide text beside the image, never
# baked in -- that is what keeps equation numbers editable.
FORMULAS = [
    ("demo_objective",
     r"\min_{\theta\in\Theta}\ \mathcal{L}(\theta)"
     r"=\frac{1}{N}\sum_{i=1}^{N}\ell\big(y_i,f_{\theta}(x_i)\big)"
     r"+\lambda\lVert\theta\rVert_2^{2}", "(1)"),
    ("demo_constraint",
     r"\text{s.t.}\quad g_j(\theta)\le 0,\ j=1,\dots,m;\qquad "
     r"h_k(\theta)=0,\ k=1,\dots,p", "(2)"),
    ("demo_update",
     r"\theta^{(t+1)}=\theta^{(t)}-\eta_t\,\nabla\mathcal{L}\big(\theta^{(t)}\big)",
     "(3)"),
    ("demo_gap",
     r"\mathrm{gap}(t)=\frac{\big\lvert \mathcal{L}(\theta^{(t)})-\underline{\mathcal{L}}"
     r"\big\rvert}{\max\big(1,\lvert\underline{\mathcal{L}}\rvert\big)}\le\varepsilon",
     "(4)"),
    ("demo_subproblem",
     r"\hat{\theta}=\arg\min_{\theta}\ \mathcal{L}(\theta)"
     r"+\frac{\rho}{2}\big\lVert\theta-\bar{\theta}\big\rVert_2^{2}", "(5)"),
]

# (key, caption, body) rendered as a ruled algorithm float — three-line style with
# an automatic "Algorithm N: <caption>" header. Body is algorithm2e syntax.
# 中文关键字可直接写（ctex 已加载）。
ALGOS = [
    ("demo_pseudo", "带收敛判据与步长回退的梯度下降", r"""\KwIn{观测集 $\{(x_i,y_i)\}_{i=1}^{N}$，初始步长 $\eta_0$，精度 $\varepsilon$}
\KwOut{参数估计 $\theta^{(t)}$}
初始化 $\theta^{(0)}$，令 $t \leftarrow 0$\;
\While{$\mathrm{gap}(t) > \varepsilon$}{
  计算当前梯度 $\nabla\mathcal{L}(\theta^{(t)})$\;
  由式 (3) 更新 $\theta^{(t+1)} \leftarrow \theta^{(t)} - \eta_t\,\nabla\mathcal{L}(\theta^{(t)})$\;
  \If{$\mathcal{L}(\theta^{(t+1)}) > \mathcal{L}(\theta^{(t)})$}{
    步长减半 $\eta_t \leftarrow \eta_t / 2$\;
    回退到 $\theta^{(t)}$ 并重新更新\;
  }
  由式 (4) 重新计算 $\mathrm{gap}(t)$\;
  $t \leftarrow t + 1$\;
}
\Return $\theta^{(t)}$\;"""),
]


# ctex (with the Windows font set) is required because a handful of blocks carry
# Chinese inside math AND the algorithm listings are written in Chinese.
#
# ← 公式字体在这里改。默认 `lmodern` = Latin Modern（LaTeX 默认数学字体的
#   OpenType 版）。若想回到与 pptx 正文同族的 Palatino 数学，把下面这行换成
#   `\usepackage{mathpazo}` 即可；两者只影响公式图片，不影响别的。
_BODY = r"""\usepackage{amsmath,amssymb}
\usepackage{xcolor}
\usepackage[fontset=windows]{ctex}
\usepackage{lmodern}
\usepackage[ruled,vlined,linesnumbered]{algorithm2e}
\begin{document}
\color[HTML]{@@INK@@}%
@@BODY@@
\end{document}
"""

# `standalone` cannot host a display environment at top level (align/gather fail
# with "Missing \endgroup"), which is why the extractor rewrites them to
# aligned/gathered inside a display-math span. varwidth gives the algorithm
# listings a text width to wrap against.
TEX_MATH = ("\\documentclass[border=2pt]{standalone}\n" + _BODY)
# 算法用自然宽度：绝大多数清单应该以横构图放进版面，靠宽度定尺寸，
# 字号才够大。只有极长的清单（见 slides.ALG_MIN_K）才考虑切两段。
TEX_ALGO = ("\\documentclass[border=4pt,varwidth=16cm]{standalone}\n" + _BODY)


def _run(cmd, cwd):
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError("%s failed:\n%s\n%s" % (cmd[0], p.stdout[-1500:], p.stderr[-1500:]))
    return p


def render(key, body, kind="math", caption=None):
    """Render one block. kind="algo" wraps `body` in a ruled `algorithm` float,
    which is what gives the three-line (booktabs-style) look with an automatic
    "Algorithm N: <caption>" header."""
    os.makedirs(WORK, exist_ok=True)
    if kind == "algo":
        inner = "\\begin{algorithm}[H]\n\\caption{%s}\n%s\n\\end{algorithm}" % (
            caption or key, body)
        doc = TEX_ALGO.replace("@@BODY@@", inner)
    else:
        doc = TEX_MATH.replace("@@BODY@@", "$\\displaystyle %s$" % body)
    tex = os.path.join(WORK, key + ".tex")
    with open(tex, "w", encoding="utf-8") as fh:
        # .replace(), not %-formatting: the LaTeX body contains % and {} freely
        fh.write(doc.replace("@@INK@@", INK))
    _run([os.path.join(TEXBIN, "pdflatex.exe"), "-interaction=nonstopmode",
          "-halt-on-error", key + ".tex"], WORK)
    pdf = os.path.join(WORK, key + ".pdf")
    if not os.path.exists(pdf):
        raise RuntimeError("no PDF produced for %s" % key)
    # ⚠️ 输出路径必须是**相对名**。pdftocairo 把 argv 过一遍 ANSI 代码页，
    # 含非 ASCII 的**绝对输出路径**会被弄坏，报 "Error opening output file" ——
    # 而同样含中文的输入路径没事，相对输出名 + cwd 也没事。本项目目录是中文，
    # 所以这不是假想问题。（实测：绝对+ASCII 成功；绝对+中文 失败；相对+中文cwd 成功。）
    # -singlefile: without it pdftocairo appends "-1" to the page name
    _run([os.path.join(TEXBIN, "pdftocairo.exe"), "-png", "-transp",
          "-singlefile", "-r", str(DPI), pdf, key], WORK)
    png = os.path.join(WORK, key + ".png")
    if not os.path.exists(png):
        raise RuntimeError("no PNG produced for %s" % key)

    im = Image.open(png).convert("RGBA")
    bbox = im.getbbox()          # trims the fully-transparent margin
    if bbox:
        im = im.crop(bbox)
    dest = os.path.join(OUTDIR, key + ".png")
    im.save(dest)
    return dest, im


def _blank_cut(im):
    """Row index near the middle with the most transparency — a clean place to
    cut a tall image in two without slicing through a glyph."""
    import numpy as _np
    a = _np.asarray(im)
    alpha = a[..., 3]
    h, w = alpha.shape
    lo, hi = int(h * 0.35), int(h * 0.65)
    if hi <= lo:
        return h // 2
    density = (alpha[lo:hi] > 8).sum(axis=1)
    return lo + int(density.argmin())


def emit_halves(dest, key, im, manifest, overlap=18):
    """Also emit a top/bottom half pair for an algorithm listing.

    Only a LAST RESORT: a long listing placed as one image is bounded by the body
    height, so it shrinks; cutting it at a blank row and putting the halves side
    by side recovers the width. Most listings should stay whole -- slides.alg()
    only reaches for the pair when the single placement would be genuinely too
    small (ALG_MIN_K).
    """
    w, h = im.size
    cut = _blank_cut(im)
    halves = [("a", im.crop((0, 0, w, min(h, cut + overlap)))),
              ("b", im.crop((0, max(0, cut - overlap), w, h)))]
    for suffix, part in halves:
        name = "%s_%s" % (key, suffix)
        path = os.path.join(OUTDIR, name + ".png")
        part.save(path)
        manifest[name] = {"file": "assets/formulas/%s.png" % name,
                          "w": part.size[0], "h": part.size[1],
                          "dpi": DPI, "kind": "algo", "split_of": key}
    print("%-10s %5dx%-6d split -> %s_a / %s_b" % ("", w, h, key, key))


def spec_entries():
    """[(key, body, kind, tag)] from the extracted source blocks, or [] if none.

    Optional by design: the template ships with its own FORMULAS / ALGOS, so a
    missing _math_spec.json (i.e. extract_math.py never run) is not an error.
    """
    if not os.path.exists(SPEC):
        print("note: %s not found - rendering only the inline FORMULAS / ALGOS"
              % os.path.basename(SPEC))
        return []
    with open(SPEC, encoding="utf-8") as fh:
        spec = json.load(fh)
    out = []
    for info in spec.values():
        for b in info["blocks"]:
            out.append((b["id"], b["body"], b["kind"], None, b.get("caption")))
    return out


def main():
    os.makedirs(OUTDIR, exist_ok=True)
    jobs = ([(k, b, "math", t, None) for k, b, t in FORMULAS]
            + [(k, b, "algo", None, cap) for k, cap, b in ALGOS]
            + spec_entries())
    manifest = {"_render": {"dpi": DPI, "base_pt": BASE_PT, "ink": INK}}
    failures = []
    for key, body, kind, tag, caption in jobs:
        try:
            dest, im = render(key, body, kind, caption)
        except RuntimeError as exc:
            failures.append((key, str(exc)[:400]))
            print("FAIL %-10s %s" % (key, str(exc).splitlines()[0]))
            continue
        w, h = im.size
        entry = {"file": "assets/formulas/%s.png" % key, "w": w, "h": h,
                 "dpi": DPI, "kind": kind}
        if tag:
            entry["tag"] = tag
        if caption:
            entry["caption"] = caption
        manifest[key] = entry
        print("%-10s %5dx%-6d %s" % (key, w, h, kind))
        if kind == "algo":
            emit_halves(dest, key, im, manifest)
    with open(os.path.join(OUTDIR, "manifest.json"), "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
    print("wrote %s" % os.path.join(OUTDIR, "manifest.json"))
    if failures:
        print("\n%d FAILURES:" % len(failures))
        for key, msg in failures:
            print("=== %s ===\n%s" % (key, msg))
        return 1
    print("ok - %d images" % (len(manifest) - 1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
