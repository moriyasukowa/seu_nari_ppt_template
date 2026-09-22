"""Extract a bibliography from the text of an already-typeset PDF.

Why read the PDF instead of the .bib: if the source already rendered a correct
citation style (author order, [J]/[M]/[S] type marks, volume/issue/pages),
reusing that finished text beats reimplementing the style from the .bib.

What you must supply: `SOURCE_PDF` and the page range its references occupy.
What you must *check*: the four "chrome" patterns below — they strip the page
furniture (running head, page number, footer) that pdftotext interleaves with
the body. They are written for a beamer deck with a Chinese footer, so pointing
this at your own PDF without adapting them gives **dirty output, not an error**:
page numbers end up inside the entry text. `main()` warns when that happens.

    python prep_refs.py

Optional: the demo deck inlines its own short reference list, so this is only
needed when you have a source PDF to mine.
"""
import json
import os
import re
import shutil
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
# ---------------------------------------------------------------------------
# 改这里：源 PDF 与它的参考文献页码范围（1-based，含两端）。
# ---------------------------------------------------------------------------
SOURCE_PDF = os.path.join(ROOT, "你的文献页.pdf")
PDF = SOURCE_PDF
OUT = os.path.join(HERE, "_refs.json")
# TeX Live 的 bin 目录：优先环境变量 TEXBIN，其次 PATH 上的 pdflatex，
# 最后回退到本机安装位置。回退为空串时 pdflatex.exe 由 PATH 解析。
TEXBIN = (os.environ.get("TEXBIN")
          or os.path.dirname(shutil.which("pdflatex") or "")
          or r"D:\texlive\2024\bin\windows")

REF_PAGES = (1, 1)        # 1-based，含两端

# 期望的条目数；None = 不校验（推荐先设 None 跑一次看抽出来多少）。
EXPECTED = None

# ---------------------------------------------------------------------------
# 按你的源 PDF 调整：这些是"版面噪声"的模式，pdftotext 会把页眉/页脚/页码混进正文。
# 下面的默认值是按一份带中文页脚的 beamer 写的 —— 换成你自己的 PDF 时**必须核对**，
# 否则不会报错，只是抽出来的条目里混进页码和页眉（main() 会就此给出告警）。
# ---------------------------------------------------------------------------
_NAV = re.compile(r"研究背景.*总结与展望.*参考文献")     # 顶部导航条
_TITLE = re.compile(r"^参考文献\s*[IVX]+\s*$")            # 本页小标题
_PAGENO = re.compile(r"^\s*\d+\s*/\s*\d+\s*$")        # "12 / 57" 式页码
_FOOTER = ()                                             # 页脚里的固定串，如 ("姓名", "讲题")
_ENTRY = re.compile(r"^\s*\[(\d+)\]\s*(.*)$")           # "[12] 正文…"


def raw_text():
    if not os.path.exists(PDF):
        raise SystemExit(
            "SOURCE_PDF not found: %s\n"
            "  改 prep_refs.py 顶部的 SOURCE_PDF / REF_PAGES，或跳过这一步\n"
            "  （示例 deck 用的是 slides.py 里内联的文献列表）。" % PDF)
    p = subprocess.run([os.path.join(TEXBIN, "pdftotext.exe"),
                        "-f", str(REF_PAGES[0]), "-l", str(REF_PAGES[1]),
                        "-layout", PDF, "-"],
                       capture_output=True, text=True, encoding="utf-8")
    if p.returncode != 0:
        raise SystemExit("pdftotext failed: %s" % p.stderr[-800:])
    return p.stdout


def parse(text):
    entries = {}
    current = None
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if _NAV.search(stripped) or _TITLE.match(stripped) or _PAGENO.match(stripped):
            continue
        if any(f in stripped for f in _FOOTER):
            continue
        m = _ENTRY.match(line)
        if m:
            current = int(m.group(1))
            entries[current] = m.group(2).strip()
        elif current is not None:
            entries[current] += " " + stripped
    # tidy the wrapping artefacts
    out = []
    for n in sorted(entries):
        s = re.sub(r"\s+", " ", entries[n]).strip()
        s = s.replace("\uFB00", "ff").replace("\uFB01", "fi").replace("\uFB02", "fl")
        out.append((n, s))
    return out


def main():
    entries = parse(raw_text())
    nums = [n for n, _ in entries]
    if EXPECTED is not None and (len(entries) != EXPECTED
                                 or nums != list(range(1, EXPECTED + 1))):
        print("WARNING: got %d entries, numbers %s..%s"
              % (len(entries), nums[0] if nums else "-", nums[-1] if nums else "-"))
        missing = [i for i in range(1, EXPECTED + 1) if i not in nums]
        if missing:
            print("missing: %s" % missing)
    # The failure mode this script has is a *dirty result*, not an exception:
    # unadapted chrome patterns let page numbers and running heads through.
    junk = [n for n, t in entries
            if len(t) < 12 or re.match(r"^\s*\d+\s*/\s*\d+\s*$", t)]
    if junk:
        print("WARNING: entries %s look like page furniture, not citations.\n"
              "  上面 '按你的源 PDF 调整' 那几个模式大概没适配你的 PDF。"
              % junk[:5])
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump([{"n": n, "text": t} for n, t in entries], fh,
                  ensure_ascii=False, indent=1)
    for n, t in entries[:3]:
        print("[%d] %s" % (n, t[:110]))
    print("...")
    for n, t in entries[-2:]:
        print("[%d] %s" % (n, t[:110]))
    print("\n%d entries -> %s" % (len(entries), OUT))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
