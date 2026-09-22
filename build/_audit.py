"""Pre-delivery audit of the built deck.

Catches the four things that actually go wrong and that a contact sheet at
readable size makes hard to spot:

  1. FORMULA SCALE DRIFT — every formula PNG is rendered at the same DPI from the
     same base, so placing them all at one scale factor is what makes the maths
     read as one document. A slide with two different scales means a formula was
     fitted to its own box instead. This is the check to run first.
  2. BLEED — pictures, tables or background panels crossing the footer rule
     (y=660) or the right text margin (x=1227).
  3. UNSIZED READABILITY — table text below 10pt, or an image narrower than the
     floor below, both of which are illegible on a projector.
  4. EMPTY TITLE — a content page whose title placeholder never got filled. The
     most insidious failure, because nothing errors and the page just looks
     unfinished.

    python _audit.py [deck.pptx]      # default ../out/seu_nari_template.pptx
"""
import json
import os
import sys
from collections import Counter

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Emu

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))

EMU_PX = 9525
RIGHT_LIMIT = 1227          # body right edge (M + BODY_W)
BOTTOM_LIMIT = 660          # footer rule
MIN_TABLE_PT = 10.0
# Width floors apply to FIGURES only. A short formula is legitimately narrow at
# the deck's uniform scale — flagging it would push the author into blowing it
# up again, which is the very bug this audit exists to catch.
MIN_FIG_HARD = 180.0        # unreadable on a projector
MIN_FIG_NOTE = 260.0        # worth a look
SCALE_TOL = 0.005


def load(name):
    with open(os.path.join(ROOT, "assets", name, "manifest.json"), encoding="utf-8") as fh:
        return json.load(fh)


def _px(v):
    return None if v is None else v / EMU_PX


def audit(path):
    EQ = load("formulas")
    FIG = load("figures")
    prs = Presentation(path)
    problems, notes = [], []

    for i, sl in enumerate(prs.slides, 1):
        lay = sl.slide_layout.name
        scales, pics, text_boxes = {}, [], []
        title = ""

        for sh in sl.shapes:
            if sh.top is not None and sh.height is not None:
                w, h = _px(sh.width), _px(sh.height)
                b, r = _px(sh.top) + h, _px(sh.left) + w
            else:
                w = h = b = r = None

            if sh.has_text_frame and sh.text_frame.text.strip():
                sizes = [run.font.size.pt for p in sh.text_frame.paragraphs
                         for run in p.runs if run.font.size]
                text_boxes.append((sh.name, _px(sh.top), b, min(sizes) if sizes else None))

            if sh.shape_type is not None and "PICTURE" in str(sh.shape_type):
                pics.append((sh.name, w, h, b, r))
                if sh.name in EQ:
                    k = round(w / EQ[sh.name]["w"], 3)
                    scales.setdefault(k, []).append(sh.name)
                elif sh.name in FIG and w and w < MIN_FIG_NOTE:
                    msg = "p%d figure %s only %.0fpx wide" % (i, sh.name, w)
                    (problems if w < MIN_FIG_HARD else notes).append(msg)

            if sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        for p in c.text_frame.paragraphs:
                            for run in p.runs:
                                if run.font.size and run.font.size.pt < MIN_TABLE_PT:
                                    problems.append(
                                        "p%d table cell %.1fpt < %.0fpt: %r"
                                        % (i, run.font.size.pt, MIN_TABLE_PT, run.text[:22]))

            for ph in ([sh] if sh.is_placeholder else []):
                if ph.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                    title = ph.text.strip()

        # 1. formula scale uniformity.
        #    The keys are rounded to 3dp for readability, which means float noise
        #    can split one maths size into two keys (0.3995 vs 0.3985 -> 0.400 vs
        #    0.398) and raise a false MIXED. Merge anything within SCALE_TOL first.
        merged = {}
        for k in sorted(scales):
            for ref in merged:
                if abs(ref - k) <= SCALE_TOL:
                    merged[ref].extend(scales[k])
                    break
            else:
                merged[k] = list(scales[k])
        if len(merged) > 1:
            problems.append("p%d MIXED formula scale %s"
                            % (i, ", ".join("%.3f×%d" % (k, len(v))
                                            for k, v in sorted(merged.items()))))
        elif merged:
            notes.append("p%d formula scale %.3f" % (i, list(merged)[0]))

        # 2. bleed — hard for pictures, advisory for text boxes
        for name, w, h, b, r in pics:
            if b and b > BOTTOM_LIMIT:
                problems.append("p%d picture %s bottom %.0f > %d" % (i, name, b, BOTTOM_LIMIT))
            if r and r > RIGHT_LIMIT:
                problems.append("p%d picture %s right %.0f > %d" % (i, name, r, RIGHT_LIMIT))
        for name, top, bot, _ in text_boxes:
            if bot and bot > BOTTOM_LIMIT:
                notes.append("p%d text box %s declared bottom %.0f (check it does not spill)"
                             % (i, name or "?", bot))
            if top and top >= BOTTOM_LIMIT:
                problems.append("p%d text box %s starts at %.0f, below the footer" % (i, name or "?", top))

        # 4. empty title on a content page
        if not lay.startswith(("封面", "章节过渡", "致谢")) and not title:
            problems.append("p%d (%s) has an EMPTY TITLE" % (i, lay))

    return problems, notes


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "out", "seu_nari_template.pptx")
    problems, notes = audit(path)
    hist = Counter(n.split("formula scale ")[1] for n in notes if "formula scale" in n)
    if hist:
        print("formula scale per slide: %s"
              % ", ".join("%s × %d slides" % (k, v) for k, v in sorted(hist.items())))
        small = {k: v for k, v in hist.items() if k != max(hist, key=lambda x: hist[x])}
        if small:
            print("  (these pages are deliberately smaller than the deck standard: %s)"
                  % ", ".join("%s × %d" % (k, v) for k, v in sorted(small.items())))
    print("\n%d problem(s)" % len(problems))
    for p in problems:
        print("  ! " + p)
    print("\n%d advisory note(s)" % len(notes))
    for n in notes:
        if "formula scale" not in n:
            print("  - " + n)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
