"""Compose PNG pages into a single contact-sheet image for visual inspection.

Usage: python _contact.py <src_dir> <out.png> [cols] [cell_w] [range]
       python _contact.py ../_preview/deck ../_preview/sheet_10_30.png 3 560 10-30

`range` (A-B, 1-based, inclusive) keeps only those pages' files. Long decks are
reviewed in overlapping slices (see SKILL §7.3), and this is the tool for it --
otherwise you end up hand-splitting the PNGs.

Note: every cell is sized from the FIRST image, so a sheet mixing portrait and
landscape pages will crop the odd ones out. Keep slices inside one deck.
"""
import os
import sys

from PIL import Image, ImageDraw


def _keep(name, rng):
    """True if the file's leading number falls in rng ('10-30')."""
    if not rng:
        return True
    try:
        lo, hi = (int(x) for x in rng.split("-", 1))
    except ValueError:
        raise SystemExit("range must look like 10-30, got %r" % rng)
    digits = "".join(c for c in name if c.isdigit())
    return bool(digits) and lo <= int(digits[:len(str(hi))]) <= hi


def sheet(src_dir, out_png, cols=4, cell_w=520, pad=10, label_h=22, rng=None):
    files = sorted(f for f in os.listdir(src_dir)
                   if f.lower().endswith(".png") and _keep(f, rng))
    if not files:
        raise SystemExit("no PNGs in %s%s" % (src_dir, " for range %s" % rng if rng else ""))
    ims = []
    for f in files:
        im = Image.open(os.path.join(src_dir, f)).convert("RGB")
        h = int(im.height * cell_w / im.width)
        ims.append((f, im.resize((cell_w, h), Image.LANCZOS)))

    cell_h = ims[0][1].height
    rows = (len(ims) + cols - 1) // cols
    W = cols * cell_w + (cols + 1) * pad
    H = rows * (cell_h + label_h) + (rows + 1) * pad
    out = Image.new("RGB", (W, H), (245, 245, 245))
    d = ImageDraw.Draw(out)
    for i, (name, im) in enumerate(ims):
        r, c = divmod(i, cols)
        x = pad + c * (cell_w + pad)
        y = pad + r * (cell_h + label_h + pad)
        d.text((x + 3, y + 4), name, fill=(20, 20, 20))
        out.paste(im, (x, y + label_h))
    out.save(out_png)
    return out_png, out.size, len(ims)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    print(sheet(sys.argv[1], sys.argv[2],
                int(sys.argv[3]) if len(sys.argv) > 3 else 4,
                int(sys.argv[4]) if len(sys.argv) > 4 else 520,
                rng=sys.argv[5] if len(sys.argv) > 5 else None))
