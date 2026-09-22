"""Validate design/tokens.json: recompute WCAG contrast and compare with the
ratios the file declares.

Run after any colour change. Catches the failure mode where a hex is edited but
the documented ratio (and therefore the accessibility reasoning) silently goes
stale — that happened twice while building this deck.

    python design/check_tokens.py        # exit 1 if anything drifted
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TOKENS = os.path.join(HERE, "tokens.json")

# (declared key, foreground token, background token)
PAIRS = [
    ("white-on-primary", "white", "primary"),
    ("white-on-secondary", "white", "secondary"),
    ("ink-on-surface", "ink", "surface"),
    ("ink-on-surface-card", "ink", "surface-card"),
    ("greenDark-on-card", "primary-dark", "surface-card"),
    ("accent-on-surface", "accent", "surface"),
    ("inkMuted-on-surface", "ink-muted", "surface"),
]


def _lum(hexrgb):
    def chan(c):
        c /= 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hexrgb[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * chan(r) + 0.7152 * chan(g) + 0.0722 * chan(b)


def contrast(a, b):
    la, lb = _lum(a), _lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def verdict(ratio, spec):
    if spec.get("passes") == "decorative-only":
        return "decor-only"          # low contrast is the point; not a text pairing
    if ratio >= 7:
        return "AAA"
    if ratio >= 4.5:
        return "AA"
    if ratio >= 3.0:
        return "AA-large-only"
    return "FAIL"


def main():
    with open(TOKENS, encoding="utf-8") as fh:
        tok = json.load(fh)
    hex_of = lambda k: tok["color"][k]["hex"].lstrip("#").upper()
    declared = tok["contrast"]

    print("%-22s %-9s %s" % ("pair", "measured", "verdict"))
    ok = True
    for key, fg, bg in PAIRS:
        a, b = hex_of(fg), hex_of(bg)
        r = contrast(a, b)
        spec = declared.get(key, {})
        want = spec.get("ratio")
        drift = ""
        if want is None:
            drift = "  <-- not declared in tokens.json"
            ok = False
        elif abs(want - round(r, 1)) >= 0.15:
            drift = "  <-- tokens.json says %.1f" % want
            ok = False
        print("%-22s %5.2f:1    %-13s%s" % (key, r, verdict(r, spec), drift))

    if not ok:
        print("\n!! tokens.json contrast block is stale — update the declared values.")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
