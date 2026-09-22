"""Pull math / pseudocode blocks out of a LaTeX source, verbatim.

Re-typing dozens of display equations into PowerPoint would be slow and
typo-prone, so instead we read them straight out of the source and write them to
_math_spec.json for make_formulas.py to render.

**Splitting is by ``\\begin{frame}``** — i.e. this expects a *beamer* source. A
plain article has no frames, so collect() refuses to return a silently empty
result; see _FRAME_RE below if you need to adapt the split to another structure.

Normalisation, and why:
  * `standalone` cannot host a display environment at top level (align/gather
    error out with "Missing \\endgroup"), but it *can* host `aligned` /
    `gathered` inside `$\\displaystyle ...$` — which is how the template already
    renders formulas. So align -> aligned, gather -> gathered, equation -> inline.
  * `\\label{}` / `\\nonumber` are stripped: the deck keeps equation numbers as
    editable slide text, not baked into the image.

    python extract_math.py

Optional: the demo deck ships its own inline FORMULAS, so this is only needed
when the formulas already live in a LaTeX source. If SOURCE_TEX is missing the
script exits with a clear message.
"""
import json
import os
import re
from collections import OrderedDict

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
# ---------------------------------------------------------------------------
# 改这里：你的 LaTeX 源。注意本脚本按 beamer 的 frame 结构切分（见 _FRAME_RE），
# 所以源应当是 beamer 讲义/幻灯片，不是普通论文正文。
# ---------------------------------------------------------------------------
SOURCE_TEX = os.path.join(ROOT, "你的源文件.tex")
SRC = SOURCE_TEX
OUT = os.path.join(HERE, "_math_spec.json")

MATH_ENVS = ("equation", "align", "align*", "gather", "gather*", "equation*")
ALGO_ENVS = ("algorithmic",)

# 切分依据。默认按 beamer 的 frame；换别的源结构就改这一行，
# 例如普通论文可按 section 切：r"\section\{([^}]*)\}"。
# 改错的下场是 collect() 找不到任何块 —— 已在 collect() 里加了硬检查。
_FRAME_RE = re.compile(r"\\begin\{frame\}(\[[^\]]*\])?(\{[^{}]*\})?")


def _split_frames(text):
    """[(title, body)] in document order."""
    frames = []
    for m in _FRAME_RE.finditer(text):
        start = m.end()
        end = text.find(r"\end{frame}", start)
        if end < 0:
            continue
        title = (m.group(2) or "")[1:-1].strip()
        frames.append((title, text[start:end]))
    return frames


def _extract_env(body, env):
    """All bodies of `\\begin{env}...\\end{env}` in order."""
    out = []
    open_tag = r"\begin{%s}" % env
    close_tag = r"\end{%s}" % env
    i = 0
    while True:
        s = body.find(open_tag, i)
        if s < 0:
            return out
        e = body.find(close_tag, s)
        if e < 0:
            return out
        out.append(body[s + len(open_tag):e])
        i = e + len(close_tag)


def _strip_optarg(body, env):
    """Drop the [1] argument algpseudocode takes, if present."""
    if env in ALGO_ENVS and body.startswith("["):
        close = body.find("]")
        return body[close + 1:]
    return body


# _extract_env() hands back only the env *inner* body, so the replacement has to
# be re-applied by name rather than by rewriting the (already-gone) begin/end
# tags. `equation` needs no wrapper - it is a bare expression.
WRAP = {"align": "aligned", "align*": "aligned",
        "gather": "gathered", "gather*": "gathered"}


def normalize(body, env):
    b = re.sub(r"\\label\{[^}]*\}", "", body)
    b = b.replace(r"\nonumber", "")
    b = b.strip()
    inner = WRAP.get(env)
    if inner:
        b = "\\begin{%s}\n%s\n\\end{%s}" % (inner, b, inner)
    return b


def collect():
    if not os.path.exists(SRC):
        raise SystemExit(
            "SOURCE_TEX not found: %s\n"
            "  改 extract_math.py 顶部的 SOURCE_TEX 指向你的 .tex，或跳过这一步\n"
            "  （示例 deck 用的是 make_formulas.py 里的内联 FORMULAS，不需要它）。" % SRC)
    with open(SRC, encoding="utf-8") as fh:
        text = fh.read()
    spec = OrderedDict()
    frames = _split_frames(text)
    if not frames:
        # A file with display maths but no frame markers is a plain article, not a
        # beamer deck. Returning {} here would let make_formulas.py print "ok" and
        # the author would believe every formula had been rendered -- a silent
        # wrong result, which is worse than a failure.
        n_math = sum(len(_extract_env(text, env)) for env in MATH_ENVS + ALGO_ENVS)
        raise SystemExit(
            "no frame markers in %s — nothing to split on.\n"
            "  本脚本按 beamer 的 frame 结构切分，而该文件里%s。\n"
            "  两选一：把 SOURCE_TEX 指向 beamer 源；或改本文件顶部的 _FRAME_RE\n"
            "  以适配你的结构（普通论文可以按 section 切）。\n"
            "  直接跑下去只会写出空 JSON 并报 0 blocks，看起来像成功 —— 所以这里直接停。"
            % (SRC, ("有 %d 处公式环境但没有 frame" % n_math) if n_math
               else "既没有 frame 也没有公式环境"))
    for fi, (title, body) in enumerate(frames):
        blocks = []
        found = []
        for env in MATH_ENVS + ALGO_ENVS:
            for b in _extract_env(body, env):
                found.append((env, b))
        # document order matters: sort by position in the frame body
        found.sort(key=lambda eb: body.find(eb[1][:40] if eb[1] else ""))
        for bi, (env, b) in enumerate(found):
            kind = "algo" if env in ALGO_ENVS else "math"
            raw = _strip_optarg(b, env)
            body_out = raw.strip() if kind == "algo" else normalize(raw, env)
            if not body_out:
                continue
            blocks.append({
                "id": "f%02d_%d" % (fi, bi),
                "kind": kind,
                "env": env,
                "lines": body_out.count(r"\\") + 1,
                "chars": len(body_out),
                "body": body_out,
            })
        if blocks:
            spec["f%02d" % fi] = {"frame": fi, "title": title, "blocks": blocks}
    return spec


def main():
    spec = collect()
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(spec, fh, ensure_ascii=False, indent=2)
    total = 0
    print("%-6s %-8s %s" % ("frame", "blocks", "title"))
    for key, info in spec.items():
        total += len(info["blocks"])
        kinds = "".join("A" if b["kind"] == "algo" else "m" for b in info["blocks"])
        print("%-6s %-8s %-10s %s" % (key, len(info["blocks"]), kinds,
                                      info["title"][:48]))
        for b in info["blocks"]:
            print("        %-8s %-6s lines=%-4d chars=%d" %
                  (b["id"], b["kind"], b["lines"], b["chars"]))
    print("\n%d frames, %d blocks -> %s" % (len(spec), total, OUT))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
