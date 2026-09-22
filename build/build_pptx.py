"""Build the SEU x NARI pptx template. See _pptx_kit.py for the XML helpers.

Usage:
    python build_pptx.py            # all phases, writes out/seu_nari_template.pptx
    python build_pptx.py --phase A  # theme + master only  -> build/_shell.pptx
    python build_pptx.py --phase B  # + the 15 layouts     -> build/_shell.pptx
    python build_pptx.py --phase C  # + example slides     -> out/...
"""
import copy
import os
import sys
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml import parse_xml
from pptx.parts.slide import SlideLayoutPart
from pptx.util import Emu, Inches, Pt

import _pptx_kit as K

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
ASSETS = os.path.join(ROOT, "assets")
SHELL = os.path.join(ROOT, "build", "_shell.pptx")
OUT = os.path.join(ROOT, "out", "seu_nari_template.pptx")
PPTX_SLIDE_LAYOUT = ("application/vnd.openxmlformats-officedocument."
                     "presentationml.slideLayout+xml")

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# --- deck content (not design tokens) ---------------------------------------
# 本科毕业设计（论文）答辩 · 北京邮电大学
DECK = {
    "author_inst": "张珂 · 东南大学电气工程学院",
    "cover_kicker": "研究生第一学年研究进展汇报",
    "cover_title": "面向时序预测的轻量化模型研究",
    "cover_en": "Lightweight Modeling for Time-Series Forecasting",
    "cover_meta": [("汇报人", "张珂"), ("导师", "×××　教授"),
                   ("学院", "东南大学电气工程学院"), ("日期", "2026 年 9 月")],
    # 致谢页的三个字符串也在这里，模板代码里不写死任何文案。
    "end_title": "感谢聆听",
    "end_sub": "敬请批评指正",
    "end_date": "2026 年 9 月",
    # 章节数由列表长度决定（NSEC）。增删这里的条目，版式、进度段数、刻度
    # 全部自动跟随，不需要改任何常量。
    "sections": ["研究背景", "相关工作", "方法设计", "实验与分析", "总结与展望"],
    "sections_en": ["BACKGROUND", "RELATED WORK", "METHODOLOGY",
                    "EXPERIMENTS", "CONCLUSION"],
}

# Section count is derived, not typed: it drives the green progress segments, the
# divider tick marks and both layout loops. Four separate `range(1, 7)` literals
# used to have to be edited together, and missing one failed silently (bar) or
# with a KeyError (layouts).
NSEC = len(DECK["sections"])

# Typefaces come from design/tokens.json so there is one place to change them.
_TOK = K.load_tokens()
EA = _TOK["type"]["cjk-stack"][0]          # CJK face
LATIN = _TOK["type"]["latin-stack"][0]     # Latin / digits face
EA_FALLBACK = _TOK["type"]["cjk-stack"][1]

SEAL = os.path.join(ASSETS, "seu-seal.png")
SEU_LOCKUP = os.path.join(ASSETS, "seu-logo.png")
NARI_LOGO = os.path.join(ASSETS, "nari-logo.png")

# Registered once on the master so no media partname is duplicated.
IMAGE_FILES = {"seal": SEAL, "seu-logo": SEU_LOCKUP, "nari-logo": NARI_LOGO}

# --- geometry (px on the 1280x720 canvas) -----------------------------------
M = 53          # outer margin
SEAL_CM = 2.91        # 用户指定尺寸；换算成画布 px
SEAL_PX = round(K.cm(SEAL_CM))   # = 110 px
HEADER_TOP = 72
BAR_H = 9
BAR_W = 48
BAR_GAP = 7
BAR_TOP = 93.5  # vertically centred on the 52px header row
FOOT_RULE_Y = 660
FOOT_TEXT_Y = 668
CONTENT_TOP = 148
TITLE_RULE_Y = 0


# ============================================================================
# theme
# ============================================================================

def theme_xml(tok):
    scheme = (
        '<a:clrScheme xmlns:a="%s" name="SEU-NARI">'
        '<a:dk1><a:srgbClr val="%s"/></a:dk1>'
        '<a:lt1><a:srgbClr val="%s"/></a:lt1>'
        '<a:dk2><a:srgbClr val="%s"/></a:dk2>'
        '<a:lt2><a:srgbClr val="%s"/></a:lt2>'
        '<a:accent1><a:srgbClr val="%s"/></a:accent1>'
        '<a:accent2><a:srgbClr val="%s"/></a:accent2>'
        '<a:accent3><a:srgbClr val="%s"/></a:accent3>'
        '<a:accent4><a:srgbClr val="%s"/></a:accent4>'
        '<a:accent5><a:srgbClr val="%s"/></a:accent5>'
        '<a:accent6><a:srgbClr val="%s"/></a:accent6>'
        '<a:hlink><a:srgbClr val="%s"/></a:hlink>'
        '<a:folHlink><a:srgbClr val="%s"/></a:folHlink>'
        '</a:clrScheme>'
        % (A_NS, K.c(tok, "ink"), K.c(tok, "surface"),
           K.c(tok, "primary-dark"), K.c(tok, "surface-alt"),
           K.c(tok, "primary"), K.c(tok, "secondary"), K.c(tok, "accent"),
           K.c(tok, "primary-light"), K.c(tok, "secondary-dark"), K.c(tok, "ink-faint"),
           K.c(tok, "secondary"), K.c(tok, "primary-dark"))
    )

    def font(tag, name):
        return ('<a:%s xmlns:a="%s" name="%s">'
                '<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface="%s"/>'
                '<a:font script="Hans" typeface="%s"/>'
                '<a:font script="Hant" typeface="%s"/>'
                '</a:%s>' % (tag, A_NS, name, LATIN, EA, LATIN, EA, EA, tag))

    return scheme, font("majorFont", "Palatino Linotype"), font("minorFont", "Palatino Linotype")


def rewrite_theme(prs, tok):
    theme_part = prs.slide_masters[0].part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)
    ns = {"a": A_NS}
    te = root.find("a:themeElements", ns)
    clr_xml, major_xml, minor_xml = theme_xml(tok)
    te.replace(te.find("a:clrScheme", ns), etree.fromstring(clr_xml))
    fs = te.find("a:fontScheme", ns)
    fs.replace(fs.find("a:majorFont", ns), etree.fromstring(major_xml))
    fs.replace(fs.find("a:minorFont", ns), etree.fromstring(minor_xml))
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                     standalone=True)
    return theme_part


# ============================================================================
# master
# ============================================================================

def tx_styles(tok):
    """Master default text styles so new textboxes inherit the right font/size."""
    def dflt(sz, bold, color_key):
        b = ' b="1"' if bold else ''
        return ('<a:defRPr sz="%d"%s>'
                '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
                '<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface="%s"/>'
                '</a:defRPr>' % (sz, b, K.c(tok, color_key), LATIN, EA, LATIN))

    title = ('<p:titleStyle xmlns:p="%s" xmlns:a="%s">'
             '<a:lvl1pPr algn="l">%s</a:lvl1pPr></p:titleStyle>'
             % (P_NS, A_NS, dflt(2800, True, "ink")))
    body = ('<p:bodyStyle xmlns:p="%s" xmlns:a="%s">'
            '<a:lvl1pPr marL="171450" indent="-171450">%s</a:lvl1pPr></p:bodyStyle>'
            % (P_NS, A_NS, dflt(1600, False, "ink")))
    other = ('<p:otherStyle xmlns:p="%s" xmlns:a="%s">'
             '<a:defPPr>%s</a:defPPr></p:otherStyle>'
             % (P_NS, A_NS, dflt(1400, False, "ink")))
    return title, body, other


def build_master(prs, tok):
    master = prs.slide_masters[0]
    el = master.element

    # white page background
    cSld = el.find(qn("p:cSld"))
    for old in cSld.findall(qn("p:bg")):
        cSld.remove(old)
    bg = parse_xml(
        '<p:bg %s><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg>' % (nsdecls("p", "a"), K.c(tok, "surface")))
    cSld.insert(0, bg)

    # default text styles
    title, body, other = tx_styles(tok)
    for old in el.findall(qn("p:txStyles")):
        el.remove(old)
    tx = parse_xml(
        '<p:txStyles %s>%s%s%s</p:txStyles>' % (nsdecls("p", "a"), title, body, other))
    # p:sldMaster requires txStyles before extLst, so insert rather than append
    ext = el.find(qn("p:extLst"))
    if ext is not None:
        ext.addprevious(tx)
    else:
        el.append(tx)

    spTree = K.spTree_of(master)
    K.clear_shapes(spTree)

    # Register every image exactly once here; the layouts reuse these parts, so
    # the package ends up with three media parts and no duplicate partnames.
    image_parts = K.preload_images(master.part, IMAGE_FILES)

    # The SEU seal is NOT on the master: it appears on content pages only, so it
    # lives on the "内容 NN" layouts (nothing to add here).

    # footer hairline + texts
    spTree.append(K.rect(K.c(tok, "rule"), K.px(M), K.px(FOOT_RULE_Y),
                         K.px(1280 - 2 * M), K.px(1), name="页脚分隔线"))
    spTree.append(K.textbox(
        K.px(M), K.px(FOOT_TEXT_Y), K.px(620), K.px(24),
        [{"runs": [(DECK["author_inst"], K.c(tok, "ink-muted"), False)],
          "size": 11, "latin": LATIN, "ea": EA}], name="页脚署名"))

    guid = "{A1B2C3D4-0001-4E5F-9A00-000000000001}"
    spTree.append(K.textbox_pxml(
        K.px(1280 - M - 240), K.px(FOOT_TEXT_Y), K.px(240), K.px(24),
        K.fld_run("slidenum", 11, K.c(tok, "ink-muted"), LATIN, EA, guid),
        name="页码"))
    return image_parts


# ============================================================================
# shared layout ornament (geometry measured off the approved concept boards)
# ============================================================================

BAR_TOP = 93.5          # centred on the 52px header row that starts at y=72
BAR_PITCH = 55          # BAR_W + BAR_GAP
TRACK = "DDE1D8"
EYEBROW_Y = 144
TITLE_Y = 166
RULE_Y = 228
CONTENT_Y = 254
BODY_W = 1280 - 2 * M   # 1174


def bar_sp(tok, k_filled):
    """The NSEC-segment progress bar track: segments 1..k_filled are green.

    No gold marker here. The gold sub-progress bar is proportional to how far the
    author is through the current section, which is a property of the PAGE, not of
    the section — so slides.py draws it (see gold_sub_bar). Keeping the green
    fill on the layout still means inserting a slide inside a section cannot
    break the section-level bar.
    """
    out = []
    for i in range(1, NSEC + 1):
        x = M + (i - 1) * BAR_PITCH
        fill = K.c(tok, "primary") if i <= k_filled else TRACK
        out.append(K.sp_str(K.rect(fill, K.px(x), K.px(BAR_TOP), K.px(BAR_W),
                                   K.px(BAR_H), name="进度段%d" % i)))
    return "".join(out)


def seal_sp():
    """SEU seal for the top-right of CONTENT pages, vertically centred on the bar.

    Uses picture_slot() because the layout part does not exist yet when the body
    XML is built; register_layout() swaps in the real rId afterwards.
    """
    return K.sp_str(K.picture_slot(
        "seal", K.px(1280 - M - SEAL_PX),
        K.px(BAR_TOP + BAR_H / 2 - SEAL_PX / 2),
        K.px(SEAL_PX), K.px(SEAL_PX), "东大校徽"))


def title_block(tok):
    """Title placeholder + gold hairline rule.

    No eyebrow here: text on a layout cannot be clicked in Normal view, so every
    string the author might want to change lives on the slide instead (see
    slides.py). Layouts carry only shapes, logos and placeholders.
    """
    out = [K.title_ph(K.px(M), K.px(TITLE_Y), K.px(BODY_W), K.px(50),
                      28, K.c(tok, "ink"), LATIN, EA),
           K.sp_str(K.rect(K.c(tok, "accent"), K.px(M), K.px(RULE_Y),
                           K.px(192), K.px(2), name="金色标题细线"))]
    return "".join(out)


def diamonds(tok, specs):
    """specs: list of (x, y, size, kind) where kind is 'green'|'gold'|'outline'."""
    out = []
    for x, y, size, kind in specs:
        if kind == "green":
            out.append(K.sp_str(K.rect(K.c(tok, "primary"), K.px(x), K.px(y),
                                       K.px(size), K.px(size), rot=2700000,
                                       name="绿菱形")))
        elif kind == "gold":
            out.append(K.sp_str(K.rect(K.c(tok, "accent"), K.px(x), K.px(y),
                                       K.px(size), K.px(size), rot=2700000,
                                       name="金菱形")))
        else:
            out.append(K.sp_str(K.rect(None, K.px(x), K.px(y), K.px(size), K.px(size),
                                       rot=2700000, line=K.c(tok, "rule-strong"),
                                       line_w=14288, name="描边菱形")))
    return "".join(out)


def badge(tok):
    """Section-divider badge: green square + rotated gold + outline diamond.

    The numeral is NOT drawn here — it lives on the slide so the author can edit
    it. Positions come from measuring the approved board, not from re-deriving
    the CSS transform stack: the CSS `rotate(45deg) translate(dx,dy)` moves a
    shape's centre by R(45 deg) * (dx, dy).
    """
    return "".join([
        diamonds(tok, [(136.5, 379.8, 148, "gold"),      # centre (210.5, 453.8)
                       (108.3, 187.6, 196, "outline")]),  # centre (206.3, 285.6)
        K.sp_str(K.rect(K.c(tok, "primary"), K.px(105), K.px(255), K.px(212),
                        K.px(212), name="绿方块")),
    ])


def marks_sp(tok, k_filled):
    out = []
    for i in range(1, NSEC + 1):
        x = 378 + (i - 1) * 39
        fill = K.c(tok, "primary") if i <= k_filled else K.c(tok, "rule")
        out.append(K.sp_str(K.rect(fill, K.px(x), K.px(404), K.px(34), K.px(4),
                                   name="刻度%d" % i)))
    return "".join(out)


# ============================================================================
# phase B — the layout set
# ============================================================================

def layout_cover(tok):
    """Cover chrome only: the two logos, their divider, and the diamonds.
    All cover text lives on the slide so it stays clickable."""
    seu_w = 330 * 80 / 128          # keep the lockup's aspect at 80px tall
    nari_w = 1088 * 50 / 206
    return "".join([
        K.sp_str(K.picture_slot("seu-logo", K.px(64), K.px(80),
                                K.px(seu_w), K.px(80))),
        K.sp_str(K.rect(K.c(tok, "rule"), K.px(64 + seu_w + 26), K.px(89),
                        K.px(1), K.px(62), name="logo分隔线")),
        K.sp_str(K.picture_slot("nari-logo", K.px(64 + seu_w + 52), K.px(95),
                                K.px(nari_w), K.px(50))),
        diamonds(tok, [(972, 96, 190, "outline"),
                       (900, 128, 230, "green"),
                       (1034, 286, 150, "gold")]),
    ])


def layout_toc(tok):
    """Progress bar + title placeholder + gold rule. The six TOC entries are
    drawn on the slide (see slides.c_toc) so they can be renamed in place."""
    return bar_sp(tok, 1) + title_block(tok)


def layout_divider(tok, no):
    """Badge shapes, gold rule and the section marks. The numeral, section title
    and English line are drawn on the slide so they can be edited."""
    return "".join([
        badge(tok),
        K.sp_str(K.rect(K.c(tok, "accent"), K.px(378), K.px(384), K.px(112),
                        K.px(2), name="金线")),
        marks_sp(tok, no),
    ])


def layout_content(tok, no):
    """Progress bar + seal + title + gold rule. Content pages are the only ones
    carrying the seal, per feedback."""
    return bar_sp(tok, no) + seal_sp() + title_block(tok)


def layout_end(tok):
    """Gold rule + diamonds; the 谢谢聆听 block is on the slide."""
    return "".join([
        K.sp_str(K.rect(K.c(tok, "accent"), K.px(602), K.px(386), K.px(76),
                        K.px(2), name="致谢金线")),
        diamonds(tok, [(1010, 120, 190, "outline"),
                       (932, 152, 230, "green"),
                       (1068, 300, 150, "gold")]),
    ])


def build_layouts(prs, tok, image_parts):
    master = prs.slide_masters[0]
    specs = [("封面", layout_cover(tok), False),
             ("目录", layout_toc(tok), True)]
    for i in range(1, NSEC + 1):
        specs.append(("章节过渡 %02d" % i, layout_divider(tok, i), True))
    for i in range(1, NSEC + 1):
        specs.append(("内容 %02d" % i, layout_content(tok, i), True))
    specs.append(("致谢", layout_end(tok), True))

    K.drop_all_layouts(master)
    made = []
    for idx, (name, body, show_master) in enumerate(specs):
        made.append(K.register_layout(prs, master, name, body,
                                      show_master_sp=show_master,
                                      part_index=200 + idx,
                                      image_parts=image_parts))
    return made


# ============================================================================
# main
# ============================================================================

def _normalize_zip(path, stamp=(2026, 1, 1, 0, 0, 0)):
    """Rewrite `path` with a fixed timestamp on every zip entry.

    python-pptx stamps the current time on each entry, so two builds of identical
    content produce different file hashes. Every part is byte-identical (verified)
    — only the zip metadata moves. For a repo that tracks out/*.pptx that means a
    spurious 1 MB binary diff on every rebuild, so pin the timestamps.
    """
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as src:
        infos = src.infolist()
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in infos:
                fixed = zipfile.ZipInfo(info.filename, date_time=stamp)
                fixed.compress_type = info.compress_type
                fixed.external_attr = info.external_attr
                fixed.internal_attr = info.internal_attr
                fixed.create_system = info.create_system
                dst.writestr(fixed, src.read(info.filename))
    os.replace(tmp, path)


def new_prs():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    return prs


def main():
    phase = "ABC"
    if "--phase" in sys.argv:
        phase = sys.argv[sys.argv.index("--phase") + 1].upper()
    tok = K.load_tokens()

    prs = new_prs()
    rewrite_theme(prs, tok)
    image_parts = build_master(prs, tok)
    os.makedirs(os.path.dirname(SHELL), exist_ok=True)

    if "B" in phase or "C" in phase:
        made = build_layouts(prs, tok, image_parts)
        print("phase B ok -> %d layouts: %s" % (len(made), [l.name for l in made]))

    prs.save(SHELL)
    print("saved shell -> %s" % SHELL)

    if "C" in phase:
        import slides as S
        deck = Presentation(SHELL)
        n = len(S.build_slides(deck, tok))
        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        deck.save(OUT)
        _normalize_zip(OUT)          # 固定 zip 时间戳，重建不再产生假 diff
        print("phase C ok -> %d slides -> %s" % (n, OUT))
    return 0


if __name__ == "__main__":
    sys.exit(main())
