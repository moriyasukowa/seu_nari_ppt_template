"""Build the SEU x NARI PowerPoint template.

python-pptx cannot add shapes to slide masters/layouts (LayoutShapes has no
``add_shape``) and has no ``add_slide_layout()``. So this builder does the
master/layout work by direct OOXML surgery — appending ``<p:sp>`` elements to
each ``<p:spTree>`` and creating extra layout parts with
``SlideLayoutPart.load()`` — then hands the shell to python-pptx to author the
slides, where its API is actually good.

Phases (run individually with --phase so each can be rendered and inspected):
    A  theme (colours + fonts) and master ornament (seal, footer rule, footer text)
    B  the 15 slide layouts, incl. the per-section progress bar
    C  example slides demonstrating every archetype

Design values are read from design/tokens.json — nothing is hand-copied.
"""
import copy
import json
import os
import sys

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml import parse_xml
from pptx.parts.slide import SlideLayoutPart
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
TOKENS = os.path.join(ROOT, "design", "tokens.json")
ASSETS = os.path.join(ROOT, "assets")
SHELL = os.path.join(ROOT, "build", "_shell.pptx")
OUT = os.path.join(ROOT, "out", "seu_nari_template.pptx")

PPTX_SLIDE_LAYOUT = ("application/vnd.openxmlformats-officedocument."
                     "presentationml.slideLayout+xml")

# --- px (96 dpi, 1280x720 design canvas) <-> EMU -----------------------------
PXE = 9525


def px(v):
    return Emu(int(round(v * PXE)))


def cm(v):
    """Centimetres -> px on the 1280x720 design canvas (96 px per inch).

    The canvas is 13.3333 in = 33.867 cm wide, so px = cm / 2.54 * 96.
    """
    return v / 2.54 * 96.0


def load_tokens():
    with open(TOKENS, encoding="utf-8") as fh:
        return json.load(fh)


def c(tok, key):
    """Token colour -> uppercase hex without '#'."""
    return tok["color"][key]["hex"].lstrip("#").upper()


# ============================================================================
# low-level XML helpers
# ============================================================================

_shape_id = [1000]


def _next_id():
    _shape_id[0] += 1
    return _shape_id[0]


def rect(hex_fill, x, y, w, h, rot=None, line=None, line_w=12700, name="rect", dash=None):
    """A filled rectangle (optionally rotated / outlined / dashed) as a <p:sp>.

    NOTE: rotation is an ATTRIBUTE on <a:xfrm> (rot="..."), not a child element —
    an <a:rot> child is silently ignored and the shape renders axis-aligned.
    """
    rot_attr = ' rot="%d"' % rot if rot is not None else ''
    xfrm = ('<a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            % (rot_attr, x, y, w, h))
    if line:
        d = '<a:prstDash val="%s"/>' % dash if dash else ''
        ln = ('<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</a:ln>'
              % (line_w, line, d))
    else:
        ln = '<a:ln><a:noFill/></a:ln>'
    fill = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % hex_fill) if hex_fill else '<a:noFill/>'
    return parse_xml(
        '<p:sp %s><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '%s%s</p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
        % (nsdecls("p", "a"), _next_id(), name, xfrm, fill, ln))


def _rpr(sz_pt, hex_color, bold, latin, ea, italic=False, spc=None, base=None):
    b = ' b="1"' if bold else ''
    i = ' i="1"' if italic else ''
    s = ' spc="%d"' % spc if spc is not None else ''
    # baseline is in 1/1000 percent: +30000 superscript, -25000 subscript.
    v = ' baseline="%d"' % base if base is not None else ''
    return ('<a:rPr lang="zh-CN" altLang="en-US" sz="%d"%s%s%s%s dirty="0">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface="%s"/>'
            '</a:rPr>' % (int(sz_pt * 100), b, i, s, v, hex_color, latin, ea, latin))


def _run(r, p):
    """One <a:r>.  ``r`` is (text, colour, bold) or (text, colour, bold, baseline)."""
    base = r[3] if len(r) > 3 else None
    return ('<a:r>%s<a:t>%s</a:t></a:r>'
            % (_rpr(p["size"], r[1], r[2], p["latin"], p["ea"],
                    p.get("italic", False), p.get("spc"), base), _esc(r[0])))


VIRTUAL_BODY = ('<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                'wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t">'
                '<a:normAutofit/></a:bodyPr>')


def textbox(x, y, w, h, paragraphs, name="tb", anchor="t", wrap=True):
    """paragraphs: list of dicts {runs:[(text,color,bold[,baseline])], size, latin, ea,
    align, line_spacing, spc_before, spc_after, marL, indent, bullet(color)}.

    ``bullet`` emits a square buChar in the given colour plus a hanging indent,
    a small square reads lighter than a large filled circle.

    A run may carry a fourth element, a DrawingML ``baseline`` shift, so that
    notation like UE_a^{s,e}(t) can stay editable text instead of an image.
    """
    body = ('<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'wrap="%s" lIns="0" tIns="0" rIns="0" bIns="0" anchor="%s">'
            '<a:normAutofit/></a:bodyPr>' % ("square" if wrap else "none", anchor))
    ps = []
    for p in paragraphs:
        runs = "".join(_run(r, p) for r in p["runs"])
        attrs = ''
        for attr in ("marL", "indent"):
            if p.get(attr) is not None:
                attrs += ' %s="%d"' % (attr, p[attr])
        ppr = '<a:pPr%s algn="%s">' % (attrs, p.get("align", "l"))
        if p.get("line_spacing"):
            ppr += '<a:lnSpc><a:spcPct val="%d"/></a:lnSpc>' % int(p["line_spacing"] * 100000)
        if p.get("spc_before"):
            ppr += '<a:spcBef><a:spcPts val="%d"/></a:spcBef>' % int(p["spc_before"] * 100)
        if p.get("spc_after"):
            ppr += '<a:spcAft><a:spcPts val="%d"/></a:spcAft>' % int(p["spc_after"] * 100)
        if p.get("bullet"):
            # order matters: buClr, buSz, buFont, buChar before defRPr
            ppr += ('<a:buClr><a:srgbClr val="%s"/></a:buClr>'
                    '<a:buFont typeface="Arial"/><a:buChar char="&#9642;"/>'
                    % p["bullet"])
        ppr += '</a:pPr>'
        ps.append('<a:p>' + ppr + runs + '</a:p>')
    return parse_xml(
        '<p:sp %s><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody>%s<a:lstStyle/>%s</p:txBody></p:sp>'
        % (nsdecls("p", "a"), _next_id(), name, x, y, w, h, body, "".join(ps)))


def textbox_pxml(x, y, w, h, paragraphs_xml, name="tb", anchor="t"):
    """Textbox from pre-built paragraph XML (needed for <a:fld> fields)."""
    body = ('<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="%s">'
            '<a:normAutofit/></a:bodyPr>' % anchor)
    return parse_xml(
        '<p:sp %s><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody>%s<a:lstStyle/>%s</p:txBody></p:sp>'
        % (nsdecls("p", "a"), _next_id(), name, x, y, w, h, body, paragraphs_xml))


def fld_run(field_type, sz_pt, hex_color, latin, ea, guid, align="r", placeholder="1"):
    """One paragraph holding a live field (e.g. slidenum). Returns a complete
    <a:p>, ready for textbox_pxml.

    NOTE: <a:fld> is a sibling of <a:r> inside <a:p> — it must not be nested
    inside an <a:r>, whose only allowed children are <a:rPr> and <a:t>.
    """
    rpr = ('<a:rPr lang="en-US" sz="%d" dirty="0">'
           '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
           '<a:latin typeface="%s"/><a:ea typeface="%s"/></a:rPr>'
           % (int(sz_pt * 100), hex_color, latin, ea))
    return ('<a:p><a:pPr algn="%s"/>'
            '<a:fld id="%s" type="%s">%s<a:t>%s</a:t></a:fld>'
            '</a:p>' % (align, guid, field_type, rpr, placeholder))


def sldlayout_xml(name, body_sp_xml, show_master_sp=True):
    """A complete <p:sldLayout> part body. ``body_sp_xml`` is concatenated
    <p:sp>/<p:pic> XML for the layout's own shapes."""
    return ('<p:sldLayout %s type="blank" preserve="1" showMasterSp="%d">'
            '<p:cSld name="%s"><p:spTree %s>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '%s</p:spTree></p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
            % (nsdecls("p", "a", "r"), 1 if show_master_sp else 0,
               name, nsdecls("a"), body_sp_xml))


def title_ph(x, y, w, h, sz_pt, hex_color, latin, ea, text="单击编辑标题"):
    """A real title placeholder so the outline view and title box work."""
    return (
        '<p:sp %s><p:nvSpPr><p:cNvPr id="%d" name="标题占位符"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t">'
        '<a:normAutofit/></a:bodyPr><a:lstStyle/>'
        '<a:p><a:pPr algn="l"/><a:r>%s<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
        % (nsdecls("p", "a"), _next_id(), x, y, w, h,
           _rpr(sz_pt, hex_color, True, latin, ea), _esc(text)))


def sp_str(el):
    """Serialise a shape element back to XML text (for concatenating bodies)."""
    from lxml import etree
    return etree.tostring(el, encoding="unicode")


SLIDE_LAYOUT_CT = ("application/vnd.openxmlformats-officedocument."
                   "presentationml.slideLayout+xml")
RID_SLOT = "RIDSLOT:"


def picture_slot(key, x, y, w, h, name="pic"):
    """A picture whose r:embed is a placeholder, resolved by register_layout()
    once the host part exists and can own the image relationship."""
    return picture(RID_SLOT + key, x, y, w, h, name)


def register_layout(prs, master, name, body_sp_xml, show_master_sp=True,
                    part_index=200, layout_id_base=2147483649, image_parts=None):
    """Build a new slide-layout part and wire it to the master.

    python-pptx has no add_slide_layout(); SlideLayoutPart.load() lets us create
    the part from a blob, then we add the two relationships (master->layout for
    the listing, layout->master for inheritance) and a <p:sldLayoutId> entry.

    ``image_parts`` maps placeholder keys to already-registered ImageParts. Add
    them once via preload_images() — calling get_or_add_image_part() on a
    hand-loaded part does not share the package's ImageParts registry, which
    yields duplicate media partnames and a corrupt file.

    NOTE: ``layout_id_base`` must be 2147483649, not 2147483648. PowerPoint
    reserves 0x80000000 and silently refuses to open any file that uses it as a
    <p:sldLayoutId id>, with a generic "发生意外" error that looks like a corrupt
    package but is not one.

    Returns the newly registered SlideLayout.
    """
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.opc.package import PackURI
    from pptx.parts.slide import SlideLayoutPart

    xml = sldlayout_xml(name, body_sp_xml, show_master_sp)
    partname = PackURI("/ppt/slideLayouts/slideLayout%d.xml" % part_index)
    # XmlPart.load signature is (partname, content_type, package, blob) — package
    # comes before blob, not after.
    part = SlideLayoutPart.load(partname, SLIDE_LAYOUT_CT, prs.part.package,
                                xml.encode("utf-8"))

    for blip in part._element.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if rid and rid.startswith(RID_SLOT):
            key = rid[len(RID_SLOT):]
            blip.set(qn("r:embed"), part.relate_to(image_parts[key], RT.IMAGE))

    rId = master.part.relate_to(part, RT.SLIDE_LAYOUT)
    part.relate_to(master.part, RT.SLIDE_MASTER)
    lst = master.element.find(qn("p:sldLayoutIdLst"))
    lst.append(parse_xml('<p:sldLayoutId %s id="%d" r:id="%s"/>'
                         % (nsdecls("p", "r"), layout_id_base + len(lst), rId)))
    return part.slide_layout


def preload_images(owner_part, paths):
    """Register every image once on the package, keyed by name.

    Returns {key: ImagePart}. Must be called on a part that the package itself
    created (the master is), so the package's ImageParts registry stays the one
    source of image partnames.
    """
    return {key: owner_part.get_or_add_image_part(path)[0]
            for key, path in paths.items()}


def add_pic_part(owner_part, image_part, x, y, w, h, name="pic"):
    """Place an already-registered ImagePart, relating it to ``owner_part``."""
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    return picture(owner_part.relate_to(image_part, RT.IMAGE), x, y, w, h, name)


def drop_all_layouts(master):
    """Remove every inherited Office layout so only ours remain."""
    for layout in list(master.slide_layouts):
        master.slide_layouts.remove(layout)


def _esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def picture(rId, x, y, w, h, name="pic"):
    return parse_xml(
        '<p:pic %s><p:nvPicPr><p:cNvPr id="%d" name="%s"/>'
        '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
        % (nsdecls("p", "a", "r"), _next_id(), name, rId, x, y, w, h))


def spTree_of(part_owner):
    """The <p:spTree> of a master or layout element."""
    return part_owner.element.find(qn('p:cSld')).find(qn('p:spTree'))


def clear_shapes(spTree):
    """Drop every shape but keep nvGrpSpPr / grpSpPr (required by the schema)."""
    keep = {qn('p:nvGrpSpPr'), qn('p:grpSpPr')}
    for child in list(spTree):
        if child.tag not in keep:
            spTree.remove(child)


def add_pic(owner_part, image_path, x, y, w, h, name="pic"):
    image_part, rId = owner_part.get_or_add_image_part(image_path)
    return picture(rId, x, y, w, h, name)
