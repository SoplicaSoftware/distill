#!/usr/bin/env python3
"""pptx_map.py — token-cheap PPTX map for the distill skill.

Produces <out>/map.md (text + layout + annotations, coords in % of slide)
and <out>/slide-NN.png renders only for slides where visuals matter.

usage: pptx_map.py FILE.pptx [-o OUT] [--slides 3-7,9] [--render auto|all|none]
                   [--width 1400]
"""
import argparse, colorsys, os, re, shutil, subprocess, sys, tempfile
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}


def parse_ranges(spec, n):
    if not spec:
        return list(range(1, n + 1))
    out = set()
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-")
            out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(part))
    return sorted(i for i in out if 1 <= i <= n)


def color_name(hexv):
    if not hexv:
        return None
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if s < 0.15:
        return "black" if l < 0.2 else "white" if l > 0.85 else "gray"
    h *= 360
    for lim, nm in ((15, "red"), (40, "orange"), (65, "yellow"), (165, "green"),
                    (200, "cyan"), (255, "blue"), (290, "purple"), (335, "pink"), (361, "red")):
        if h < lim:
            return nm


def xml_color(el):
    """Return hex or scheme name of first solidFill under el."""
    if el is None:
        return None
    c = el.find(".//a:solidFill/a:srgbClr", NS)
    if c is not None:
        return c.get("val")
    c = el.find(".//a:solidFill/a:schemeClr", NS)
    if c is not None:
        return "scheme:" + c.get("val")
    return None


def xml_color_any(el):
    c = el.find("a:srgbClr", NS)
    if c is not None:
        return c.get("val")
    c = el.find("a:schemeClr", NS)
    return "scheme:" + c.get("val") if c is not None else None


def col_label(v):
    if not v:
        return None
    if v.startswith("scheme:"):
        return v
    return f"{color_name(v)}#{v}"


class Ctx:
    def __init__(self, W, H):
        self.W, self.H = W, H

    def pct(self, x, y):
        return round(100 * x / self.W, 1), round(100 * y / self.H, 1)


def bbox(shape, ctx, off=(0, 0)):
    x = (shape.left or 0) + off[0]
    y = (shape.top or 0) + off[1]
    w, h = shape.width or 0, shape.height or 0
    x0, y0 = ctx.pct(x, y)
    x1, y1 = ctx.pct(x + w, y + h)
    return (x0, y0, x1, y1)


def fmt_box(b):
    return f"x{b[0]:g}–{b[2]:g}% y{b[1]:g}–{b[3]:g}%"


def line_endpoints(shape, b):
    """Endpoints (in %) for line/connector using flipH/flipV."""
    xfrm = shape._element.find(".//a:xfrm", NS)
    fh = xfrm is not None and xfrm.get("flipH") == "1"
    fv = xfrm is not None and xfrm.get("flipV") == "1"
    x0, y0, x1, y1 = b
    sx, ex = (x1, x0) if fh else (x0, x1)
    sy, ey = (y1, y0) if fv else (y0, y1)
    return (sx, sy), (ex, ey)


def arrowheads(shape):
    ln = shape._element.find(".//a:ln", NS)
    head = tail = False
    if ln is not None:
        he, te = ln.find("a:headEnd", NS), ln.find("a:tailEnd", NS)
        head = he is not None and he.get("type", "none") != "none"
        tail = te is not None and te.get("type", "none") != "none"
    return head, tail


def text_of(shape):
    if not getattr(shape, "has_text_frame", False) or not shape.has_text_frame:
        return ""
    t = "\n".join(p.text for p in shape.text_frame.paragraphs).strip()
    return re.sub(r"\n{2,}", "\n", t)


def table_md(tbl):
    rows = []
    for r in tbl.rows:
        rows.append("| " + " | ".join(c.text.replace("\n", " ").strip() for c in r.cells) + " |")
    if rows:
        rows.insert(1, "|" + "---|" * len(tbl.columns))
    return "\n".join(rows)


def walk(shapes, ctx, off=(0, 0)):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            # child offsets in groups are in group coordinate space; approximate
            yield from walk(sh.shapes, ctx, off)
        else:
            yield sh, off


def inside(pt, b, pad=1.0):
    return b[0] - pad <= pt[0] <= b[2] + pad and b[1] - pad <= pt[1] <= b[3] + pad


def rel(pt, b):
    w, h = (b[2] - b[0]) or 1, (b[3] - b[1]) or 1
    return round(100 * (pt[0] - b[0]) / w), round(100 * (pt[1] - b[1]) / h)


def _area(b):
    return (b[2] - b[0]) * (b[3] - b[1])


def where(pt, pics, exclude=None):
    """Smallest picture containing the point (pasted snippets sit on screenshots).
    With exclude=<name>, only pictures larger than that one count as parents."""
    min_area = _area(dict(pics)[exclude]) if exclude else -1
    hits = [(n, b) for n, b in pics if n != exclude and _area(b) > min_area and inside(pt, b)]
    if not hits:
        return "on slide"
    name, pb = min(hits, key=lambda nb: (nb[1][2] - nb[1][0]) * (nb[1][3] - nb[1][1]))
    rx, ry = rel(pt, pb)
    return f"on {name} @({rx}%,{ry}% of image)"


def centre(b):
    return ((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)


ANNOT_GEOMS = {"ellipse", "rect", "roundRect", "rightArrow", "leftArrow", "upArrow",
               "downArrow", "line", "straightConnector1", "bentConnector3",
               "curvedConnector3", "frame", "wedgeRectCallout", "wedgeRoundRectCallout",
               "cloudCallout", "star5", "mathMultiply"}


def analyze_slide(slide, idx, ctx):
    texts, pics, annots, tables, others = [], [], [], [], []
    for sh, off in walk(slide.shapes, ctx):
        b = bbox(sh, ctx, off)
        el = sh._element
        geom_el = el.find(".//a:prstGeom", NS)
        geom = geom_el.get("prst") if geom_el is not None else None
        tag = el.tag.split("}")[1]
        spPr = el.find(".//p:spPr", NS)
        ln = spPr.find("a:ln", NS) if spPr is not None else None
        line_col = col_label(xml_color(ln))
        if line_col is None:
            ref = el.find(".//p:style/a:lnRef", NS)
            if ref is not None:
                line_col = col_label(xml_color_any(ref))
        fill_el = spPr.find("a:solidFill", NS) if spPr is not None else None
        fill_col = col_label(xml_color(spPr if fill_el is not None else None))
        txt = text_of(sh)

        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE or tag == "pic":
            name = f"IMG{len(pics) + 1}"
            pics.append((name, b))
            continue
        if getattr(sh, "has_table", False) and sh.has_table:
            tables.append((b, table_md(sh.table)))
            continue
        if getattr(sh, "has_chart", False) and sh.has_chart:
            others.append(f"CHART {fmt_box(b)} type={sh.chart.chart_type}")
            continue
        is_line = tag == "cxnSp" or geom in ("line", "straightConnector1") or (geom or "").endswith("Connector3")
        if is_line:
            annots.append(("line", b, line_col, sh, txt))
            continue
        if txt and (fill_col is None and line_col is None or len(txt) > 40 or geom in (None, "rect")):
            texts.append((b, txt, fill_col))
            continue
        if geom in ANNOT_GEOMS or not txt:
            annots.append((geom or "shape", b, line_col or fill_col, sh, txt))
        else:
            texts.append((b, txt, fill_col))

    out = [f"## Slide {idx}"]
    title = slide.shapes.title.text.strip() if slide.shapes.title is not None and slide.shapes.title.has_text_frame else ""
    if title:
        out.append(f"TITLE: {title}")
    texts.sort(key=lambda t: (round(t[0][1] / 5), t[0][0]))
    for b, t, fc in texts:
        if t == title:
            continue
        bg = f" bg={fc}" if fc else ""
        loc = where(centre(b), pics)
        loc = "" if loc == "on slide" else f" {loc}"
        out.append(f"TEXT [{fmt_box(b)}{bg}{loc}]: {t}")
    for name, b in pics:
        area = (b[2] - b[0]) * (b[3] - b[1])
        kind = "snippet" if area < 150 else "picture"
        par = where(centre(b), pics, exclude=name)
        par = "" if par == "on slide" else f" {par}"
        out.append(f"{name}: {kind} {fmt_box(b)}{par}")
    for b, md in tables:
        out.append(f"TABLE [{fmt_box(b)}]:\n{md}")
    out.extend(others)
    if annots:
        out.append("ANNOTATIONS:")
    for kind, b, col, sh, txt in annots:
        c = f" {col}" if col else ""
        label = f' "{txt}"' if txt else ""
        if kind == "line":
            s, e = line_endpoints(sh, b)
            head, tail = arrowheads(sh)
            if head and tail:
                k = "double-arrow"
            elif tail:
                k = "arrow"
            elif head:
                k, s, e = "arrow", e, s
            else:
                k = "line"
            if k == "line":
                out.append(f"- line{c} ({s[0]:g},{s[1]:g})–({e[0]:g},{e[1]:g}) | {where(s, pics)} → {where(e, pics)}{label}")
            else:
                out.append(f"- {k}{c} FROM ({s[0]:g},{s[1]:g}) {where(s, pics)} TO ({e[0]:g},{e[1]:g}) {where(e, pics)}{label}")
        else:
            cen = centre(b)
            out.append(f"- {kind}{c} [{fmt_box(b)}] centre {where(cen, pics)}{label}")

    if slide.has_notes_slide:
        n = slide.notes_slide.notes_text_frame.text.strip() if slide.notes_slide.notes_text_frame else ""
        if n:
            out.append(f"NOTES: {n}")
    needs_render = bool(pics) or bool(tables and annots) or len(annots) >= 2 or any(o.startswith("CHART") for o in others)
    return "\n".join(out), needs_render


def render(pptx, slides, outdir, width):
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, pptx],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    pdf = os.path.join(tmp, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    done = []
    for i in slides:
        base = os.path.join(outdir, f"slide-{i:02d}")
        subprocess.run(["pdftoppm", "-f", str(i), "-l", str(i), "-png", "-singlefile",
                        "-scale-to-x", str(width), "-scale-to-y", "-1", pdf, base], check=True)
        done.append(base + ".png")
    shutil.copy(pdf, os.path.join(outdir, "deck.pdf"))
    shutil.rmtree(tmp, ignore_errors=True)
    return done


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--slides", default=None)
    ap.add_argument("--render", choices=["auto", "all", "none"], default="auto")
    ap.add_argument("--width", type=int, default=1400)
    a = ap.parse_args()
    out = a.out or os.path.join("/tmp/distill", os.path.splitext(os.path.basename(a.pptx))[0])
    os.makedirs(out, exist_ok=True)
    prs = Presentation(a.pptx)
    ctx = Ctx(prs.slide_width, prs.slide_height)
    sel = parse_ranges(a.slides, len(prs.slides))
    parts = [f"# {os.path.basename(a.pptx)} — {len(prs.slides)} slides, "
             f"{Emu(prs.slide_width).inches:.2f}x{Emu(prs.slide_height).inches:.2f}in. "
             f"Coords = % of slide (x→right, y→down)."]
    to_render = []
    for i, sl in enumerate(prs.slides, 1):
        if i not in sel:
            continue
        md, need = analyze_slide(sl, i, ctx)
        if a.render == "all" or (a.render == "auto" and need):
            to_render.append(i)
            md += f"\nRENDER: slide-{i:02d}.png"
        parts.append(md)
    if to_render:
        try:
            render(a.pptx, to_render, out, a.width)
        except Exception as e:
            parts.append(f"\n(!) render failed: {e}")
    with open(os.path.join(out, "map.md"), "w") as f:
        f.write("\n\n".join(parts) + "\n")
    print(f"map: {out}/map.md | renders: {len(to_render)} -> {to_render}")


if __name__ == "__main__":
    main()
