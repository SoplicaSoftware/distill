#!/usr/bin/env python3
"""edu_extract.py — prepare lecture slides (PDF or PPTX) for concept-map notes.

Splits handout pages into slides (detects slide frames; 1, 2, 3, 6 per page),
extracts per-slide text, crops figures from a page render (works for images
with transparency masks, e.g. chemical formulas), drops repeated decorations
(logos) and writes:

  OUT/slides.md            compact per-slide text for Claude (cheap to read)
  OUT/extract.json         machine-readable skeleton: slides, text, figures
  OUT/images/sNN_fK.png    figures (cropped at --dpi)
  OUT/slides/sNN.png       slide renders (for a future reader; not for Claude)
  OUT/contact/sheet-N.png  labelled thumbnail grids, so Claude can caption
                           many figures with one image read

usage: edu_extract.py FILE.(pdf|pptx) -o OUT [--dpi 200] [--min-fig 2.0]
"""
import argparse, hashlib, json, os, re, shutil, subprocess, sys, tempfile
import pdfplumber
from PIL import Image, ImageDraw, ImageFont


def to_pdf(path, tmp):
    if path.lower().endswith(".pdf"):
        return path
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    return os.path.join(tmp, os.path.splitext(os.path.basename(path))[0] + ".pdf")


def slide_frames(page):
    """Big rectangles with slide-like aspect = slide frames on a handout page."""
    W, H = page.width, page.height
    cands = []
    for r in page.rects + [c for c in page.curves if len(c.get("pts", [])) in (4, 5)]:
        w, h = r["x1"] - r["x0"], r["bottom"] - r["top"]
        if w < 0.25 * W or h < 0.15 * H or w > 0.98 * W:
            continue
        if not 1.2 <= w / h <= 1.9:
            continue
        cands.append((round(r["x0"]), round(r["top"]), round(r["x1"]), round(r["bottom"])))
    # dedupe near-identical, drop frames containing other frames
    uniq = []
    for c in sorted(set(cands)):
        if not any(abs(c[0] - u[0]) < 4 and abs(c[1] - u[1]) < 4 and abs(c[2] - u[2]) < 4 for u in uniq):
            uniq.append(c)
    uniq = [c for c in uniq if not any(o != c and o[0] >= c[0] - 2 and o[1] >= c[1] - 2 and o[2] <= c[2] + 2 and o[3] <= c[3] + 2 for o in uniq)]
    if not uniq:
        return [(0, 0, W, H)], False
    uniq.sort(key=lambda b: (round(b[1] / 20), b[0]))
    return uniq, True


def slide_number_below(page, box):
    """Handouts print the slide number just under the frame."""
    words = page.within_bbox((max(0, box[0] - 5), box[3], min(page.width, box[2]), min(page.height, box[3] + 25))).extract_words()
    for w in words:
        if re.fullmatch(r"\d{1,3}", w["text"]):
            return int(w["text"])
    return None


def slide_text(page, box):
    crop = page.within_bbox(box)
    words = crop.extract_words(extra_attrs=["size"], keep_blank_chars=False, use_text_flow=False)
    title = ""
    if words:
        top_zone = [w for w in words if w["top"] < box[1] + 0.3 * (box[3] - box[1])]
        if top_zone:
            mx = max(w["size"] for w in top_zone)
            if mx >= 1.25 * sorted(w["size"] for w in words)[len(words) // 2]:
                title = " ".join(w["text"] for w in sorted(top_zone, key=lambda w: (round(w["top"]), w["x0"])) if w["size"] >= mx - 0.5)
    txt = crop.extract_text(layout=False) or ""
    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r"\n\s*\n+", "\n", txt).strip()
    return title.strip(), txt


def merge_boxes(boxes, gap):
    boxes = [list(b) for b in boxes]
    changed = True
    while changed:
        changed = False
        out = []
        while boxes:
            a = boxes.pop()
            i = 0
            while i < len(boxes):
                b = boxes[i]
                if a[0] - gap <= b[2] and b[0] - gap <= a[2] and a[1] - gap <= b[3] and b[1] - gap <= a[3]:
                    a = [min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])]
                    boxes.pop(i)
                    changed = True
                else:
                    i += 1
            out.append(a)
        boxes = out
    return boxes


def ahash(img):
    g = img.convert("L").resize((12, 12))
    px = list(g.tobytes())
    m = sum(px) / len(px)
    return "".join("1" if p > m else "0" for p in px)


def is_blank(img):
    g = img.convert("L")
    lo, hi = g.getextrema()
    return hi - lo < 12


def contact_sheets(figs, outdir, cols=4, cell=340, per=16):
    os.makedirs(outdir, exist_ok=True)
    try:
        font = ImageFont.truetype("DejaVuSans-Bold.ttf", 18)
    except Exception:
        font = ImageFont.load_default()
    sheets = []
    for s in range(0, len(figs), per):
        chunk = figs[s:s + per]
        rows = (len(chunk) + cols - 1) // cols
        sheet = Image.new("RGB", (cols * cell, rows * (cell + 26)), "white")
        d = ImageDraw.Draw(sheet)
        for i, (fid, path) in enumerate(chunk):
            im = Image.open(path).convert("RGB")
            im.thumbnail((cell - 12, cell - 12))
            x, y = (i % cols) * cell, (i // cols) * (cell + 26)
            sheet.paste(im, (x + (cell - im.width) // 2, y + 26 + (cell - im.height) // 2))
            d.rectangle([x + 2, y + 2, x + cell - 2, y + cell + 24], outline="#999")
            d.text((x + 8, y + 4), fid, fill="#c00", font=font)
        p = os.path.join(outdir, f"sheet-{s // per + 1}.png")
        sheet.save(p)
        sheets.append(p)
    return sheets


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--dpi", type=int, default=200)
    ap.add_argument("--min-fig", type=float, default=0.8, help="min figure area, %% of slide")
    ap.add_argument("--slide-px", type=int, default=1200)
    a = ap.parse_args()

    out = a.out
    for d in ("images", "slides", "contact"):
        shutil.rmtree(os.path.join(out, d), ignore_errors=True)
        os.makedirs(os.path.join(out, d), exist_ok=True)
    tmp = tempfile.mkdtemp()
    pdf = to_pdf(a.file, tmp)
    scale = a.dpi / 72

    slides, figs_all, hashes = [], [], {}
    with pdfplumber.open(pdf) as doc:
        meta = doc.metadata or {}
        seq = 0
        for pno, page in enumerate(doc.pages, 1):
            base = os.path.join(tmp, f"p{pno}")
            subprocess.run(["pdftoppm", "-f", str(pno), "-l", str(pno), "-r", str(a.dpi), "-png", "-singlefile", pdf, base], check=True)
            render = Image.open(base + ".png")
            frames, handout = slide_frames(page)
            page_words = page.extract_words()
            for box in frames:
                seq += 1
                num = slide_number_below(page, box) if handout else None
                sn = num or seq
                title, text = slide_text(page, box)
                sid = f"s{sn:02d}"
                crop_box = tuple(int(v * scale) for v in box)
                sl_img = render.crop(crop_box)
                sl_img.thumbnail((a.slide_px, a.slide_px))
                sl_path = os.path.join(out, "slides", f"{sid}.png")
                sl_img.save(sl_path)
                # figures
                sarea = (box[2] - box[0]) * (box[3] - box[1])
                raw = []
                for im in page.images:
                    cx, cy = (im["x0"] + im["x1"]) / 2, (im["top"] + im["bottom"]) / 2
                    if not (box[0] <= cx <= box[2] and box[1] <= cy <= box[3]):
                        continue  # belongs to a neighbouring slide on the handout
                    b = (max(im["x0"], box[0]), max(im["top"], box[1]), min(im["x1"], box[2]), min(im["bottom"], box[3]))
                    if b[2] - b[0] < 3 or b[3] - b[1] < 3:
                        continue
                    raw.append(b)
                sarea = (box[2] - box[0]) * (box[3] - box[1])
                # vector drawings (diagrams, arrows, formula bonds): cluster lines/curves/rects
                vec = []
                for o in page.lines + page.curves + page.rects:
                    cx, cy = (o["x0"] + o["x1"]) / 2, (o["top"] + o["bottom"]) / 2
                    if not (box[0] <= cx <= box[2] and box[1] <= cy <= box[3]):
                        continue
                    b = (max(o["x0"], box[0]), max(o["top"], box[1]), min(o["x1"], box[2]), min(o["bottom"], box[3]))
                    if (o["x1"] - o["x0"]) > 0.9 * (box[2] - box[0]) and (o["bottom"] - o["top"]) > 0.9 * (box[3] - box[1]):
                        continue  # slide frame / background
                    vec.append((b[0], b[1], max(b[2], b[0] + 0.5), max(b[3], b[1] + 0.5)))
                vgroups = []
                for g in merge_boxes(vec, gap=0.008 * page.width):
                    n = sum(1 for v in vec if v[0] >= g[0] - 1 and v[1] >= g[1] - 1 and v[2] <= g[2] + 1 and v[3] <= g[3] + 1)
                    gw, gh = g[2] - g[0], g[3] - g[1]
                    if n < 4 or gw * gh > 0.6 * sarea or gh < 0.06 * gw:
                        continue  # too few strokes, whole-slide frame, or table rule strip
                    vgroups.append(tuple(g))
                merged = merge_boxes(merge_boxes(raw, gap=0.005 * page.width) + vgroups, gap=0)
                merged = [m for m in merged if (m[2] - m[0]) * (m[3] - m[1]) <= 0.7 * sarea]
                merged = [m for m in merged if 100 * (m[2] - m[0]) * (m[3] - m[1]) / sarea >= a.min_fig]
                bw, bh = box[2] - box[0], box[3] - box[1]
                # title banners / header shapes: touch the top, wide, end in upper 40%
                merged = [m for m in merged if not ((m[1] - box[1]) < 0.12 * bh and (m[3] - box[1]) < 0.42 * bh
                                                    and (m[2] - m[0]) > 0.45 * bw)]
                merged.sort(key=lambda m: (round(m[1] / 15), m[0]))
                figs = []
                for k, m in enumerate(merged, 1):
                    pad = 3
                    cb = tuple(int(v * scale) for v in (m[0] - pad, m[1] - pad, m[2] + pad, m[3] + pad))
                    fim = render.crop(cb)
                    if is_blank(fim):
                        continue
                    fid = f"{sid}_f{k}"
                    fpath = os.path.join(out, "images", f"{fid}.png")
                    fim.save(fpath)
                    h = ahash(fim)
                    hashes.setdefault(h, []).append(fid)
                    bh_ = box[3] - box[1]
                    band = (max(box[0], m[0] - 0.03 * (box[2] - box[0])), min(box[3], m[1]),
                            min(box[2], m[2] + 0.03 * (box[2] - box[0])), min(box[3], m[3] + 0.09 * bh_))
                    near = " ".join(w["text"] for w in page_words
                                    if band[0] <= (w["x0"] + w["x1"]) / 2 <= band[2] and band[1] <= (w["top"] + w["bottom"]) / 2 <= band[3])
                    figs.append({"id": fid, "file": f"images/{fid}.png", "nearby_text": near[:160],
                                 "bbox_pct": [round(100 * (m[0] - box[0]) / (box[2] - box[0]), 1),
                                              round(100 * (m[1] - box[1]) / (box[3] - box[1]), 1),
                                              round(100 * (m[2] - box[0]) / (box[2] - box[0]), 1),
                                              round(100 * (m[3] - box[1]) / (box[3] - box[1]), 1)],
                                 "px": [fim.width, fim.height], "hash": h})
                toks = text.split()
                short = [t for t in toks if len(t) <= 2 and not re.fullmatch(r"\d+[.)]?|[-–•vüi]", t)]
                garbled = len(toks) > 8 and (len(short) / len(toks) > 0.3
                                             or any(ch in text for ch in "⇥⇤\uf8f0\uf8fb"))
                slides.append({"id": sid, "number": sn, "page": pno, "title": title, "text": text,
                               "text_garbled": garbled,
                               "render": f"slides/{sid}.png", "figures": figs})

    # repeated images (logos, headers) on >=3 slides -> decoration, removed
    deco = {h for h, ids in hashes.items() if len({i.split('_')[0] for i in ids}) >= 3}
    for s in slides:
        keep = []
        for f in s["figures"]:
            if f["hash"] in deco:
                os.remove(os.path.join(out, f["file"]))
            else:
                f.pop("hash")
                keep.append(f)
                figs_all.append((f["id"], os.path.join(out, f["file"])))
        s["figures"] = keep

    sheets = contact_sheets(figs_all, os.path.join(out, "contact"))
    extract = {"source": {"file": os.path.basename(a.file), "title": meta.get("Title", ""),
                          "author": meta.get("Author", ""), "pages": len(slides) and slides[-1]["page"]},
               "slides": slides, "contact_sheets": [os.path.relpath(p, out) for p in sheets]}
    with open(os.path.join(out, "extract.json"), "w") as f:
        json.dump(extract, f, ensure_ascii=False, indent=1)
    with open(os.path.join(out, "slides.md"), "w") as f:
        f.write(f"# {extract['source']['title'] or extract['source']['file']} — {len(slides)} slides, "
                f"{len(figs_all)} figures, contact sheets: {', '.join(extract['contact_sheets'])}\n")
        for s in slides:
            figs = "; ".join(f'{x["id"]}' + (f' «{x["nearby_text"][:70]}»' if x["nearby_text"] else "") for x in s["figures"]) or "-"
            warn = f"\n!! TEXT GARBLED (formulas/symbols) -> view {s['render']}" if s["text_garbled"] else ""
            f.write(f"\n## {s['id']} (p{s['page']}) {s['title']}\nFIGS: {figs}{warn}\n{s['text']}\n")
    shutil.rmtree(tmp, ignore_errors=True)
    print(f"{len(slides)} slides, {len(figs_all)} figures ({len(deco)} decoration hashes removed), "
          f"{len(sheets)} contact sheets -> {out}")


if __name__ == "__main__":
    main()
