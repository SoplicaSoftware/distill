#!/usr/bin/env python3
"""crop.py — zoom into a region of a rendered slide/page, at high resolution.

usage: crop.py DECK.pdf PAGE X0 Y0 X1 Y1 [-o out.png] [--width 1400]
Coordinates in % of the page (same as map.md). Output width capped so the
crop costs about as many tokens as a normal slide render but shows detail.
"""
import argparse, subprocess
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("pdf"); ap.add_argument("page", type=int)
for k in ("x0", "y0", "x1", "y1"):
    ap.add_argument(k, type=float)
ap.add_argument("-o", "--out")
ap.add_argument("--width", type=int, default=1400)
a = ap.parse_args()
info = subprocess.run(["pdfinfo", "-f", str(a.page), "-l", str(a.page), a.pdf], capture_output=True, text=True).stdout
line = next(l for l in info.splitlines() if "size:" in l and "Page" in l)
pw, ph = (float(v) for v in line.split(":")[1].split("pts")[0].split("x"))
frac = max((a.x1 - a.x0) / 100, 0.01)
dpi = min(600, int(a.width / (pw * frac) * 72))
scale = dpi / 72
X, Y = int(pw * scale * a.x0 / 100), int(ph * scale * a.y0 / 100)
W, H = int(pw * scale * (a.x1 - a.x0) / 100), int(ph * scale * (a.y1 - a.y0) / 100)
out = a.out or str(Path(a.pdf).with_name(f"crop-p{a.page}-{int(a.x0)}-{int(a.y0)}-{int(a.x1)}-{int(a.y1)}"))
out = out[:-4] if out.endswith(".png") else out
subprocess.run(["pdftoppm", "-f", str(a.page), "-l", str(a.page), "-png", "-singlefile", "-r", str(dpi),
                "-x", str(X), "-y", str(Y), "-W", str(W), "-H", str(H), a.pdf, out], check=True)
print(f"{out}.png  ({W}x{H}px @ {dpi}dpi)")
