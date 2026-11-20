"""Print, per page, the y (pt from top) of long horizontal rules in a PDF: table borders.
Usage: python -I rules.py file.pdf [min_len_pt]. Renders with Ghostscript at 144 dpi."""
import subprocess, sys, tempfile, os, glob
from PIL import Image

pdf = sys.argv[1]
min_len = float(sys.argv[2]) if len(sys.argv) > 2 else 250
dpi = 144
with tempfile.TemporaryDirectory() as tmp:
    subprocess.run(["gswin64c", "-q", "-dNOPAUSE", "-dBATCH", "-sDEVICE=pnggray", f"-r{dpi}",
                    f"-sOutputFile={tmp}/p%03d.png", pdf], check=True)
    for i, f in enumerate(sorted(glob.glob(os.path.join(tmp, "p*.png"))), 1):
        im = Image.open(f)
        w, h = im.size
        px = im.load()
        ys = []
        for y in range(h):
            run = best = 0
            for x in range(w):
                if px[x, y] < 128:
                    run += 1
                    best = max(best, run)
                else:
                    run = 0
            if best * 72 / dpi >= min_len:
                ys.append(y)
        # merge adjacent rows
        merged = []
        for y in ys:
            if merged and y - merged[-1][1] <= 2:
                merged[-1][1] = y
            else:
                merged.append([y, y])
        print(f"page {i}: " + ", ".join(f"{(a + b) / 2 * 72 / dpi:.1f}" for a, b in merged))
