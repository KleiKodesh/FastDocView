"""Compares equation sizes in the live rig (vue-demo/playwright/mathGeometry.mjs output) with Word's own PDF of the same file.
In the PDF each "case <label>: <equation> end" line gives the equation's box: from the end of "case <label>:" plus one
space to the start of "end" minus one space (width), and the union of the glyph boxes between them (height; its baseline is
the line's text baseline). Prints per equation: width / height in points (Word, ours, ratio) and the baseline offset, and
the count outside the tolerance.  python scripts/compare-math-geometry.py <word.pdf> <ours.json> [tolerance=0.25]"""
import json, sys
import fitz

pdf, ours_path = sys.argv[1], sys.argv[2]
tolerance = float(sys.argv[3]) if len(sys.argv) > 3 else 0.25
ours = json.load(open(ours_path, encoding="utf-8"))
word = {}
for page in fitz.open(pdf):
    chars = []
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                for c in span["chars"]:
                    chars.append((c["c"], fitz.Rect(c["bbox"]), span["origin"][1], span["size"]))
    text = "".join(c[0] for c in chars)
    start = 0
    while True:
        at = text.find("case ", start)
        if at < 0: break
        colon = text.find(":", at)
        label = text[at + 5:colon]
        end = text.find(" end", colon)
        start = colon
        if end < 0 or label in word: continue
        first, last = chars[colon], chars[end]
        line_y = chars[at][2]
        glyphs = [c for c in chars[colon + 1:end] if c[0].strip() and abs(c[2] - line_y) < 40]
        if not glyphs: continue
        space = chars[colon + 1][1].width if chars[colon + 1][0] == " " else 0
        left, right = first[1].x1 + space, chars[end][1].x0
        top = min(g[1].y0 for g in glyphs); bottom = max(g[1].y1 for g in glyphs)
        word[label] = {"width": right - left, "height": bottom - top, "baselineFromTop": line_y - top}

outside = 0
print(f"{'equation':16} {'W word':>7} {'W ours':>7} {'ratio':>6} {'H word':>7} {'H ours':>7} {'ratio':>6} {'base w/o':>9}")
for label, w in word.items():
    o = ours.get(label)
    if not o: continue
    rw, rh = o["width"] / w["width"], o["height"] / w["height"]
    bad = abs(rw - 1) > tolerance or abs(rh - 1) > tolerance
    outside += bad
    print(f"{label:16} {w['width']:7.1f} {o['width']:7.1f} {rw:6.2f} {w['height']:7.1f} {o['height']:7.1f} {rh:6.2f} {w['baselineFromTop']:4.1f}/{o['baselineFromTop']:4.1f}{'  OUT' if bad else ''}")
print(f"{outside} of {len([l for l in word if l in ours])} outside +-{tolerance:.0%}")
