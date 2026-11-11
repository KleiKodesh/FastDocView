"""Compares stretched operator widths in the live rig (vue-demo/playwright/mathStretchCheck.mjs) with Word's PDF of the same
brackets file: per "case <label>:" line the advance width of Word's first operator glyph (a size variant, or the widest
assembly part). python scripts/compare-math-stretch.py <word.pdf> <ours.json> [tolerance=0.15]"""
import json, sys
import fitz
pdf, ours = sys.argv[1], json.load(open(sys.argv[2], encoding="utf-8"))
tolerance = float(sys.argv[3]) if len(sys.argv) > 3 else 0.15
word = {}
for page in fitz.open(pdf):
    chars = [(c["c"], c["bbox"]) for b in page.get_text("rawdict")["blocks"] for l in b.get("lines", []) for s in l["spans"] for c in s["chars"]]
    text = "".join(c[0] for c in chars)
    at = 0
    while (i := text.find("case ", at)) >= 0:
        colon = text.find(":", i); end = text.find(" end", colon); label = text[i + 5:colon]
        glyphs = [c for c in chars[colon + 1:end] if c[0] in "([{|\u222b" or not (c[0].isalnum() or c[0] in " :" or ord(c[0]) > 0x1d000)]
        if glyphs: word[label] = round(glyphs[0][1][2] - glyphs[0][1][0], 1)
        at = end
bad = 0
for label, w in word.items():
    o = ours.get(label)
    if not o: print(f"{label:12} word {w:5.1f}  ours -"); bad += 1; continue
    ratio = o[0] / w
    out = abs(ratio - 1) > tolerance
    bad += out
    print(f"{label:12} word {w:5.1f}  ours {o[0]:5.1f}  {ratio:4.2f}{'  OUT' if out else ''}")
print(f"{bad} of {len(word)} outside +-{tolerance:.0%}")
