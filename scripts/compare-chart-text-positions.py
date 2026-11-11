# Compares chart text positions: Word's PDF (PyMuPDF spans) vs our SVG (<text x y> in each <svg class="word-chart">).
# Positions are relative to each chart's title, so page placement does not matter. x is compared at the text's anchor:
# its centre for text-anchor middle, its left edge for start, its right edge for end (the PDF span's bbox edges).
# Prints ASCII only (non-ASCII texts are escaped).
# Usage: python compare-chart-text-positions.py <word.pdf> <rendered.html> [--detail] <title> [<title> ...]
import re, sys, statistics, fitz

args = sys.argv[1:]
detail = '--detail' in args
args = [a for a in args if a != '--detail']
pdf, html_path, titles = args[0], args[1], args[2:]
spans = []
for page in fitz.open(pdf):
    for block in page.get_text('dict')['blocks']:
        for line in block.get('lines', []):
            for s in line['spans']:
                t = s['text'].strip()
                if t:
                    x0, y0, x1, y1 = s['bbox']
                    spans.append((page.number, t, x0, (x0 + x1) / 2, x1, s['origin'][1]))
html = open(html_path, encoding='utf-8').read()
svgs = re.findall(r'<svg class="word-chart".*?</svg>', html, re.S)
dx, dy, n = [], [], 0
ascii_ = lambda t: t.encode('ascii', 'backslashreplace').decode()
for title in titles:
    svg = next((s for s in svgs if '>' + title + '<' in s), None)
    ref = next((s for s in spans if s[1] == title), None)
    if svg is None or ref is None:
        print('missing', ascii_(title)); continue
    texts = []
    for m in re.finditer(r'<text([^>]*)>([^<]*)</text>', svg):
        a, t = m.group(1), m.group(2).strip()
        x = re.search(r' x="([-0-9.]+)"', a); y = re.search(r' y="([-0-9.]+)"', a)
        if not x or not y or not t or 'rotate(' in a: continue
        anchor = re.search(r'text-anchor="(\w+)"', a)
        texts.append((t, float(x.group(1)), float(y.group(1)), anchor.group(1) if anchor else 'start'))
    title_svg = next(t for t in texts if t[0] == title)
    page = ref[0]
    # only the spans inside this chart's own box (the SVG height below the title's baseline, its width around the title)
    box = re.search(r'viewBox="0 0 ([0-9.]+) ([0-9.]+)"', svg)
    width, height = (float(box.group(1)), float(box.group(2))) if box else (600.0, 300.0)
    near = [s for s in spans if s[0] == page and -title_svg[2] - 2 <= s[5] - ref[5] <= height - title_svg[2] + 2
            and abs(s[3] - ref[3]) <= width / 2 + 2]
    used = set()
    for t, x, y, anchor in texts:
        if t == title: continue
        cands = [s for s in near if s[1] == t and id(s) not in used]
        if not cands: continue
        k = {'start': 2, 'middle': 3, 'end': 4}[anchor]
        rx, ry = x - title_svg[1], y - title_svg[2]
        best = min(cands, key=lambda s: abs((s[k] - ref[3]) - rx) + abs((s[5] - ref[5]) - ry))
        used.add(id(best))
        ex, ey = (best[k] - ref[3]) - rx, (best[5] - ref[5]) - ry
        dx.append(abs(ex)); dy.append(abs(ey)); n += 1
        if detail: print('  %-18s %-6s word-ours dx %6.1f dy %6.1f' % (ascii_(title)[:18], ascii_(t)[:6], ex, ey))
print('texts', n)
if n:
    print('median dx %.2f pt  dy %.2f pt  max dx %.1f dy %.1f' % (statistics.median(dx), statistics.median(dy), max(dx), max(dy)))
