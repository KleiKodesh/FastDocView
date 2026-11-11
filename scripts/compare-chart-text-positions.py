# Compares chart text positions: Word's PDF (PyMuPDF spans) vs our SVG (<text x y> in each <svg class="word-chart">).
# Positions are taken relative to each chart's title, so page placement does not matter. Prints ASCII only.
# Usage: python compare-chart-text-positions.py <word.pdf> <rendered.html> <title> [<title> ...]
import re, sys, statistics, fitz

pdf, html_path, titles = sys.argv[1], sys.argv[2], sys.argv[3:]
spans = []
for page in fitz.open(pdf):
    for block in page.get_text('dict')['blocks']:
        for line in block.get('lines', []):
            for s in line['spans']:
                t = s['text'].strip()
                if t:
                    x0, y0, x1, y1 = s['bbox']
                    spans.append((page.number, t, (x0 + x1) / 2, s['origin'][1]))
html = open(html_path, encoding='utf-8').read()
svgs = re.findall(r'<svg class="word-chart".*?</svg>', html, re.S)
dx, dy, n = [], [], 0
for title in titles:
    svg = next((s for s in svgs if '>' + title + '<' in s), None)
    ref = next((s for s in spans if s[1] == title), None)
    if svg is None or ref is None:
        print('missing', title.encode('ascii', 'replace').decode()); continue
    texts = []
    for m in re.finditer(r'<text([^>]*)>([^<]*)</text>', svg):
        a, t = m.group(1), m.group(2).strip()
        x = re.search(r' x="([-0-9.]+)"', a); y = re.search(r' y="([-0-9.]+)"', a)
        if not x or not y or not t: continue
        anchor = re.search(r'text-anchor="(\w+)"', a)
        texts.append((t, float(x.group(1)), float(y.group(1)), anchor.group(1) if anchor else 'start'))
    title_svg = next(t for t in texts if t[0] == title)
    page = ref[0]
    # spans of the same page within the chart's vertical reach of the title
    near = [s for s in spans if s[0] == page and 0 <= s[3] - ref[3] < 300]
    used = set()
    for t, x, y, anchor in texts:
        if t == title or anchor != 'middle': continue
        cands = [s for s in near if s[1] == t and id(s) not in used]
        if not cands: continue
        rx, ry = x - title_svg[1], y - title_svg[2]
        best = min(cands, key=lambda s: abs((s[2] - ref[2]) - rx) + abs((s[3] - ref[3]) - ry))
        used.add(id(best))
        dx.append(abs((best[2] - ref[2]) - rx)); dy.append(abs((best[3] - ref[3]) - ry)); n += 1
print('texts', n)
if n:
    print('median dx %.2f pt  dy %.2f pt  max dx %.1f dy %.1f' % (statistics.median(dx), statistics.median(dy), max(dx), max(dy)))
