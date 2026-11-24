# Renderer vs Word, 3-D scene walls: renders every probe .docx of a measured probe set (measure-word-chart-perspective-walls.py
# output: Word's floor / back / side wall corners, by colour red / green / blue) with the runner's --chart-html and compares our
# wall polygons (matched by the same fill colours) corner by corner, in points on the page (chart origin = the PDF's chart frame).
# Faces thinner than 12 pt are compared one way (measured corners to our outline), as word_chart_camera_fit.residuals does.
# Usage: python chart3d_walls_vs_word.py <runner exe> <probe folder> <measured.json> <out folder> [name filter]
# Prints per probe rms / max (pt) and per group (perspective / right-angled / parallel, chart type) medians. ASCII only.
import sys, os, re, json, glob, subprocess, math, numpy as np
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from word_chart_camera_fit import match, thickness, outline_distance
exe, folder, measured, out = sys.argv[1:5]
flt = sys.argv[5] if len(sys.argv) > 5 else ''
os.makedirs(out, exist_ok=True)
P = json.load(open(measured))
COLOURS = {'#FF0000': 'floor', '#00FF00': 'back', '#0000FF': 'side'}
svgs = []
docs = sorted(set(f for f in glob.glob(os.path.join(folder, '*.docx'))))
for d in docs:
    html = os.path.join(out, os.path.basename(d) + '.html')
    if not os.path.exists(html) or os.path.getmtime(html) < os.path.getmtime(exe):
        subprocess.run([exe, '--chart-html', d, html], check=True, stdout=subprocess.DEVNULL)
    svgs += re.findall(r'<svg class="word-chart".*?</svg>', open(html, encoding='utf-8').read(), re.S)
rows = []
for p, svg in zip(P, svgs):
    if flt and flt not in p['name']: continue
    if not p.get('faces') or 'chart' not in p: continue
    ox, oy = p['chart'][0], p['chart'][1]
    ours = {}
    for m in re.finditer(r'<path class="word-chart-(?:wall|floor)" d="([^"]+)" fill="(#[0-9A-F]{6})"', svg):
        pts = [tuple(map(float, q)) for q in re.findall(r'[ML]([-0-9.]+) ([-0-9.]+)', m.group(1))]
        if m.group(2) in COLOURS: ours[COLOURS[m.group(2)]] = np.array(pts) + (ox, oy)
    res = []
    for name, q in p['faces'].items():
        q = np.array(q)
        if name not in ours: continue
        o = ours[name]
        if len(o) != 4: continue
        if thickness(q) < 12: res += [outline_distance(b, o) for b in q]
        else: res += list(np.hypot(*(match(o, q) - o).T))
    if not res: continue
    res = np.array(res)
    kind = p.get('type', 'line-grid')
    mode = 'rightAngled' if p.get('rAngAx') else ('parallel' if not p['perspective'] else 'perspective')
    rows.append((p['name'], kind, mode, math.sqrt((res ** 2).mean()), res.max()))
for r in rows: print('%-40s %-14s %-11s rms %6.2f max %6.2f' % r)
def summary(label, sel):
    if not sel: return
    a = np.array([r[3] for r in sel]); b = np.array([r[4] for r in sel])
    print('GROUP %-28s n %3d  rms median %5.2f  p90 %5.2f  max %6.2f  | max median %5.2f  within 1 pt %d' % (label, len(sel), np.median(a), np.percentile(a, 90), a.max(), np.median(b), (a < 1).sum()))
for mode in ('perspective', 'parallel', 'rightAngled'):
    summary(mode, [r for r in rows if r[2] == mode])
    for kind in sorted(set(r[1] for r in rows)):
        summary(mode + ' ' + kind, [r for r in rows if r[2] == mode and r[1] == kind])
summary('ALL', rows)
