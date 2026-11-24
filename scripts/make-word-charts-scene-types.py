# 3-D chart scene probes per chart type (2026-11-24 round 3): wall-box charts (floor red, back wall green, side wall blue,
# thickness 0, data hidden) for 3-D columns / bars (clustered, stacked, standard), areas (standard, stacked), surface and lines, a
# small rotX x rotY x perspective grid each, plus a right-angled (rAngAx 1) grid of lines and clustered columns at rotX 0-45.
# The .docx is written directly (never AddChart / Excel). Print: save-word-charts-scene-grid.ps1 -Filter 'types-*.docx';
# measure: measure-word-chart-perspective-walls.py <folder> <out.json> types; fit: word_chart_scene_types_fit.py.
import sys, os, json, zipfile
import importlib.util
here = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('grid', os.path.join(here, 'make-word-charts-scene-grid.py'))
gsrc = open(spec.origin).read()
ns = {'__file__': spec.origin, '__name__': 'grid'}
# reuse the grid script's chart() and doc() without running its probe list
exec(gsrc.split('P = []')[0], ns)
exec('W = ' + gsrc.split(chr(10) + 'W = ', 1)[1].split('N = 24')[0], ns)
base_chart = ns['chart']
out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, '..', 'test-data', 'word-charts-scene-types'); os.makedirs(out, exist_ok=True)

def chart(rotX, rotY, perspective, depth, h, series=1, kind='line3D', grouping='standard', barDir=None, rAngAx=0):
    x = base_chart(rotX, rotY, perspective, depth, h, kind='line3D', series=series)
    if rAngAx: x = x.replace('<c:rAngAx val="0"/>', '<c:rAngAx val="1"/>')
    three = grouping == 'standard' or kind == 'surface3D'
    axids = '<c:axId val="11"/><c:axId val="12"/>' + ('<c:axId val="13"/>' if three else '')
    lbl_start = x.index('<c:dLbls>'); lbl_end = x.index('</c:dLbls>') + len('</c:dLbls>')
    sers = x[x.index('<c:ser>'):lbl_start]; lbls = x[lbl_start:lbl_end]
    old = x[x.index('<c:line3DChart>'):x.index('</c:line3DChart>') + len('</c:line3DChart>')]
    if kind == 'bar3D':
        group = ('<c:bar3DChart><c:barDir val="%s"/><c:grouping val="%s"/><c:varyColors val="0"/>%s%s<c:gapWidth val="150"/>%s'
                 '<c:gapDepth val="150"/><c:shape val="box"/>%s</c:bar3DChart>') % (barDir, grouping, sers, lbls,
                 '<c:overlap val="100"/>' if grouping in ('stacked', 'percentStacked') else '', axids)
    elif kind == 'area3D':
        group = '<c:area3DChart><c:grouping val="%s"/><c:varyColors val="0"/>%s%s<c:gapDepth val="150"/>%s</c:area3DChart>' % (grouping, sers, lbls, axids)
    elif kind == 'surface3D':
        group = '<c:surface3DChart><c:wireframe val="0"/>%s<c:bandFmts/>%s</c:surface3DChart>' % (sers, axids)
    else:
        group = old if three else old.replace('<c:axId val="13"/>', '')
    x = x.replace(old, group)
    if not three:
        x = x[:x.index('<c:serAx>')] + x[x.index('</c:serAx>') + len('</c:serAx>'):]
    if kind == 'bar3D' and barDir == 'bar':
        x = x.replace('<c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="12"/><c:crosses val="autoZero"/><c:auto',
                      '<c:axPos val="l"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="12"/><c:crosses val="autoZero"/><c:auto', 1)
        x = x.replace('<c:axPos val="l"/><c:numFmt', '<c:axPos val="b"/><c:numFmt', 1)
    return x

TYPES = [('col-clustered', dict(kind='bar3D', barDir='col', grouping='clustered', series=1)),
         ('col-stacked', dict(kind='bar3D', barDir='col', grouping='stacked', series=2)),
         ('col-standard', dict(kind='bar3D', barDir='col', grouping='standard', series=2)),
         ('area-standard', dict(kind='area3D', grouping='standard', series=2)),
         ('area-stacked', dict(kind='area3D', grouping='stacked', series=2)),
         ('surface', dict(kind='surface3D', series=2)),
         ('line-standard', dict(kind='line3D', grouping='standard', series=2))]
P = []
def add(t, kw, **k):
    k.setdefault('size', (415, 242)); k.setdefault('depth', 100); k.setdefault('h', None); k.setdefault('rAngAx', 0)
    k.update(kw); k['type'] = t
    k['name'] = '%s_x%d_y%d_p%d_r%d' % (t, k['rotX'], k['rotY'], k['perspective'], k['rAngAx'])
    P.append(k)
for t, kw in TYPES:
    for x in (10, 30, 50):
        for y in (20, 50, 80):
            for p in (30, 120): add(t, kw, rotX=x, rotY=y, perspective=p)
    for x, y in ((15, 20), (40, 60)): add(t, kw, rotX=x, rotY=y, perspective=0)
bar = dict(kind='bar3D', barDir='bar', grouping='clustered', series=1)
for x in (10, 30):
    for y in (20, 40):
        for p in (30, 120): add('bar-clustered', bar, rotX=x, rotY=y, perspective=p)
# right-angled axes: lines (one row) and clustered columns, rotX 0-45 by 5 x rotY 0-60 by 15
for x in range(0, 46, 5):
    for y in (0, 15, 30, 45, 60): add('line-ra', dict(kind='line3D', grouping='standard', series=1), rotX=x, rotY=y, perspective=30, rAngAx=1)
for x in range(0, 46, 10):
    for y in (0, 20, 40, 60): add('col-ra', dict(kind='bar3D', barDir='col', grouping='clustered', series=1), rotX=x, rotY=y, perspective=30, rAngAx=1)
ns['chart'] = lambda rotX, rotY, perspective, depth, h, series=1: None
def doc(probes, fn):
    # the grid script's doc() calls chart(p[...]) by position; give it a per-probe closure
    it = iter(probes)
    ns['chart'] = lambda *a, **k: (lambda p: chart(p['rotX'], p['rotY'], p['perspective'], p['depth'], p['h'], series=p['series'], kind=p['kind'],
                                                  grouping=p.get('grouping', 'standard'), barDir=p.get('barDir'), rAngAx=p['rAngAx']))(next(it))
    ns['doc'](probes, fn)
N = 24
for n in range(0, len(P), N): doc(P[n:n+N], os.path.join(out, 'types-%03d.docx' % (n // N)))
json.dump(P, open(os.path.join(out, 'types.json'), 'w'), indent=0)
print(len(P), 'probes', (len(P)+N-1)//N, 'docs')
