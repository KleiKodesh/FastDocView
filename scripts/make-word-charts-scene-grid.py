# Dense 3-D chart scene grid (616 wall-box line3D probes: rotX 0-80 by 5, rotY 0-90 by 10, perspective 30/60/120, plus
# perspective 0, three plot aspects, depth / height 200), .docx written directly - never AddChart / Excel.
# Print: scripts/save-word-charts-scene-grid.ps1; measure: measure-word-chart-perspective-walls.py (grid.json / grid-*.pdf);
# fit: word_chart_scene_grid_fit.py.
import sys, os, json, zipfile, itertools
import importlib.util
spec = importlib.util.spec_from_file_location('mk', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'make-word-charts-perspective-v1.py'))
src = open(spec.origin).read().split('BASE = dict')[0]
ns = {'__file__': spec.origin}; exec(src, ns); chart = ns['chart']
out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'test-data', 'word-charts-scene-grid'); os.makedirs(out, exist_ok=True)
P = []
def add(**k):
    k.setdefault('size', (415, 242)); k.setdefault('depth', 100); k.setdefault('h', None); k.setdefault('series', 1)
    k['name'] = 'x%d_y%d_p%d_%dx%d_d%d_h%s' % (k['rotX'], k['rotY'], k['perspective'], k['size'][0], k['size'][1], k['depth'], k['h'])
    P.append(k)
for x in range(0, 81, 5):
    for y in range(0, 91, 10):
        for p in (30, 60, 120): add(rotX=x, rotY=y, perspective=p)
for x in range(0, 81, 10):
    for y in (0, 30, 60, 90): add(rotX=x, rotY=y, perspective=0)
for sz in ((300, 300), (415, 150), (250, 350)):
    for x in (10, 30, 50):
        for y in (20, 50, 80):
            for p in (60, 120): add(rotX=x, rotY=y, perspective=p, size=sz)
for x in (15, 40):
    for y in (20, 60):
        for p in (60, 120):
            add(rotX=x, rotY=y, perspective=p, depth=200); add(rotX=x, rotY=y, perspective=p, h=200)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def doc(probes, fn):
    paras = []
    rels = ['<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>']
    ct = []
    for i, p in enumerate(probes, 1):
        w, h = p['size']
        paras.append('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="c%d"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdC%d"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' % (w*12700, h*12700, i, i, i))
        rels.append('<Relationship Id="rIdC%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart%d.xml"/>' % (i, i))
        ct.append('<Override PartName="/word/charts/chart%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' % i)
    body = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body>%s<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>') % (W, ''.join(paras))
    with zipfile.ZipFile(fn, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' + ''.join(ct) + '</Types>')
        z.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(rels) + '</Relationships>')
        z.writestr('word/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="%s"/>' % W)
        z.writestr('word/document.xml', body)
        for i, p in enumerate(probes, 1):
            z.writestr('word/charts/chart%d.xml' % i, chart(p['rotX'], p['rotY'], p['perspective'], p['depth'], p['h'], series=p['series']))
N = 24
for n in range(0, len(P), N): doc(P[n:n+N], os.path.join(out, 'grid-%03d.docx' % (n // N)))
json.dump(P, open(os.path.join(out, 'grid.json'), 'w'), indent=0)
print(len(P), 'probes', (len(P)+N-1)//N, 'docs')
