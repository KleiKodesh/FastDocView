# 3-D scene series probes (2026-11-24): wall-box line3D charts with 1 / 2 / 3 / 5 series (one depth row each), perspective 30 / 60,
# rotX 15 rotY 20 and rotX 30 rotY 40, depthPercent 100 / 200 - how Word's perspective depends on the scene depth.
# .docx written directly; print: save-word-charts-scene-grid.ps1 -Filter 'series-*.docx'; measure: measure-word-chart-perspective-walls.py
# <folder> <out.json> series.
import sys, os, json, zipfile, itertools
import importlib.util
spec = importlib.util.spec_from_file_location('mk', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'make-word-charts-perspective-v1.py'))
src = open(spec.origin).read().split('BASE = dict')[0]
ns = {'__file__': spec.origin}; exec(src, ns); chart = ns['chart']
out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'test-data', 'word-charts-scene-series'); os.makedirs(out, exist_ok=True)
P = []
def add(**k):
    k.setdefault('size', (415, 242)); k.setdefault('h', None)
    k['name'] = 's%d_x%d_y%d_p%d_d%d' % (k['series'], k['rotX'], k['rotY'], k['perspective'], k['depth'])
    P.append(k)
for n in (1, 2, 3, 5):
    for p in (30, 60):
        for x, y in ((15, 20), (30, 40)):
            for d in (100, 200): add(series=n, rotX=x, rotY=y, perspective=p, depth=d)
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
for n in range(0, len(P), N): doc(P[n:n+N], os.path.join(out, 'series-%03d.docx' % (n // N)))
json.dump(P, open(os.path.join(out, 'series.json'), 'w'), indent=0)
print(len(P), 'probes', (len(P)+N-1)//N, 'docs')
