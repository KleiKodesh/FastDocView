# Makes test-data/word-charts-binary/word-charts-perspective-v1-<n>.docx: minimal 3-D line charts that vary one camera parameter
# at a time (perspective, rotX, rotY, depthPercent, hPercent, chart size), written from word-charts-binary-v1.docx by replacing the
# chart XML directly (no Excel automation). Every chart has one hidden series, no axes, no title or legend, and solid walls (floor
# red, back wall green, side wall blue, thickness 0) so the projected box corners can be read from Word's scene picture.
# Then scripts/save-word-charts-perspective.ps1 lets Word save each as PDF.
import re, zipfile, os, json
folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'test-data', 'word-charts-binary')
source = os.path.join(folder, 'word-charts-binary-v1.docx')

def wall(tag, color):
    return ('<c:%s><c:thickness val="0"/><c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:%s>'
            % (tag, color, tag))

def chart(rotX=15, rotY=20, perspective=30, depth=100, h=None, kind='line3D', series=1):
    view = '<c:view3D><c:rotX val="%d"/><c:rotY val="%d"/>%s<c:depthPercent val="%d"/><c:rAngAx val="0"/><c:perspective val="%d"/></c:view3D>' % (
        rotX, rotY, '' if h is None else '<c:hPercent val="%d"/>' % h, depth, perspective)
    sers = ''.join('<c:ser><c:idx val="%d"/><c:order val="%d"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:val><c:numLit><c:formatCode>General</c:formatCode>'
                   '<c:ptCount val="4"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>'
                   '<c:pt idx="3"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>' % (i, i) for i in range(series))
    lbls = ('<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>'
            '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
    group = '<c:%sChart><c:grouping val="standard"/><c:varyColors val="0"/>%s%s<c:gapDepth val="150"/><c:axId val="11"/><c:axId val="12"/><c:axId val="13"/></c:%sChart>' % (
        kind, sers, lbls, kind)
    axes = ('<c:catAx><c:axId val="11"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/>'
            '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="12"/><c:crosses val="autoZero"/>'
            '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
            '<c:valAx><c:axId val="12"/><c:scaling><c:orientation val="minMax"/><c:max val="4"/><c:min val="0"/></c:scaling><c:delete val="1"/>'
            '<c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
            '<c:tickLblPos val="nextTo"/><c:crossAx val="11"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>'
            '<c:serAx><c:axId val="13"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/>'
            '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="12"/><c:crosses val="autoZero"/></c:serAx>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:roundedCorners val="0"/>'
            '<c:chart><c:autoTitleDeleted val="1"/>' + view + wall('floor', 'FF0000') + wall('sideWall', '0000FF') + wall('backWall', '00FF00') +
            '<c:plotArea><c:layout/>' + group + axes + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>')

BASE = dict(rotX=15, rotY=20, perspective=30, depth=100)
PROBES = []
def probe(name, size=(415, 242), **kw):
    p = dict(BASE); p.update(kw); p['name'] = name; p['size'] = size; PROBES.append(p)
for v in (0, 15, 30, 60, 90, 120, 240): probe('p%d' % v, perspective=v)
for v in (0, 30, 60, -20): probe('x%d' % v, rotX=v)
for v in (0, 45, 80, 160): probe('y%d' % v, rotY=v)
for v in (50, 200, 500): probe('d%d' % v, depth=v)
for v in (50, 100, 200): probe('h%d' % v, h=v)
probe('s-wide', size=(415, 150)); probe('s-square', size=(300, 300))
for v in (60, 120): probe('p%d-x40y50' % v, perspective=v, rotX=40, rotY=50)
probe('p90-d200', perspective=90, depth=200); probe('p90-h200', perspective=90, h=200); probe('p120-square', perspective=120, size=(300, 300))
probe('p60-s3', perspective=60, series=3)

def write(target, probes):
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(os.path.join(folder, target), 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.match(r'word/charts/chart(\d)\.xml$', item.filename)
            if m:
                i = int(m.group(1)) - 1
                if i < len(probes):
                    p = probes[i]
                    data = chart(p['rotX'], p['rotY'], p['perspective'], p['depth'], p.get('h'), series=p.get('series', 1)).encode('utf-8')
            if item.filename == 'word/document.xml':
                d = data.decode('utf-8'); k = [0]
                def extent(m):
                    i = k[0]; k[0] += 1
                    if i >= len(probes): return m.group(0)
                    w, h = probes[i]['size']
                    return '<wp:extent cx="%d" cy="%d"/>' % (w * 12700, h * 12700)
                data = re.sub(r'<wp:extent cx="\d+" cy="\d+"/>', extent, d).encode('utf-8')
            zout.writestr(item, data)
    print('written', target)

for n in range(0, len(PROBES), 7):
    write('word-charts-perspective-v1-%d.docx' % (n // 7 + 1), PROBES[n:n + 7])
with open(os.path.join(folder, 'word-charts-perspective-v1.json'), 'w') as f:
    json.dump(PROBES, f, indent=0)
