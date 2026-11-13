# Makes two fixtures from the Word-made test-data/word-charts-binary/word-charts-binary-v1.docx by rewriting its chart XML
# directly (no Excel automation):
#   word-charts-3d-rows-v1.docx  - 3-D line and area charts: perspective and right-angled axes, a deep scene, a shown series axis
#   word-charts-crosses-v1.docx  - c:crosses min / max and c:crossesAt on the category and value axes (columns, bars, lines,
#                                  areas, 3-D columns, a scatter chart), with negative values
# Then scripts/save-word-charts-3d-crosses.ps1 lets Word save each as PDF and .doc.
import re, zipfile, os
folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'test-data', 'word-charts-binary')
source = os.path.join(folder, 'word-charts-binary-v1.docx')

with zipfile.ZipFile(source) as z:
    charts = {i: z.read('word/charts/chart%d.xml' % i).decode('utf-8') for i in range(1, 8)}

def sers(x):
    return re.findall(r'<c:ser>.*?</c:ser>', x, re.S)

def block(x, name):
    return re.search(r'<c:%s>.*?</c:%s>' % (name, name), x, re.S).group(0)

WALLS = re.search(r'<c:floor>.*?</c:backWall>', charts[7], re.S).group(0)
LABELS = ('<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>'
          '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')

def with_values(series, values):
    if values is None: return series
    out = []
    for s, vs in zip(series, values):
        val = block(s, 'val')
        cache = re.sub(r'<c:pt idx="(\d)"><c:v>[^<]*</c:v></c:pt>', lambda m: '<c:pt idx="%s"><c:v>%s</c:v></c:pt>' % (m.group(1), vs[int(m.group(1))]), val)
        out.append(s.replace(val, cache))
    return out

def crossing(axis, crosses):
    if crosses is None: return axis
    if crosses.startswith('at:'):
        return axis.replace('<c:crosses val="autoZero"/>', '<c:crossesAt val="%s"/>' % crosses[3:])
    return axis.replace('<c:crosses val="autoZero"/>', '<c:crosses val="%s"/>' % crosses)

def make(kind, title, view3d=None, gap_depth=150, series_axis_shown=False, values=None, cat_crosses=None, val_crosses=None):
    """kind: col, bar, line, area, col3d, line3d, area3d. Built on the matching v1 chart (series formatting, axes)."""
    base = {'col': 1, 'bar': 2, 'line': 3, 'area': 5, 'col3d': 7, 'line3d': 3, 'area3d': 5}[kind]
    x = charts[base]
    old_group = re.search(r'<c:(barChart|bar3DChart|lineChart|areaChart)>.*?</c:\1>', x, re.S).group(0)
    ids = re.findall(r'<c:axId val="(\d+)"/>', old_group)
    series = ''.join(with_values(sers(x), values))
    axes = ''.join('<c:axId val="%s"/>' % i for i in ids)
    three_d = kind.endswith('3d')
    if three_d and kind != 'col3d': axes += '<c:axId val="1999000001"/>'
    if kind == 'line3d': group = '<c:line3DChart><c:grouping val="standard"/><c:varyColors val="0"/>' + series + LABELS + '<c:gapDepth val="%d"/>' % gap_depth + axes + '</c:line3DChart>'
    elif kind == 'area3d': group = '<c:area3DChart><c:grouping val="standard"/><c:varyColors val="0"/>' + series + LABELS + '<c:gapDepth val="%d"/>' % gap_depth + axes + '</c:area3DChart>'
    else: group = old_group.replace(''.join(sers(x)), series)   # the series sit together in the group
    x = x.replace(old_group, group)
    cat, val = block(x, 'catAx'), block(x, 'valAx')
    x = x.replace(cat, crossing(cat, cat_crosses)).replace(block(x, 'valAx'), crossing(val, val_crosses))
    if three_d and kind != 'col3d':
        ser_axis = ('<c:serAx><c:axId val="1999000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="%d"/><c:axPos val="b"/>'
                    '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="%s"/><c:crosses val="autoZero"/></c:serAx>'
                    % (0 if series_axis_shown else 1, ids[1]))
        x = x.replace('</c:valAx>', '</c:valAx>' + ser_axis, 1)
    if three_d:
        view = '<c:view3D><c:rotX val="%d"/><c:rotY val="%d"/><c:depthPercent val="%d"/><c:rAngAx val="%d"/>%s</c:view3D>' % (
            view3d['rotX'], view3d['rotY'], view3d.get('depth', 100), view3d['rAngAx'], '<c:perspective val="%d"/>' % view3d['perspective'] if 'perspective' in view3d else '')
        x = re.sub(r'<c:view3D>.*?</c:backWall>', '', x, flags=re.S)
        x = x.replace('<c:plotArea>', view + WALLS + '<c:plotArea>', 1)
    x = re.sub(r'(<a:t>)[^<]*Chart(</a:t>)', r'\g<1>' + title + r'\g<2>', x, count=1)
    return x

def write(target, replaced):
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(os.path.join(folder, target), 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.match(r'word/charts/chart(\d)\.xml$', item.filename)
            if m and int(m.group(1)) in replaced: data = replaced[int(m.group(1))].encode('utf-8')
            zout.writestr(item, data)
    print('written', target)

PERSPECTIVE = {'rotX': 15, 'rotY': 20, 'rAngAx': 0, 'perspective': 30}
RIGHT = {'rotX': 15, 'rotY': 20, 'rAngAx': 1}
write('word-charts-3d-rows-v1.docx', {
    1: make('line3d', 'Line3D Persp', PERSPECTIVE),
    2: make('area3d', 'Area3D Persp', PERSPECTIVE),
    3: make('line3d', 'Line3D Right', RIGHT),
    4: make('area3d', 'Area3D Right', RIGHT),
    5: make('line3d', 'Line3D Deep', {'rotX': 30, 'rotY': 40, 'rAngAx': 0, 'perspective': 60, 'depth': 200}, gap_depth=300),
    6: make('area3d', 'Area3D Flat', {'rotX': 10, 'rotY': 10, 'rAngAx': 1, 'depth': 60}, gap_depth=50),
    7: make('line3d', 'Line3D SerAx', PERSPECTIVE, series_axis_shown=True),
})
NEGATIVE = [['4.3', '-2.5', '3.5', '4.5'], ['2.4', '4.4', '-1.8', '2.8'], ['2', '2', '3', '-5']]
write('word-charts-crosses-v1.docx', {
    1: make('col', 'Col CrossMax', values=NEGATIVE, cat_crosses='max'),
    2: make('bar', 'Bar CrossAt', values=NEGATIVE, cat_crosses='at:2', val_crosses='max'),
    3: make('line', 'Line ValMax', values=NEGATIVE, val_crosses='max', cat_crosses='min'),
    4: make('area', 'Area CrossAt', cat_crosses='at:3', val_crosses='at:2'),
    5: make('col', 'Col CrossAt', values=NEGATIVE, cat_crosses='at:-2', val_crosses='at:3'),
    6: make('col3d', 'Col3D CrossAt', {'rotX': 15, 'rotY': 20, 'rAngAx': 1}, values=NEGATIVE, cat_crosses='at:2'),
    7: make('line', 'Line CrossMin', values=NEGATIVE, cat_crosses='min'),
})
