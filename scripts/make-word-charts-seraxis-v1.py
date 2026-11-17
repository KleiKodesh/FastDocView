# Makes test-data/word-charts-binary/word-charts-seraxis-v1.docx from the Word-made word-charts-binary-v1.docx by rewriting its
# chart XML directly (no Excel automation):
#   3-D columns (standard grouping: one depth row per series), 3-D lines and areas, each with a shown series axis (c:serAx, with
#   the tick-label text properties Word writes for its own axes), right-angled and with perspective; and two charts in colours
#   outside the BIFF8 56-colour palette (series fills, title / axis / legend text) - the .doc / .rtf copies keep those exactly
#   only in their DrawingML extension records.
# Then scripts/save-word-charts-seraxis.ps1 lets Word save it as PDF, .doc and .rtf.
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
GREY = '<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'

def rgb(value):
    return '<a:srgbClr val="%s"/>' % value

def colour_series(x, fills):
    for s, fill in zip(sers(x), fills):
        sp = re.search(r'<c:spPr>.*?</c:spPr>', s, re.S).group(0)
        new = re.sub(r'<a:solidFill>.*?</a:solidFill>', '<a:solidFill>%s</a:solidFill>' % rgb(fill), sp, count=1, flags=re.S)
        x = x.replace(s, s.replace(sp, new))
    return x

def make(kind, title, view3d, label_color=None, fills=None, title_color=None, legend_color=None):
    """kind: col3d (standard grouping), line3d, area3d - each with a shown c:serAx; col for a 2-D column chart."""
    base = {'col': 1, 'col3d': 7, 'line3d': 3, 'area3d': 5}[kind]
    x = charts[base]
    old_group = re.search(r'<c:(barChart|bar3DChart|lineChart|areaChart)>.*?</c:\1>', x, re.S).group(0)
    ids = re.findall(r'<c:axId val="(\d+)"/>', old_group)
    series = ''.join(sers(x))
    axes = '<c:axId val="%s"/><c:axId val="%s"/>' % (ids[0], ids[1])
    if kind != 'col': axes += '<c:axId val="1999000001"/>'
    if kind == 'line3d': group = '<c:line3DChart><c:grouping val="standard"/><c:varyColors val="0"/>' + series + LABELS + '<c:gapDepth val="150"/>' + axes + '</c:line3DChart>'
    elif kind == 'area3d': group = '<c:area3DChart><c:grouping val="standard"/><c:varyColors val="0"/>' + series + LABELS + '<c:gapDepth val="150"/>' + axes + '</c:area3DChart>'
    elif kind == 'col3d': group = re.sub(r'<c:axId val="\d+"/>', '', old_group.replace('<c:grouping val="clustered"/>', '<c:grouping val="standard"/>')).replace('</c:bar3DChart>', axes + '</c:bar3DChart>')
    else: group = old_group
    x = x.replace(old_group, group)
    if kind != 'col':
        cat = block(x, 'catAx')
        text = re.search(r'<c:txPr>.*?</c:txPr>', cat, re.S).group(0)
        ser_axis = ('<c:serAx><c:axId val="1999000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>'
                    '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
                    '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>%s<c:crossAx val="%s"/><c:crosses val="autoZero"/></c:serAx>'
                    % (text, ids[1]))
        x = x.replace('</c:valAx>', '</c:valAx>' + ser_axis, 1)
        view = '<c:view3D><c:rotX val="%d"/><c:rotY val="%d"/><c:depthPercent val="100"/><c:rAngAx val="%d"/>%s</c:view3D>' % (
            view3d['rotX'], view3d['rotY'], view3d['rAngAx'], '<c:perspective val="%d"/>' % view3d['perspective'] if 'perspective' in view3d else '')
        x = re.sub(r'<c:view3D>.*?</c:backWall>', '', x, flags=re.S)
        x = x.replace('<c:plotArea>', view + WALLS + '<c:plotArea>', 1)
    if fills: x = colour_series(x, fills)
    if label_color:
        for name in ('catAx', 'valAx', 'serAx'):
            if '<c:%s>' % name in x:
                b = block(x, name)
                x = x.replace(b, b.replace(GREY, rgb(label_color)))
    if title_color:
        b = block(x, 'title')
        x = x.replace(b, b.replace(GREY, rgb(title_color)))
    if legend_color:
        b = block(x, 'legend')
        x = x.replace(b, b.replace(GREY, rgb(legend_color)))
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
write('word-charts-seraxis-v1.docx', {
    1: make('col3d', 'Col3D SerAx R', RIGHT),
    2: make('col3d', 'Col3D SerAx P', PERSPECTIVE),
    3: make('line3d', 'Line3D SerAx R', RIGHT),
    4: make('area3d', 'Area3D SerAx P', PERSPECTIVE),
    5: make('area3d', 'Area3D SerAx R', {'rotX': 20, 'rotY': 30, 'rAngAx': 1}),
    6: make('col', 'Col Colors', None, label_color='8A4F1C', fills=['12AB9C', 'C8641E', '5B2C83'], title_color='7A2E9B', legend_color='2F6EA8'),
    7: make('col3d', 'Col3D Colors', RIGHT, label_color='1E7A6B', fills=['E07B91', '3D9970', 'B8860B'], title_color='C0392B'),
})
