"""Copies each test-data/smartart-layouts/*-v1.docx to *-v1-nodrawing.docx without its SmartArt drawing parts (the
diagramDrawing parts, their relationships and content-type overrides), so the reader has to lay the diagram out from its
data model and layout definition, as Word does for a file another program wrote without the drawing."""
import pathlib
import re
import zipfile

folder = pathlib.Path(__file__).resolve().parent.parent / 'test-data' / 'smartart-layouts'
DRAWING_TYPE = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing'
for source in sorted(folder.glob('*-v1.docx')):
    target = source.with_name(source.stem + '-nodrawing.docx')
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as zout:
        drawings = set()
        for name in zin.namelist():
            if name.endswith('.rels'):
                for match in re.finditer(r'<Relationship [^>]*Type="' + re.escape(DRAWING_TYPE) + r'"[^>]*/>', zin.read(name).decode('utf-8')):
                    target_part = re.search(r'Target="([^"]+)"', match.group(0)).group(1)
                    drawings.add(target_part.split('/')[-1])
        for item in zin.infolist():
            name = item.filename
            if name.split('/')[-1] in drawings or (name.endswith('.rels') and name.split('/')[-1][:-5] in drawings):
                continue
            data = zin.read(name)
            if name.endswith('.rels'):
                data = re.sub(rb'<Relationship [^>]*Type="' + re.escape(DRAWING_TYPE).encode() + rb'"[^>]*/>', b'', data)
            if name == '[Content_Types].xml':
                data = re.sub(rb'<Override [^>]*diagramDrawing[^>]*/>', b'', data)
            zout.writestr(item, data)
    print('stripped', target.name, len(drawings), 'drawing part(s)')
