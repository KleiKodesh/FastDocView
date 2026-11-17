# Writes picture-bullet-source-v1.docx: a two-item list whose level 0 bullet is a picture (w:numPicBullet with a VML
# v:imagedata, ECMA-376 17.9.21; the picture is ..\word-picture-crop\quadrants.png). make-picture-bullet-fixture.ps1
# then has Word open it and save it as .docx, Word 2003 XML and PDF (Word's own files are the fixtures).
import os, zipfile
here = os.path.dirname(os.path.abspath(__file__))
png = open(os.path.join(here, '..', 'word-picture-crop', 'quadrants.png'), 'rb').read()
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
item = lambda text: f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'
parts = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/_rels/document.xml.rels': '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
    'word/_rels/numbering.xml.rels': '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdBullet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/bullet.png"/></Relationships>',
    'word/numbering.xml': f'<?xml version="1.0" encoding="UTF-8"?><w:numbering {W}><w:numPicBullet w:numPicBulletId="0"><w:pict><v:shape id="b0" type="#_x0000_t75" style="width:12pt;height:6pt" o:bullet="t"><v:imagedata r:id="rIdBullet" o:title=""/></v:shape></w:pict></w:numPicBullet><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlPicBulletId w:val="0"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
    'word/document.xml': f'<?xml version="1.0" encoding="UTF-8"?><w:document {W}><w:body>{item("Picture bullet one")}{item("Picture bullet two")}<w:sectPr/></w:body></w:document>',
}
with zipfile.ZipFile(os.path.join(here, 'picture-bullet-source-v1.docx'), 'w', zipfile.ZIP_DEFLATED) as z:
    for name, text in parts.items(): z.writestr(name, text.encode('utf-8'))
    z.writestr('word/media/bullet.png', png)
