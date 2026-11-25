# Builds east-asian-layout-source-v1.docx: w:eastAsianLayout (two lines in one = combine with each bracket kind, horizontal in
# vertical = vert, vertCompress) and w:fitText on Latin and Han text in ordinary horizontal paragraphs. Word 16 then saves the
# .docx / .doc and the reference PDF (make-east-asian-layout-fixture.ps1). Text is written with \u escapes.
import zipfile

HAN = '日本語文字列'   # six Han characters
FONTS = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="MS Mincho" w:cs="Arial"/><w:sz w:val="28"/><w:szCs w:val="28"/>'

def run(text, extra=''):
    return f'<w:r><w:rPr>{FONTS}{extra}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'

def paragraph(tag, text, extra):
    return f'<w:p>{run(tag + " x")}{run(text, extra)}{run("y")}</w:p>'

cases = [
    ('L1', 'ABCDEF', '<w:eastAsianLayout w:id="1" w:combine="1"/>'),
    ('L2', 'ABCDEF', '<w:eastAsianLayout w:id="2" w:combine="1" w:combineBrackets="round"/>'),
    ('L3', 'ABCDE', '<w:eastAsianLayout w:id="3" w:combine="1" w:combineBrackets="square"/>'),
    ('L4', HAN, '<w:eastAsianLayout w:id="4" w:combine="1" w:combineBrackets="angle"/>'),
    ('L5', HAN, '<w:eastAsianLayout w:id="5" w:combine="1" w:combineBrackets="curly"/>'),
    ('L6', '12', '<w:eastAsianLayout w:id="6" w:vert="1"/>'),
    ('L7', 'ABCD', '<w:eastAsianLayout w:id="7" w:vert="1" w:vertCompress="1"/>'),
    ('L8', 'ABC', '<w:fitText w:val="2880" w:id="8"/>'),
    ('L9', 'ABCDEFGHIJ', '<w:fitText w:val="720" w:id="9"/>'),
    ('L10', HAN[:3], '<w:fitText w:val="2160" w:id="10"/>'),
]
body = ''.join(paragraph(tag, text, extra) for tag, text, extra in cases)
document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body +
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
            '</w:body></w:document>')
content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '</Types>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
with zipfile.ZipFile('east-asian-layout-source-v1.docx', 'w', zipfile.ZIP_DEFLATED) as package:
    package.writestr('[Content_Types].xml', content_types)
    package.writestr('_rels/.rels', rels)
    package.writestr('word/document.xml', document.encode('utf-8'))
