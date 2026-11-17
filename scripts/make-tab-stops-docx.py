# Writes test-data/tab-stops/tab-stops-v1.docx: tab stops of every kind (left / centre / right / decimal / bar), every leader,
# default stops (with and without an indent), a hanging indent, stops from a style and a clear, a right-to-left paragraph and an
# absolute position tab. ASCII text only. scripts/save-tab-stops.ps1 then has Word save it as PDF, .doc and .rtf, and LibreOffice
# converts it to .odt and the .odt to PDF (the references the tab layout is compared with).
import os, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'test-data', 'tab-stops', 'tab-stops-v1.docx')
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def tabs(*stops):
    return '<w:tabs>' + ''.join(f'<w:tab w:val="{v}" w:pos="{p}"' + (f' w:leader="{l}"' if l else '') + '/>' for v, p, l in stops) + '</w:tabs>'


def para(text, ppr='', style=None):
    runs = []
    for i, part in enumerate(text.split('\t')):
        if i:
            runs.append('<w:r><w:tab/></w:r>')
        if part:
            runs.append(f'<w:r><w:t xml:space="preserve">{part}</w:t></w:r>')
    props = (f'<w:pStyle w:val="{style}"/>' if style else '') + ppr
    return f'<w:p><w:pPr>{props}</w:pPr>{"".join(runs)}</w:p>'


body = [
    para('Chapter One\t12', tabs(('right', 9350, 'dot')), 'TOC1'),                                  # 1 style stop, TOC line
    para('A much longer chapter title for the second entry\t345', '', 'TOC1'),                       # 2
    para('Hyphen leader\t7', tabs(('right', 9350, 'hyphen'))),                                       # 3
    para('Underscore leader\t8', tabs(('right', 9350, 'underscore'))),                               # 4
    para('Heavy leader\t9', tabs(('right', 9350, 'heavy'))),                                         # 5
    para('Middle dot leader\t10', tabs(('right', 9350, 'middleDot'))),                               # 6
    para('Left\tCentred text\tRight', tabs(('center', 4680, None), ('right', 9350, None))),          # 7
    para('Price\t123.45', tabs(('decimal', 4000, None))),                                            # 8
    para('Price\t1.5', tabs(('decimal', 4000, None))),                                               # 9
    para('Price\t12345', tabs(('decimal', 4000, None))),                                             # 10
    para('Bar\tafter the bar', tabs(('bar', 2000, None), ('left', 3000, None))),                     # 11
    para('a\tb\tc\td\te'),                                                                           # 12 default stops
    para('in\tdented\ttext', '<w:ind w:left="1000"/>'),                                              # 13 default stops, indent
    para('1.\tHanging indent text after the number', '<w:ind w:left="1440" w:hanging="1440"/>'),     # 14
    para('One\tTwo\tThree', tabs(('left', 1440, None), ('left', 2880, None))),                       # 15
    para('Cleared\tstop', tabs(('clear', 9350, None), ('left', 2000, None)), 'TOC1'),                 # 16
    para('Right to left\t44', '<w:bidi/>' + tabs(('end', 9350, 'dot'))),                             # 17
    para('wide\tWide centred text that is long\tx', tabs(('center', 2000, None))),                   # 18 centre near the start
    para('Long text that goes past the right stop position easily\t1', tabs(('right', 3000, None))), # 19 does not fit
    '<w:p><w:r><w:t>Absolute</w:t></w:r><w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/></w:r><w:r><w:t>99</w:t></w:r></w:p>',  # 20
]

document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {W}><w:body>{"".join(body)}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'''

styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr>{tabs(("right", 9350, "dot"))}</w:pPr></w:style>
</w:styles>'''

settings = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {W}><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''

rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''

document_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', document)
    z.writestr('word/styles.xml', styles)
    z.writestr('word/settings.xml', settings)
    z.writestr('word/_rels/document.xml.rels', document_rels)
print(OUT)
