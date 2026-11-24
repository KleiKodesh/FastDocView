# Writes test-data/tab-stops/tab-stops-decimal-v1.docx: decimal stops at 200 pt with "12,34" / "12.34" / "1.234,5" in en-US, de-DE
# and default-language runs, in a document whose defaults are de-DE and whose settings say w:decimalSymbol "," - Word 16 still
# aligns every one on '.' (the system's decimal symbol; a run without '.' sits as on a right stop); then underlined tabs: single,
# a trailing tab, a tab underlined after plain text, w:u words, double, a dot leader and a default stop. ASCII text only.
# scripts/save-tab-stops-decimal.ps1 has Word save it as PDF (the reference); measure-tab-stops-decimal-pdf.py reads that.
# (2026-11-25 the same paragraphs were also checked in en-US / de-DE documents with and without w:decimalSymbol: identical PDFs.)
import os, zipfile
OUT=os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'test-data', 'tab-stops')
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def run(t,rpr=''):
    if t=='\t': return f'<w:r><w:rPr>{rpr}</w:rPr><w:tab/></w:r>'
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
def p(parts,ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr>'+''.join(run(t,r) for t,r in parts)+'</w:p>'
dec='<w:tabs><w:tab w:val="decimal" w:pos="4000"/></w:tabs>'
L=lambda l:f'<w:lang w:val="{l}"/>'
def make(name, deflang, decsym):
    body=[
     p([('A',''),('\t',''),('12,34',L('en-US'))],dec),
     p([('B',''),('\t',''),('12,34',L('de-DE'))],dec),
     p([('C',''),('\t',''),('12.34',L('de-DE'))],dec),
     p([('D',''),('\t',''),('12.34',L('en-US'))],dec),
     p([('E',''),('\t',''),('12,34','')],dec),
     p([('F',''),('\t',''),('12.34','')],dec),
     p([('G',''),('\t',''),('1.234,5',L('de-DE'))],dec),
     # underline probes
     p([('Total','<w:u w:val="single"/>'),('\t','<w:u w:val="single"/>'),('12','<w:u w:val="single"/>')],'<w:tabs><w:tab w:val="right" w:pos="5000"/></w:tabs>'),
     p([('Name:','<w:u w:val="single"/>'),('\t','<w:u w:val="single"/>')],'<w:tabs><w:tab w:val="left" w:pos="5000"/></w:tabs>'),
     p([('Name:',''),('\t','<w:u w:val="single"/>')],'<w:tabs><w:tab w:val="left" w:pos="5000"/></w:tabs>'),
     p([('Word','<w:u w:val="words"/>'),('\t','<w:u w:val="words"/>'),('next','<w:u w:val="words"/>')],'<w:tabs><w:tab w:val="left" w:pos="5000"/></w:tabs>'),
     p([('Dbl','<w:u w:val="double"/>'),('\t','<w:u w:val="double"/>'),('x','<w:u w:val="double"/>')],'<w:tabs><w:tab w:val="left" w:pos="5000"/></w:tabs>'),
     p([('Lead','<w:u w:val="single"/>'),('\t','<w:u w:val="single"/>'),('9','<w:u w:val="single"/>')],'<w:tabs><w:tab w:val="right" w:pos="5000" w:leader="dot"/></w:tabs>'),
     p([('Def','<w:u w:val="single"/>'),('\t','<w:u w:val="single"/>'),('y','')],''),
    ]
    doc=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
    styles=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:lang w:val="{deflang}"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
    ds=f'<w:decimalSymbol w:val="{decsym}"/><w:listSeparator w:val=";"/>' if decsym else ''
    settings=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>{ds}</w:settings>'
    ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
    rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    drels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
    with zipfile.ZipFile(os.path.join(OUT,name+'.docx'),'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',ct);z.writestr('_rels/.rels',rels);z.writestr('word/document.xml',doc);z.writestr('word/styles.xml',styles);z.writestr('word/settings.xml',settings);z.writestr('word/_rels/document.xml.rels',drels)
make('tab-stops-decimal-v1', 'de-DE', ',')
