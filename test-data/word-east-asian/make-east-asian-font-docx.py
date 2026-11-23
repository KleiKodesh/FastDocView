# Builds east-asian-font-v1.docx: body runs with an East Asian font slot (rFonts eastAsia, eastAsiaTheme, w:hint="eastAsia",
# right-to-left and w:cs runs holding CJK). Text written with \u escapes. make-east-asian-font-fixture.ps1 then has Word 16
# open it and save .doc / .rtf / .odt and the reference PDF.
import os, zipfile
here = os.path.dirname(os.path.abspath(__file__))
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
mixed = 'Ab1 é«“—…· 日本あ、가Ａ Z'
def run(props, text): return f'<w:r><w:rPr>{props}<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
def para(tag, props, text, rtl=False):
    ppr = '<w:pPr><w:bidi/></w:pPr>' if rtl else ''
    return f'<w:p>{ppr}{run("<w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/>", tag + " ")}{run(props, text)}</w:p>'
fonts = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="MS Mincho" w:cs="Courier New"/>'
hinted = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="MS Mincho" w:cs="Courier New" w:hint="eastAsia"/>'
themed = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsiaTheme="minorEastAsia" w:cs="Courier New"/>'
sweep = ''.join(chr(c) for c in [0xA1,0xA4,0xA7,0xA8,0xAA,0xAB,0xAD,0xAF,0xB0,0xB1,0xB4,0xB6,0xB7,0xBA,0xBC,0xBF,0xC0,0xD7,0xE0,0xE9,0xF7,0x100,0x2C7,0x391,0x3B1,0x410,0x451,0x2010,0x2018,0x2020,0x2030,0x203B,0x20AC,0x2103,0x2116,0x2160,0x2190,0x2200,0x2312,0x2460,0x2500,0x25A0,0x2605,0x2640,0x3000,0x3008,0xFF01,0xFFE5])
heb = 'אב 日本あ ג'
body = ''.join([
    para('P1', fonts, mixed),
    para('P2', hinted, mixed),
    para('P3', themed, 'Ab 日本あ “…'),
    para('P4', fonts + '<w:rtl/>', heb, rtl=True),
    para('P5', fonts + '<w:cs/>', 'Ab 日本あ Z'),
    para('P6', hinted + '<w:rtl/>', heb + ' “12”', rtl=True),
    para('P7', hinted, sweep),
    para('P8', fonts, sweep),
])
document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
theme = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>'
  '<a:clrScheme name="C"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
  + ''.join(f'<a:accent{i}><a:srgbClr val="4472C4"/></a:accent{i}>' for i in range(1, 7)) +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
  '<a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Mincho"/><a:font script="Hans" typeface="SimSun"/></a:majorFont>'
  '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Mincho"/><a:font script="Hans" typeface="SimSun"/></a:minorFont></a:fontScheme>'
  '<a:fmtScheme name="S"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
  '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
  '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
  '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
  '</a:themeElements></a:theme>')
files = {
 '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
 '_rels/.rels': '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
 'word/_rels/document.xml.rels': '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
 'word/settings.xml': f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:themeFontLang w:val="en-US" w:eastAsia="zh-CN" w:bidi="he-IL"/></w:settings>',
 'word/document.xml': document,
 'word/theme/theme1.xml': theme,
}
with zipfile.ZipFile(os.path.join(here, 'east-asian-font-source-v1.docx'), 'w', zipfile.ZIP_DEFLATED) as z:
    for name, text in files.items(): z.writestr(name, text.encode('utf-8'))
