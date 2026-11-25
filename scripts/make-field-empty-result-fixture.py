# Writes test-data/word-fields/field-empty-result-v1.rtf and field-empty-result-v1.docx (hand-written, same cases):
# REF / PAGEREF / NOTEREF / STYLEREF / SEQ fields whose saved result is empty or missing, locked and unlocked,
# with and without \* MERGEFORMAT, bookmark present / absent. Word's PDFs, .doc and Field results:
# test-data/word-fields/save-field-empty-result.ps1
import os, zipfile
here = os.path.join(os.path.dirname(__file__), '..', 'test-data', 'word-fields')
# (letter, code, lock, result): result None = no result part, '' = empty result
cases = [
    ('A', 'REF bm', False, ''), ('B', 'REF bm', False, None), ('C', r'REF bm \* MERGEFORMAT', False, ''),
    ('D', r'REF bm \* MERGEFORMAT', False, None), ('E', 'REF bm', True, ''), ('F', 'REF bm', True, None),
    ('G', 'REF nobm', False, ''), ('H', 'REF nobm', False, None), ('I', r'REF bm \h', False, ''),
    ('J', 'bm', False, ''), ('K', 'PAGEREF bm', False, ''), ('L', 'NOTEREF nbm', False, ''),
    ('M', 'STYLEREF Banner', False, ''), ('N', 'SEQ fig', False, ''), ('O', 'SEQ fig', False, None),
    ('P', r'REF bm \* Upper', False, ''), ('Q', 'REF bm', False, 'OLD'), ('R', 'REF nobm', True, ''),
]
def rtf_esc(s): return s.replace(chr(92), chr(92) * 2)
r = [r'{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}{\stylesheet{\s0 Normal;}{\s1 heading 1;}{\s2 Banner;}}']
r.append(r'\pard\s1 Head One\par')
r.append(r'\pard\s2 Banner text\par')
r.append(r'\pard\s0 Mark: {\*\bkmkstart bm}BOOKTEXT{\*\bkmkend bm} end. Note{\*\bkmkstart nbm}{\chftn{\footnote\pard\plain {\chftn} the note}}{\*\bkmkend nbm} here.\par')
for l, code, lock, res in cases:
    rs = '' if res is None else (r'{\fldrslt }' if res == '' else r'{\fldrslt %s}' % res)
    r.append(r'\pard case %s: {\field%s{\*\fldinst %s}%s} tail%s\par' % (l, r'\fldlock' if lock else '', rtf_esc(code), rs, l))
r.append('}')
open(os.path.join(here, 'field-empty-result-v1.rtf'), 'w', newline='\n').write('\n'.join(r))

def x(s): return s.replace('&', '&amp;').replace('<', '&lt;').replace('"', '&quot;')
def run(t): return '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % x(t)
body = ['<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>%s</w:p>' % run('Head One'),
        '<w:p><w:pPr><w:pStyle w:val="Banner"/></w:pPr>%s</w:p>' % run('Banner text')]
body.append('<w:p>%s<w:bookmarkStart w:id="1" w:name="bm"/>%s<w:bookmarkEnd w:id="1"/>%s<w:bookmarkStart w:id="2" w:name="nbm"/>'
            '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:bookmarkEnd w:id="2"/>%s</w:p>'
            % (run('Mark: '), run('BOOKTEXT'), run(' end. Note'), run(' here.')))
for l, code, lock, res in cases:
    lk = ' w:fldLock="1"' if lock else ''
    f = '<w:r><w:fldChar w:fldCharType="begin"%s/></w:r><w:r><w:instrText xml:space="preserve"> %s </w:instrText></w:r>' % (lk, x(code))
    if res is not None: f += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + (run(res) if res else '')
    f += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    body.append('<w:p>%s%s%s</w:p>' % (run('case %s: ' % l), f, run(' tail' + l)))
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s<w:sectPr/></w:body></w:document>' % (W, ''.join(body))
styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
          '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Banner"><w:name w:val="Banner"/><w:basedOn w:val="Normal"/></w:style></w:styles>') % W
foot = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes %s><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
        '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>%s</w:p></w:footnote></w:footnotes>') % (W, run(' the note'))
ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>')
with zipfile.ZipFile(os.path.join(here, 'field-empty-result-v1.docx'), 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels); z.writestr('word/document.xml', doc)
    z.writestr('word/_rels/document.xml.rels', drels); z.writestr('word/styles.xml', styles); z.writestr('word/footnotes.xml', foot)
