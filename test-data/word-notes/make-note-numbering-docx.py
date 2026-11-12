# Writes note-numbering-v1.docx: footnote / endnote number format, start and restart at the document level
# (settings.xml) and per section (w:sectPr). make-note-numbering-fixtures.ps1 then has Word 16 open it and save the
# .doc, .rtf and Word's own PDF (the reference for the printed numbers). Run: python make-note-numbering-docx.py
import os, zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
body, footnotes, endnotes = [], [], []
note_id = {'footnote': 0, 'endnote': 0}


def note(kind, text):
    note_id[kind] += 1
    target = footnotes if kind == 'footnote' else endnotes
    target.append(f'<w:{kind} w:id="{note_id[kind]}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:{kind}Ref/></w:r>'
                  f'<w:r><w:t xml:space="preserve"> {text}</w:t></w:r></w:p></w:{kind}>')
    return f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:{kind}Reference w:id="{note_id[kind]}"/></w:r>'


def para(text, notes=(), section=None, page_break=False):
    properties = f'<w:pPr>{section}</w:pPr>' if section else ''
    page = '<w:r><w:br w:type="page"/></w:r>' if page_break else ''
    body.append(f'<w:p>{properties}{page}<w:r><w:t>{text}</w:t></w:r>' + ''.join(note(k, t) for k, t in notes) + '</w:p>')


def section(foot, end):
    f = f'<w:footnotePr>{foot}</w:footnotePr>' if foot is not None else ''
    e = f'<w:endnotePr>{end}</w:endnotePr>' if end is not None else ''
    return (f'<w:sectPr>{f}{e}<w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


# 1: no section-level note properties - the document's (settings.xml) apply
para('s1 a', [('footnote', 'fn s1 a')])
para('s1 b', [('footnote', 'fn s1 b'), ('endnote', 'en s1')], section=section(None, None))
# 2: footnotes restart each section, start 2, lower letter; endnotes restart each section
para('s2 a', [('footnote', 'fn s2 a')])
para('s2 b', [('footnote', 'fn s2 b'), ('endnote', 'en s2')],
     section=section('<w:numFmt w:val="lowerLetter"/><w:numStart w:val="2"/><w:numRestart w:val="eachSect"/>', '<w:numRestart w:val="eachSect"/>'))
# 3: continuous, start 10, no format given; endnote properties present but empty
para('s3 a', [('footnote', 'fn s3 a')])
para('s3 b', [('footnote', 'fn s3 b'), ('endnote', 'en s3')], section=section('<w:numStart w:val="10"/>', ''))
# 4: footnotes restart each page, two notes on two pages; endnotes upper letter
para('s4 a', [('footnote', 'fn s4 a')])
para('s4 b', [('footnote', 'fn s4 b'), ('endnote', 'en s4')], page_break=True,
     section=section('<w:numRestart w:val="eachPage"/>', '<w:numFmt w:val="upperLetter"/>'))
# 5 (the body's own w:sectPr): continuous after the restarts, upper roman; no endnote properties
para('s5 a', [('footnote', 'fn s5 a'), ('endnote', 'en s5')])
last = section('<w:numFmt w:val="upperRoman"/>', None)


def separators(kind):
    return (f'<w:{kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:{kind}>'
            f'<w:{kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{kind}>')


HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
files = {
    '[Content_Types].xml': HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/word/document.xml" ContentType="{CT}document.main+xml"/><Override PartName="/word/settings.xml" ContentType="{CT}settings+xml"/>'
        f'<Override PartName="/word/footnotes.xml" ContentType="{CT}footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="{CT}endnotes+xml"/></Types>',
    '_rels/.rels': HEAD + f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}officeDocument" Target="word/document.xml"/></Relationships>',
    'word/_rels/document.xml.rels': HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{REL}settings" Target="settings.xml"/><Relationship Id="rId2" Type="{REL}footnotes" Target="footnotes.xml"/>'
        f'<Relationship Id="rId3" Type="{REL}endnotes" Target="endnotes.xml"/></Relationships>',
    'word/settings.xml': HEAD + f'<w:settings {W}><w:footnotePr><w:numFmt w:val="upperRoman"/><w:numStart w:val="3"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
        '<w:endnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="5"/><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr></w:settings>',
    'word/document.xml': HEAD + f'<w:document {W}><w:body>{"".join(body)}{last}</w:body></w:document>',
    'word/footnotes.xml': HEAD + f'<w:footnotes {W}>{separators("footnote")}{"".join(footnotes)}</w:footnotes>',
    'word/endnotes.xml': HEAD + f'<w:endnotes {W}>{separators("endnote")}{"".join(endnotes)}</w:endnotes>',
}
with zipfile.ZipFile(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'note-numbering-v1.docx'), 'w', zipfile.ZIP_DEFLATED) as package:
    for name, text in files.items():
        package.writestr(name, text)
