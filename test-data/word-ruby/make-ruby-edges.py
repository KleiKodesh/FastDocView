# Rewrites the rubies of ruby-edges-skeleton.docx (Word's own output, see make-ruby-edges-fixtures.ps1) into forms Word's UI never
# writes but a .docx may hold: a tab, a footnote reference and a bookmark after a ")" in the base; a footnote reference and a
# bookmark in a base that a comma guide replaces; a ruby with a ")" ending its paragraph; ";" in guide and base.
# Writes ruby-edges-v1.docx. Run from this folder: python -I make-ruby-edges.py
import os
import re
import zipfile

os.chdir(os.path.dirname(os.path.abspath(__file__)))
NOTE_RUN = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="%d"/></w:r>'
BASES = {
    0: ('g0', '<w:r><w:t>ab)c</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>d</w:t></w:r>'),
    1: ('g1', '<w:r><w:t>ab)c</w:t></w:r>' + NOTE_RUN % 1),
    2: ('g2', '<w:r><w:t>ab)c</w:t></w:r><w:bookmarkStart w:id="0" w:name="EdgeTwo"/><w:r><w:t>y</w:t></w:r><w:bookmarkEnd w:id="0"/>'),
    3: ('g,3', '<w:r><w:t>base</w:t></w:r>' + NOTE_RUN % 2),
    4: ('g,4', '<w:r><w:t>base</w:t></w:r><w:bookmarkStart w:id="1" w:name="EdgeFour"/><w:r><w:t>w</w:t></w:r><w:bookmarkEnd w:id="1"/>'),
    5: ('g5', '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">end of line)x  </w:t></w:r>'),
    6: ('g;6', '<w:r><w:t>b;6</w:t></w:r>'),
}
source = zipfile.ZipFile('ruby-edges-skeleton.docx')
parts = {name: source.read(name) for name in source.namelist()}
document = parts['word/document.xml'].decode('utf8')
# the skeleton's "notes holder" paragraph carries footnotes 1 and 2: its references go, the rubies reference them instead
document, count = re.subn(r'<w:r>(?:(?!</w:r>).)*<w:footnoteReference w:id="[12]"/></w:r>', '', document, flags=re.S)
assert count == 2, count
for case, (guide, base) in BASES.items():
    pattern = re.compile(r'(ruby edge %d: (?:(?!</w:p>).)*?<w:rt>).*?(</w:rt><w:rubyBase>).*?(</w:rubyBase>)' % case, re.S)
    document, count = pattern.subn(lambda m: m.group(1) + '<w:r><w:t>' + guide + '</w:t></w:r>' + m.group(2) + base + m.group(3), document, count=1)
    assert count == 1, case
# case 5: the ruby ends its paragraph (its trailing run goes)
document, count = re.subn(r'(ruby edge 5: (?:(?!</w:p>).)*?</w:ruby></w:r>)(?:(?!</w:p>).)*(</w:p>)', r'\1\2', document, count=1, flags=re.S)
assert count == 1
parts['word/document.xml'] = document.encode('utf8')
with zipfile.ZipFile('ruby-edges-v1.docx', 'w', zipfile.ZIP_DEFLATED) as target:
    for name, data in parts.items():
        target.writestr(name, data)
