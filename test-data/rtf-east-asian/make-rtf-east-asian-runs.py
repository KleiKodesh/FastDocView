# Builds rtf-east-asian-runs-v1.rtf: East Asian characters inside right-to-left (\rtlch) and complex-script (\ltrch\fcs1) runs,
# with and without \dbch\afN, and the size / bold slots Word gives them. Reference: Word 16's PDF of the file itself
# (rtf-east-asian-runs-v1.pdf; make-rtf-east-asian-fixture.ps1). Word 16 (pdf-fonts.py in ../word-east-asian):
#   A1 \rtlch\f22           Hebrew Tahoma, Han SimSun (the document's \stshfdbch), Latin Tahoma
#   A3 \rtlch\f22\afs40\fs20 Hebrew 10pt, Han 12pt (\fs / \afs after \rtlch size only the complex-script text)
#   A5 \rtlch\fcs1\af22\afs40\ltrch\fcs0\f1\fs20\dbch\af20  Hebrew Tahoma 20pt, Latin Arial 10pt, Han MS Mincho 10pt (Word's own form)
#   A6 \rtlch\fcs1\af22\ab\afs40\ltrch\fcs0\b0\fs20\dbch\af23  Hebrew Tahoma bold 20pt, Han Microsoft YaHei 10pt not bold
#   A8 \rtlch\dbch\af20\f23 Han MS Mincho (\f after \dbch\af keeps it), Latin Microsoft YaHei
#   A9 \ltrch\dbch\af20\f23 Han Microsoft YaHei (\f after \dbch in a left-to-right run sets it)
#   A10 \rtlch\dbch\f23\loch\af20 Han Microsoft YaHei (\f after \dbch sets it while none is set), Latin MS Mincho
#       (\loch\af in a right-to-left run still names the complex-script face)
# Word's reader carries some of this state from one run group into the next (A8 drew in Microsoft YaHei after a
# \ltrch\fcs1\af22\ab0\afs40\rtlch\fcs0\b\fs20\dbch\af23 run), so the runs keep this order; Word itself never writes them.
# LibreOffice reads A1-A6 the same way (rtfdispatchvalue.cxx:171-187, 739-756) except that it has no \stshfdbch fallback (A1).
B = chr(92)
def u(text): return ''.join(c if ord(c) < 128 else B + 'u%d?' % ord(c) for c in text)
def cw(*words): return ''.join(B + w for w in words) + ' '
HEB = '\u05e9\u05dc\u05d5\u05dd'
HAN = '\u65e5\u672c'
runs = [
    ('A1', cw('rtlch', 'f22')),
    ('A3', cw('rtlch', 'f22', 'afs40', 'fs20')),
    ('A5', cw('rtlch', 'fcs1', 'af22', 'afs40', 'ltrch', 'fcs0', 'f1', 'fs20', 'dbch', 'af20')),
    ('A6', cw('rtlch', 'fcs1', 'af22', 'ab', 'afs40', 'ltrch', 'fcs0', 'b0', 'fs20', 'dbch', 'af23')),
    ('A8', cw('rtlch', 'dbch', 'af20', 'f23')),
    ('A9', cw('ltrch', 'dbch', 'af20', 'f23')),
    ('A10', cw('rtlch', 'dbch', 'f23', 'loch', 'af20')),
]
fonts = [(0, 'froman', '', 'Times New Roman'), (1, 'fswiss', '', 'Arial'), (2, 'fmodern', '', 'Courier New'),
         (20, 'fnil', 'fcharset128', 'MS Mincho'), (21, 'fnil', 'fcharset134', 'SimSun'), (22, 'fswiss', 'fcharset177', 'Tahoma'),
         (23, 'fnil', 'fcharset134', 'Microsoft YaHei')]
out = '{' + cw('rtf1', 'ansi', 'ansicpg1252', 'deff0', 'stshfdbch21', 'stshfloch1', 'stshfhich1', 'stshfbi22').rstrip() + '{' + B + 'fonttbl'
for number, family, charset, name in fonts:
    out += '{' + B + 'f%d' % number + B + family + (B + charset if charset else '') + ' ' + name + ';}'
out += '}\n{' + B + 'stylesheet{' + cw('rtlch', 'af22', 'afs24', 'ltrch', 'f1', 'fs24', 'loch', 'f1', 'hich', 'af1', 'dbch', 'af21') + 'Normal;}}\n'
for tag, words in runs:
    out += '{' + cw('pard', 'plain', 'f1', 'fs24') + tag + ' }{' + words + u(HEB + ' ' + HAN + ' Ab') + '}' + B + 'par\n'
out += '}'
open('rtf-east-asian-runs-v1.rtf', 'w', encoding='ascii', newline='\r\n').write(out)
