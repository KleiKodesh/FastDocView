# Writes test-data/word-ruby/note-in-field-keyword-v1.rtf: a note reference before, inside, touching or after a field's
# keyword, plus empty codes, an unknown keyword, EQ and "=" - for Word's rule that a broken keyword leaves the field
# without a type (nothing shown). Word's .docx / .doc / PDFs: test-data/word-ruby/save-note-in-field.ps1 -Name note-in-field-keyword-v1 -Doc
import os
def note(t): return r'{\chftn{\footnote\pard\plain {\chftn} note %s}}' % t
LOCK = r'\fldlock'
cases = [
    ('A', LOCK, r'{MERGEFIELD}' + note('A') + r'{ name}'),       # touching the keyword's end
    ('B', LOCK, r'{MERGEFIELD }' + note('B') + r'{name}'),       # after the keyword and a space
    ('C', LOCK, r'{QUOTE}' + note('C') + r'{ "q"}'),
    ('D', '', r'{MERGEFIELD}' + note('D') + r'{ name}'),         # unlocked
    ('E', LOCK, r'{XYZFIELD name}'),                             # unknown keyword (a bookmark REF)
    ('F', LOCK, note('F') + r'{MERGEFIELD name}'),               # before the keyword
    ('G', LOCK, r'{ MERGEFIELD}' + note('G') + r'{ name}'),
    ('H', LOCK, r'{MERGE}' + note('H') + r'{FIELD name}'),       # inside the keyword
    ('I', LOCK, r'{ }' + note('I') + r'{ }'),
    ('J', LOCK, r'{MERGEFIELDX name}'),
    ('K', LOCK, r'{ }'),                                         # empty code
    ('L', LOCK, r'{PAGE}' + note('L')),
    ('M', LOCK, r'{MERGEFIELD name \\* MERGEFORMAT}'),
    ('N', LOCK, note('N') + r'{ MERGEFIELD name}'),              # before, then a space
    ('O', LOCK, r'{EQ}' + note('O') + r'{ \\o(a,b)}'),
    ('P', '', r'{EQ \\o(a}' + note('P') + r'{,b)}'),           # inside an EQ's arguments (unlocked): drawn in place
    ('Q', '', r'{ }'),                                           # empty code, unlocked
    ('R', LOCK, r'{=}' + note('R') + r'{ 1+2}'),                 # "=" is a keyword by itself
    ('S', LOCK, r'{MERGEFIELD\tab}' + note('S') + r'{name}'),    # a tab ends the keyword
]
text = r'{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}' + '\n'
for letter, lock, code in cases:
    text += r'\pard case %s: {\field%s{\*\fldinst %s}{\fldrslt {res%s}}} tail%s\par' % (letter, lock, code, letter, letter) + '\n'
text += '}'
path = os.path.join(os.path.dirname(__file__), '..', 'test-data', 'word-ruby', 'note-in-field-keyword-v1.rtf')
open(path, 'w', newline='\n').write(text)
