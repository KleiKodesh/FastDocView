# Writes test-data/word-ruby/field-leading-switch-v1.rtf: field codes that start with a switch (no keyword first), a stray
# backslash before a keyword, and nested fields with notes in their codes (each field's code notes are its own). Word's
# .docx / .doc / PDFs and each field's Type: test-data/word-ruby/save-note-in-field.ps1 -Name field-leading-switch-v1 -Doc
import os
def note(t): return r'{\chftn{\footnote\pard\plain {\chftn} note %s}}' % t
LOCK = r'\fldlock'
cases = [
    ('A', LOCK, r'{\\* MERGEFORMAT}'),                         # a format switch alone
    ('B', LOCK, r'{\\* MERGEFORMAT MERGEFIELD name}'),         # a format switch, then a keyword
    ('C', LOCK, r'{\\h REF bm}'),                              # a field switch first
    ('D', LOCK, r'{\\@ "d MMMM" DATE}'),                       # a date switch first
    ('E', LOCK, r'{\\MERGEFIELD name}'),                       # a stray backslash before the keyword
    ('F', '', r'{\\* MERGEFORMAT}'),                           # unlocked
    ('G', LOCK, r'{ \\# "0" = 1+2}'),                          # a numeric switch before "="
    ('H', LOCK, r'{"MERGEFIELD" name}'),                       # a quoted keyword
    ('J', LOCK, r'{\\hREF bm}'),                               # a backslash, then a letter and more letters
    ('K', LOCK, r'{\\*MERGEFORMAT}'),                          # a backslash, then a non-letter and letters
    ('L', LOCK, r'{\\M name}'),                                # a one-letter "keyword" after a backslash
    # nested: the inner field's code note is its own (no mark: its keyword is broken); the EQ's own note is drawn in place
    ('I', '', r'{EQ \\o(}{\field{\*\fldinst {MERGEFIELD}' + note('I1') + r'{ n}}{\fldrslt {x}}}{,b}' + note('I2') + r'{)}'),
]
text = r'{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}' + '\n'
for letter, lock, code in cases:
    text += r'\pard case %s: {\field%s{\*\fldinst %s}{\fldrslt {res%s}}} tail%s\par' % (letter, lock, code, letter, letter) + '\n'
text += '}'
path = os.path.join(os.path.dirname(__file__), '..', 'test-data', 'word-ruby', 'field-leading-switch-v1.rtf')
open(path, 'w', newline='\n').write(text)
