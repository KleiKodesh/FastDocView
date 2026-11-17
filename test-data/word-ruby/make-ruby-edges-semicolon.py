# Builds ruby-edges-semicolon-v1.rtf from Word's ruby-edges-v1.rtf: case 6's EQ code gets ";" in place of the comma between guide
# and base (no comma left in the code), the separator LibreOffice falls back to; make-ruby-edges-fixtures.ps1 has Word print it.
# Run from this folder: python -I make-ruby-edges-semicolon.py
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))
text = open('ruby-edges-v1.rtf', 'rb').read().decode('latin1')
old = '(g;6),b;6)'
assert text.count(old) == 1
text = text.replace(old, '(g;6);b;6)')
open('ruby-edges-semicolon-v1.rtf', 'wb').write(text.encode('latin1'))
