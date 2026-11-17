# Builds ruby-escapes-probe-v1.rtf from Word's own ruby-escapes-v1.rtf: the guides and bases of cases 2-4 get backslash
# "escapes" (\, \) \( \) and cases 0-1 an unquoted multi-word Font: option - forms Word never writes itself, to see how its
# EQ parser reads them. make-ruby-escapes-fixtures.ps1 then has Word print it to ruby-escapes-probe-v1.pdf.
# Run from this folder: python -I make-ruby-escapes-probe.py
BACKSLASH = chr(92)
RTF_BACKSLASH = BACKSLASH * 2      # a literal backslash inside RTF text
text = open('ruby-escapes-v1.rtf', 'rb').read().decode('latin1')


def replace_once(old, new):
    global text
    assert text.count(old) == 1, old
    text = text.replace(old, new)


replace_once(' g;2}', ' a' + RTF_BACKSLASH + ',b}')
replace_once('),base;2)}', '),c' + RTF_BACKSLASH + ',d)}')
replace_once(' g' + RTF_BACKSLASH + '3}', ' e' + RTF_BACKSLASH + ')f}')
replace_once('),base' + RTF_BACKSLASH + '3)}', '),g' + RTF_BACKSLASH + ')h)}')
replace_once(' g)4(}', ' i' + RTF_BACKSLASH + '(j}')
replace_once('),base)4()}', '),k' + RTF_BACKSLASH * 2 + 'l)}')
replace_once('"Font:Times New Roman"', 'Font:Times New Roman')
replace_once('"Font:Courier New"', 'Font:Courier New')
open('ruby-escapes-probe-v1.rtf', 'wb').write(text.encode('latin1'))
