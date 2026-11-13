# Builds list-formats-v1.fodt: one single-level list per LibreOffice numbering type (style:num-format = the identifier
# LibreOffice itself writes, i18npool defaultnumberingprovider.cxx:1078 aSupportedTypes), every item restarted with
# text:start-value at a test value; then LibreOffice prints it to PDF (list-formats-v1.pdf) and to text, and the labels it
# drew are saved as list-formats-v1.expected.txt (type, value, label with non-ASCII as \uXXXX / \UXXXXXXXX).
# Also round-trips the .fodt through LibreOffice and checks every style:num-format comes back unchanged (= LibreOffice
# knows the identifier). Paragraph language en-US (the spelled-out types follow it).
#   python make-list-formats.py "C:\Program Files\LibreOffice\program\soffice.com"
import os, re, subprocess, sys, tempfile, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = 'list-formats-v1'
SOFFICE = sys.argv[1] if len(sys.argv) > 1 else r'C:\Program Files\LibreOffice\program\soffice.com'

def u(*codes):
    return ''.join(chr(c) for c in codes)

def three(a, b, c, tail=''):
    return a + ', ' + b + ', ' + c + ', ...' + tail

CYR = (u(0x410), u(0x411), u(0x430), u(0x431))
def cyr(capital, sync, lang):
    A, B, a, b = CYR
    first, second = (A, B) if capital else (a, b)
    lead = A if capital else a
    lead2 = B if capital else b
    return first + ', ' + second + ', .., ' + lead + a + ', ' + (lead2 + b if sync else lead + b) + ', ... (' + lang + ')'

# (name, identifier) in aSupportedTypes order; generated identifiers = the labels for 1, 2, 3 (makeNumberingIdentifier :1160)
TYPES = [
    ('CHARS_UPPER_LETTER', 'A'), ('CHARS_LOWER_LETTER', 'a'), ('ROMAN_UPPER', 'I'), ('ROMAN_LOWER', 'i'), ('ARABIC', '1'),
    ('SYMBOL_CHICAGO', '*, ' + u(0x2020) + ', ' + u(0x2021) + ', ' + u(0xA7) + ', **, ' + u(0x2020, 0x2020) + ', ...'),
    ('TEXT_NUMBER', '1st'), ('TEXT_CARDINAL', 'One'), ('TEXT_ORDINAL', 'First'),
    ('CHARS_UPPER_LETTER_N', 'A+sync'), ('CHARS_LOWER_LETTER_N', 'a+sync'), ('CHARS_UPPER_LETTER_N_ID', 'AAA'), ('CHARS_LOWER_LETTER_N_ID', 'aaa'),
    ('NATIVE_NUMBERING', 'Native Numbering'),
    ('FULLWIDTH_ARABIC', '1, 2, 3, ...'),   # generated with the en locale: NATNUM3 has no en digits (nativenumbersupplier.cxx getLanguageNumber -1)
    ('CIRCLE_NUMBER', three(u(0x2460), u(0x2461), u(0x2462))),
    ('NUMBER_LOWER_ZH', three(u(0x4E00), u(0x4E8C), u(0x4E09))),
    ('NUMBER_UPPER_ZH', three(u(0x58F9), u(0x8D30), u(0x53C1))),
    ('NUMBER_UPPER_ZH_TW', three(u(0x58F9), u(0x8CB3), u(0x53C3))),
    ('TIAN_GAN_ZH', three(u(0x7532), u(0x4E59), u(0x4E19))),
    ('DI_ZI_ZH', three(u(0x5B50), u(0x4E11), u(0x5BC5))),
    ('NUMBER_TRADITIONAL_JA', three(u(0x58F1), u(0x5F10), u(0x53C2))),
    ('AIU_FULLWIDTH_JA', three(u(0x30A2), u(0x30A4), u(0x30A6))),
    ('AIU_HALFWIDTH_JA', three(u(0xFF71), u(0xFF72), u(0xFF73))),
    ('IROHA_FULLWIDTH_JA', three(u(0x30A4), u(0x30ED), u(0x30CF))),
    ('IROHA_HALFWIDTH_JA', three(u(0xFF72), u(0xFF9B), u(0xFF8A))),
    ('NUMBER_HANGUL_KO', three(u(0xC77C), u(0xC774), u(0xC0BC))),
    ('HANGUL_JAMO_KO', three(u(0x3131), u(0x3134), u(0x3137))),
    ('HANGUL_SYLLABLE_KO', three(u(0xAC00), u(0xB098), u(0xB2E4))),
    ('HANGUL_CIRCLED_JAMO_KO', three(u(0x3260), u(0x3261), u(0x3262))),
    ('HANGUL_CIRCLED_SYLLABLE_KO', three(u(0x326E), u(0x326F), u(0x3270))),
    ('NUMBER_LEGAL_KO', three(u(0xD558, 0xB098), u(0xB458), u(0xC14B))),
    ('NUMBER_DIGITAL_KO', three(u(0xC77C), u(0xC774), u(0xC0BC), ' (ko-x-digital)')),
    ('NUMBER_DIGITAL2_KO', three(u(0x4E00), u(0x4E8C), u(0x4E09), ' (ko)')),
    ('CHARS_ARABIC', three(u(0x623), u(0x628), u(0x62A))),
    ('CHARS_ARABIC_ABJAD', three(u(0x627), u(0x628), u(0x62C))),
    ('NUMBER_ARABIC_INDIC', three(u(0x661), u(0x662), u(0x663))),
    ('NUMBER_EAST_ARABIC_INDIC', three(u(0x6F1), u(0x6F2), u(0x6F3))),
    ('NUMBER_INDIC_DEVANAGARI', three(u(0x967), u(0x968), u(0x969))),
    ('CHARS_THAI', three(u(0xE01), u(0xE02), u(0xE04))),
    ('CHARS_HEBREW', three(u(0x5D0), u(0x5D1), u(0x5D2))),
    ('NUMBER_HEBREW', three(u(0x5D0), u(0x5D9), u(0x5E7))),
    ('CHARS_NEPALI', three(u(0x915), u(0x916), u(0x917))),
    ('CHARS_KHMER', three(u(0x1780), u(0x1781), u(0x1782))),
    ('CHARS_LAO', three(u(0xE81), u(0xE82), u(0xE84))),
    ('CHARS_MYANMAR', three(u(0x1000), u(0x1001), u(0x1002))),
    ('CHARS_TIBETAN', three(u(0xF40), u(0xF41), u(0xF42))),
] + [
    ('CHARS_CYRILLIC_' + ('UPPER' if cap else 'LOWER') + '_LETTER_' + ('N_' if sync else '') + lang.upper(), cyr(cap, sync, lang))
    for lang in ('bg', 'ru', 'sr', 'uk') for sync in (False, True) for cap in (True, False)
] + [
    ('CHARS_PERSIAN', three(u(0x622), u(0x628), u(0x67E))),
    ('CHARS_PERSIAN_WORD', three(u(0x6CC, 0x6A9), u(0x62F, 0x648), u(0x633, 0x647))),
    ('SZEKELY_ROVAS', three(u(0x10CFA), u(0x10CFA, 0x10CFA), u(0x10CFA, 0x10CFA, 0x10CFA))),
    ('CHARS_GREEK_UPPER_LETTER', u(0x391) + ', ' + u(0x392) + ', ... (gr)'),
    ('CHARS_GREEK_LOWER_LETTER', u(0x3B1) + ', ' + u(0x3B2) + ', ... (gr)'),
    ('ARABIC_ZERO', '01, 02, 03, ...'), ('ARABIC_ZERO3', '001, 002, 003, ...'),
    ('ARABIC_ZERO4', '0001, 0002, 0003, ...'), ('ARABIC_ZERO5', '00001, 00002, 00003, ...'),
]

VALUES = list(range(1, 131)) + [150, 199, 200, 201, 210, 211, 300, 999, 1000, 1001, 1010, 1011, 1100, 1110, 1234, 2000, 2001,
          2024, 3000, 5000, 9999, 10000, 10001, 10010, 10100, 11000, 11111, 12345, 20000, 20001, 21000, 30303, 32767]

def esc(text):
    return ''.join(c if 32 <= ord(c) < 127 and c != '\\' else ('\\u%04X' % ord(c) if ord(c) < 0x10000 else '\\U%08X' % ord(c)) for c in text)

def xml(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

def build(path):
    styles, body = [], []
    for index, (name, identifier) in enumerate(TYPES):
        sync = identifier.endswith('+sync')
        fmt = identifier[:-5] if sync else identifier
        styles.append(f'<text:list-style style:name="L{index}"><text:list-level-style-number text:level="1" style:num-suffix="." '
                      f'style:num-format="{xml(fmt)}"' + (' style:num-letter-sync="true"' if sync else '') +
                      '><style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
                      '<style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="2in" '
                      'fo:text-indent="-2in" fo:margin-left="2in"/></style:list-level-properties></text:list-level-style-number></text:list-style>')
        body.append(f'<text:list text:style-name="L{index}">')
        for value in VALUES:
            body.append(f'<text:list-item text:start-value="{value}"><text:p text:style-name="P">K{index}V{value}</text:p></text:list-item>')
        body.append('</text:list>')
    document = ('<?xml version="1.0" encoding="UTF-8"?>\n'
        '<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
        'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">'
        '<office:styles><style:default-style style:family="paragraph"><style:text-properties fo:language="en" fo:country="US" '
        'style:language-asian="en" style:country-asian="US" style:language-complex="en" style:country-complex="US"/></style:default-style>'
        '<style:style style:name="P" style:family="paragraph"><style:paragraph-properties fo:margin-top="0in" fo:margin-bottom="0in"/></style:style>'
        + ''.join(styles) + '</office:styles><office:body><office:text>' + ''.join(body) + '</office:text></office:body></office:document>\n')
    with open(path, 'w', encoding='utf-8', newline='\n') as f:
        f.write(document)

def soffice(profile, convert_to, source, outdir):
    env = dict(os.environ, SAL_DISABLE_PRINTERLIST='1', SAL_DISABLE_DEFAULTPRINTER='1')
    subprocess.run([SOFFICE, '--headless', '--invisible', '--norestore', '--nologo', '-env:UserInstallation=' + profile,
                    '--convert-to', convert_to, '--outdir', outdir, source], check=True, env=env, stdout=subprocess.DEVNULL)

def main():
    fodt = os.path.join(HERE, NAME + '.fodt')
    build(fodt)
    work = tempfile.mkdtemp(prefix='odt-lists-', dir=HERE)
    try:
        profile = 'file:///' + os.path.join(work, 'profile').replace('\\', '/')
        soffice(profile, 'pdf', fodt, HERE)
        soffice(profile, 'txt:Text (encoded):UTF8', fodt, work)
        roundtrip = os.path.join(work, 'roundtrip')
        os.makedirs(roundtrip)
        soffice(profile, 'fodt', fodt, roundtrip)
        written = open(os.path.join(roundtrip, NAME + '.fodt'), encoding='utf-8').read()
        formats = dict(re.findall(r'style:name="(L\d+)">\s*<text:list-level-style-number text:level="1"[^>]*?style:num-format="([^"]*)"', written))
        lost = []
        for index, (name, identifier) in enumerate(TYPES):
            want = identifier[:-5] if identifier.endswith('+sync') else identifier
            got = formats.get('L%d' % index)
            got = got.replace('&quot;', '"').replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>') if got is not None else None
            if got != want and not (name.endswith('_N_ID') and got == want[0]):
                lost.append('%s: wrote %s, LibreOffice saved %s' % (name, esc(want), esc(got or '(none)')))
        text = open(os.path.join(work, NAME + '.txt'), encoding='utf-8-sig').read()
    finally:
        shutil.rmtree(work, ignore_errors=True)
    rows, labels = [], {}
    for line in text.splitlines():
        m = re.match(r'^ *(.*\.) K(\d+)V(\d+)$', line)
        if m:
            rows.append('%s\t%s\t%s' % (TYPES[int(m.group(2))][0], m.group(3), esc(m.group(1))))
            labels[(m.group(2), m.group(3))] = m.group(1)
    # The PDF must show the same labels: pdftotext -layout puts each label on its item's line, in visual order, so they are
    # compared as multisets of characters (spaces and bidi controls dropped). pdftotext writes astral characters as two
    # 3-byte surrogates (CESU-8): decoded with surrogatepass, then paired through UTF-16.
    layout = subprocess.run(['pdftotext', '-layout', '-enc', 'UTF-8', os.path.join(HERE, NAME + '.pdf'), '-'],
                            capture_output=True, check=True).stdout.decode('utf-8', 'surrogatepass')
    layout = layout.encode('utf-16-le', 'surrogatepass').decode('utf-16-le', 'replace')
    controls = set(chr(c) for c in (0x200E, 0x200F, 0x202A, 0x202B, 0x202C, 0x202D, 0x202E))
    def chars(t):
        return sorted(c for c in t if not c.isspace() and c not in controls)
    seen, differ = set(), []
    lines = layout.splitlines()
    for number, line in enumerate(lines):
        for m in re.finditer(r'K(\d+)V(\d+)', line):
            key = (m.group(1), m.group(2))
            if key not in labels or key in seen:
                continue
            seen.add(key)
            # after a page break the label can land on the line below its text
            below = re.sub(r'K\d+V\d+', '', lines[number + 1]) if number + 1 < len(lines) else ''
            if chars(line.replace(m.group(0), '')) != chars(labels[key]) and chars(below) != chars(labels[key]):
                differ.append(key)
    # labels longer than a line (Chicago's repeated symbols, long words) wrap in the PDF and are not comparable this way
    differ = [key for key in differ if len(labels[key]) <= 40]
    print('PDF: %d of %d labels found, %d of the one-line ones differ from the text export' % (len(seen), len(labels), len(differ)))
    by_type = {}
    for key in differ:
        by_type.setdefault(TYPES[int(key[0])][0], []).append(key[1])
    for name, values in by_type.items():
        print('  %s: %s' % (name, ' '.join(values)))
    with open(os.path.join(HERE, NAME + '.expected.txt'), 'w', encoding='ascii', newline='\n') as f:
        f.write('\n'.join(rows) + '\n')
    print('%d labels, %d types, identifiers LibreOffice did not keep: %d' % (len(rows), len(TYPES), len(lost)))
    for item in lost:
        print('  ' + item)

main()
