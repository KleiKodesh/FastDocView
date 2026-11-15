# Builds list-languages-v1.fodt: the list number styles that follow the paragraph's language (Native Numbering, full-width
# Arabic, 1st / One / First) in paragraphs of many languages - Western ones by fo:language, complex-script ones (Hebrew,
# Arabic, Persian, Thai, Indic...) by style:language-complex under Hebrew / Arabic / ... text, Asian ones by
# style:language-asian under CJK text - plus the cases that pick the language: a span at the first character, a digit first
# (Western script), an empty paragraph. LibreOffice prints it to PDF (list-languages-v1.pdf, not committed) and to text; the
# labels it drew are saved as list-languages-v1.expected.txt (case, value, label with non-ASCII as \uXXXX / \UXXXXXXXX).
#   python make-list-languages.py "C:\Program Files\LibreOffice\program\soffice.com"
import os, re, subprocess, sys, tempfile, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = 'list-languages-v1'
SOFFICE = sys.argv[1] if len(sys.argv) > 1 else r'C:\Program Files\LibreOffice\program\soffice.com'

FORMATS = [('NATIVE', 'Native Numbering'), ('FULLWIDTH', '1, 2, 3, ...'), ('TEXT_NUMBER', '1st'), ('TEXT_CARDINAL', 'One'), ('TEXT_ORDINAL', 'First')]
VALUES = list(range(1, 32)) + [40, 99, 100, 101, 111, 200, 999, 1000, 1001, 1234, 2000, 2024, 10000, 12345, 32767]

WESTERN = ['af-ZA', 'bg-BG', 'ca-ES', 'cs-CZ', 'cu-RU', 'da-DK', 'de-DE', 'de-CH', 'el-GR', 'en-US', 'en-GB', 'en-IN', 'eo', 'es-ES',
           'et-EE', 'fi-FI', 'fr-FR', 'fr-BE', 'ga-IE', 'gl-ES', 'hr-HR', 'hu-HU', 'id-ID', 'is-IS', 'it-IT', 'lb-LU', 'lg-UG', 'lt-LT',
           'lv-LV', 'mn-MN', 'ms-MY', 'mt-MT', 'nb-NO', 'nn-NO', 'nl-NL', 'pl-PL', 'pt-BR', 'pt-PT', 'ro-RO', 'ru-RU', 'sk-SK', 'sl-SI',
           'sq-AL', 'sr-RS', 'sv-SE', 'sw-KE', 'tr-TR', 'uk-UA', 'vi-VN']
# complex-script languages with a letter of their script to start the paragraph
COMPLEX = [('he-IL', 0x05D0), ('ar-EG', 0x0627), ('ar-SA', 0x0627), ('fa-IR', 0x0627), ('th-TH', 0x0E01), ('hi-IN', 0x0915),
           ('mr-IN', 0x0915), ('bn-IN', 0x0995), ('pa-IN', 0x0A15), ('gu-IN', 0x0A95), ('or-IN', 0x0B15), ('ta-IN', 0x0B95),
           ('te-IN', 0x0C15), ('kn-IN', 0x0C95), ('ml-IN', 0x0D15), ('lo-LA', 0x0E81), ('bo-CN', 0x0F40), ('my-MM', 0x1000),
           ('km-KH', 0x1780), ('ne-NP', 0x0915), ('dz-BT', 0x0F40)]
ASIAN = [('ja-JP', 0x3042), ('ko-KR', 0xAC00), ('zh-CN', 0x4E00), ('zh-TW', 0x4E00), ('zh-HK', 0x4E00)]

def split(tag):
    parts = tag.split('-')
    return parts[0], (parts[1] if len(parts) > 1 else '')

def esc(text):
    return ''.join(c if 32 <= ord(c) < 127 and c != '\\' else ('\\u%04X' % ord(c) if ord(c) < 0x10000 else '\\U%08X' % ord(c)) for c in text)

def latin_props(tag):
    language, country = split(tag)
    return f'fo:language="{language}"' + (f' fo:country="{country}"' if country else '')

def build(path):
    # cases: (name, paragraph style attributes, span language or None, first text)
    cases = []
    for tag in WESTERN:
        cases.append((tag, latin_props(tag), None, 'x'))
    for tag, letter in COMPLEX:
        language, country = split(tag)
        cases.append((tag, f'fo:language="en" fo:country="US" style:language-complex="{language}" style:country-complex="{country}"', None, chr(letter)))
    for tag, letter in ASIAN:
        language, country = split(tag)
        cases.append((tag, f'fo:language="en" fo:country="US" style:language-asian="{language}" style:country-asian="{country}"', None, chr(letter)))
    # the language is the one at the first character: a digit is Western script, a span there wins, an empty paragraph is Western
    cases.append(('he-IL digit first', 'fo:language="de" fo:country="DE" style:language-complex="he" style:country-complex="IL"', None, '5' + chr(0x05D0)))
    cases.append(('span de over en', 'fo:language="en" fo:country="US"', 'fo:language="de" fo:country="DE"', 'x'))
    cases.append(('span he over en', 'fo:language="en" fo:country="US" style:language-complex="ar" style:country-complex="EG"', 'style:language-complex="he" style:country-complex="IL"', chr(0x05D0)))
    cases.append(('empty fr', 'fo:language="fr" fo:country="FR" style:language-complex="he" style:country-complex="IL"', None, ''))
    styles, automatic, body, names = [], [], [], []
    for f, (fname, identifier) in enumerate(FORMATS):
        styles.append(f'<text:list-style style:name="L{f}"><text:list-level-style-number text:level="1" style:num-suffix="." '
                      f'style:num-format="{identifier}"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
                      '<style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="3in" '
                      'fo:text-indent="-3in" fo:margin-left="3in"/></style:list-level-properties></text:list-level-style-number></text:list-style>')
    for c, (name, attributes, span, first) in enumerate(cases):
        styles.append(f'<style:style style:name="P{c}" style:family="paragraph"><style:paragraph-properties fo:margin-top="0in" '
                      f'fo:margin-bottom="0in"/><style:text-properties {attributes}/></style:style>')
        if span:
            automatic.append(f'<style:style style:name="T{c}" style:family="text"><style:text-properties {span}/></style:style>')
        for f, (fname, identifier) in enumerate(FORMATS):
            index = len(names)
            names.append(name + ' ' + fname)
            body.append(f'<text:list text:style-name="L{f}">')
            for value in VALUES:
                marker = f'K{index}V{value}'
                text = (f'<text:span text:style-name="T{c}">{first}</text:span> {marker}' if span else
                        f'{first} {marker}' if first else f'<text:bookmark text:name="{marker}"/>')
                body.append(f'<text:list-item text:start-value="{value}"><text:p text:style-name="P{c}">{text}</text:p></text:list-item>')
            body.append('</text:list>')
    document = ('<?xml version="1.0" encoding="UTF-8"?>\n'
        '<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
        'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">'
        '<office:styles><style:default-style style:family="paragraph"><style:text-properties fo:language="en" fo:country="US" '
        'style:language-asian="en" style:country-asian="US" style:language-complex="en" style:country-complex="US"/></style:default-style>'
        + ''.join(styles) + '</office:styles><office:automatic-styles>' + ''.join(automatic) + '</office:automatic-styles>'
        '<office:body><office:text>' + ''.join(body) + '</office:text></office:body></office:document>\n')
    with open(path, 'w', encoding='utf-8', newline='\n') as f:
        f.write(document)
    return names

def soffice(profile, convert_to, source, outdir):
    env = dict(os.environ, SAL_DISABLE_PRINTERLIST='1', SAL_DISABLE_DEFAULTPRINTER='1')
    subprocess.run([SOFFICE, '--headless', '--invisible', '--norestore', '--nologo', '-env:UserInstallation=' + profile,
                    '--convert-to', convert_to, '--outdir', outdir, source], check=True, env=env, stdout=subprocess.DEVNULL)

def main():
    fodt = os.path.join(HERE, NAME + '.fodt')
    names = build(fodt)
    work = tempfile.mkdtemp(prefix='odt-lists-', dir=HERE)
    try:
        profile = 'file:///' + os.path.join(work, 'profile').replace('\\', '/')
        soffice(profile, 'pdf', fodt, HERE)
        soffice(profile, 'txt:Text (encoded):UTF8', fodt, work)
        text = open(os.path.join(work, NAME + '.txt'), encoding='utf-8-sig').read()
    finally:
        shutil.rmtree(work, ignore_errors=True)
    # each item is one line: indent, label, a space, the paragraph text (its first character and the marker); the empty
    # paragraphs have no marker and are matched by order
    rows, labels = [], {}
    order = [(i, v) for i in range(len(names)) for v in VALUES]
    lines = [line for line in text.splitlines() if line.strip()]
    if len(lines) != len(order):
        raise SystemExit('expected %d lines, got %d' % (len(order), len(lines)))
    for (index, value), line in zip(order, lines):
        line = line.strip(' ')
        m = re.match(r'^(.*\.)(?: \S{0,2} ?K(\d+)V(\d+))?$', line)
        if not m or (m.group(2) and (int(m.group(2)), int(m.group(3))) != (index, value)):
            raise SystemExit('cannot read line for %s %d: %s' % (names[index], value, esc(line)))
        rows.append('%s\t%d\t%s' % (names[index], value, esc(m.group(1))))
        labels[(index, value)] = m.group(1)
    with open(os.path.join(HERE, NAME + '.expected.txt'), 'w', encoding='ascii', newline='\n') as f:
        f.write('\n'.join(rows) + '\n')
    # the PDF must show the same labels (pdftotext -layout, compared as multisets of characters; see make-list-formats.py)
    layout = subprocess.run(['pdftotext', '-layout', '-enc', 'UTF-8', os.path.join(HERE, NAME + '.pdf'), '-'],
                            capture_output=True, check=True).stdout.decode('utf-8', 'surrogatepass')
    layout = layout.encode('utf-16-le', 'surrogatepass').decode('utf-16-le', 'replace')
    controls = set(chr(c) for c in (0x200E, 0x200F, 0x202A, 0x202B, 0x202C, 0x202D, 0x202E))
    def chars(t):
        return sorted(c for c in t if not c.isspace() and c not in controls)
    seen, differ = set(), []
    pdf_lines = layout.splitlines()
    for number, line in enumerate(pdf_lines):
        for m in re.finditer(r'K(\d+)V(\d+)', line):
            key = (int(m.group(1)), int(m.group(2)))
            if key not in labels or key in seen:
                continue
            seen.add(key)
            want = chars(labels[key])
            rest = re.sub(r'K\d+V\d+', '', line)
            below = re.sub(r'K\d+V\d+', '', pdf_lines[number + 1]) if number + 1 < len(pdf_lines) else ''
            # the line also holds the paragraph's first letter
            if not any(chars(t)[:0] == [] and sorted(want) == sorted(c for c in chars(t) if c in want or True)[:len(want)] and
                       all(chars(t).count(c) >= want.count(c) for c in set(want)) and len(chars(t)) - len(want) <= 2 for t in (rest, below)):
                if len(labels[key]) <= 40:
                    differ.append(key)
    print('%d labels; PDF: %d with a marker found, %d of the one-line ones differ' % (len(rows), len(seen), len(differ)))
    by_case = {}
    for key in differ:
        by_case.setdefault(names[key[0]], []).append(str(key[1]))
    for name, values in by_case.items():
        print('  %s: %s' % (esc(name), ' '.join(values)))

main()
