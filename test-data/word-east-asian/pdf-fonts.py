# Prints, per line of a PDF, each character's code point and font (ASCII only): which face Word drew it in.
import sys, fitz
for page in fitz.open(sys.argv[1]):
    for block in page.get_text('rawdict')['blocks']:
        for line in block.get('lines', []):
            out = []
            for span in line['spans']:
                chars = ''.join(c['c'] for c in span['chars'])
                out.append(span['font'] + ':' + ''.join(c if ord(c) < 128 else 'U+%04X' % ord(c) for c in chars))
            print(' | '.join(out))
