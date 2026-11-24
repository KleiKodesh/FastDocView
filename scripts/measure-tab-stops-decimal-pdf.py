# Reads Word's PDF of test-data/tab-stops/tab-stops-decimal-v1 and prints JSON for vue-demo/playwright/tabDecimalUnderlineCheck.mjs:
# per decimal paragraph the x (pt from the left margin) of each '.' or ',' and of its text's end; per underlined paragraph
# the x ranges of the underline Word draws (filled rectangles under the line). Usage: python measure-tab-stops-decimal-pdf.py <pdf>
import json, sys
import fitz

page = fitz.open(sys.argv[1])[0]
MARGIN = 72.0
lines = {}
for block in page.get_text('rawdict')['blocks']:
    for line in block.get('lines', []):
        for span in line['spans']:
            for char in span['chars']:
                lines.setdefault(round(span['origin'][1]), []).append((char['origin'][0], char['bbox'][2], char['c']))
ordered = [(y, sorted(chars)) for y, chars in sorted(lines.items())]
rules = [d['rect'] for d in page.get_drawings() if d.get('fill') is not None]
paragraphs = []
for k, (y, chars) in enumerate(ordered):
    text = ''.join(c for _, _, c in chars)
    entry = {'first': text.split(' ')[0], 'text': text.strip()}
    if k < 7:
        separators = [round(x - MARGIN, 2) for x, _, c in chars if c in '.,']
        entry['separators'] = separators
        entry['end'] = round([x1 for _, x1, c in chars if c != ' '][-1] - MARGIN, 2)   # Word's PDF adds a space after the text
    else:
        under = sorted((round(r.x0 - MARGIN, 1), round(r.x1 - MARGIN, 1)) for r in rules if y - 2 < r.y0 < y + 6)
        # a double underline is two rectangles over the same range: one range each
        entry['underline'] = [list(r) for r in dict.fromkeys(under)]
    paragraphs.append(entry)
print(json.dumps({'paragraphs': paragraphs}, indent=1))
