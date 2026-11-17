# Reads a reference PDF of test-data/tab-stops/tab-stops-v2 (Word's, or LibreOffice's of the .odt) and prints, per paragraph, the
# x (pt from the left margin) where each tab's following text starts: the first word of each segment after a tab, found in the
# paragraph's line. JSON on stdout, read by vue-demo/playwright/tabStopsCheck.mjs. Usage: python measure-tab-stops-pdf.py <pdf> [margin pt]
import json, sys
import fitz

SEGMENTS = [
    ['Chapter', '12'], ['A', '345'], ['Hyphen', '7'], ['Underscore', '8'], ['Heavy', '9'], ['Middle', '10'],
    ['Left', 'Centred', 'Right'], ['Price', '123.45'], ['Price', '1.5'], ['Price', '12345'], ['Bar', 'after'],
    ['a', 'b', 'c', 'd', 'e'], ['in', 'dented', 'text'], ['1.', 'Hanging'], ['One', 'Two', 'Three'], ['Cleared', 'stop'],
    ['Right', '44'], ['wide', 'Wide', 'x'], ['Long', '1'], ['Absolute', '99'], ['Bordered', 'after2'], ['Wrap', 'the', 'End'],
]

pdf, margin = sys.argv[1], float(sys.argv[2]) if len(sys.argv) > 2 else 72.0
page = fitz.open(pdf)[0]
# characters with their x, line by line (by baseline): leaders run into the text after them, so words do not separate it
lines = {}
for block in page.get_text('rawdict')['blocks']:
    for line in block.get('lines', []):
        for span in line['spans']:
            for char in span['chars']:
                lines.setdefault(round(span['origin'][1]), []).append((char['origin'][0], char['c']))
ordered = [sorted(chars) for _, chars in sorted(lines.items())]
# the last paragraph wraps: its following lines belong to it (the text after its second tab is on the next line)
if len(ordered) > len(SEGMENTS):
    ordered = ordered[:len(SEGMENTS) - 1] + [sum(ordered[len(SEGMENTS) - 1:], [])]


def find(text, chars, segment, start):
    """The first occurrence of segment at or after start on a boundary: not after a letter, or after a gap wider than a
    character (a tab drawn without a space character)."""
    at = start
    while (at := text.find(segment, at)) >= 0:
        if at == 0 or not text[at - 1].isalnum() or chars[at][0] - chars[at - 1][0] > 9:
            return at
        at += 1
    return None


result = []
for segments, chars in zip(SEGMENTS, ordered):
    text = ''.join(c for _, c in chars)
    xs, at = [], len(segments[0])
    for segment in segments[1:]:
        found = find(text, chars, segment, at)
        if found is None:
            found = find(text, chars, segment, 0)       # a right-to-left line: the text after the tab is drawn first
        xs.append(None if found is None else round(chars[found][0] - margin, 2))
        if found is not None:
            at = found + len(segment)
    result.append({'first': segments[0], 'x': xs})
# bar stops: the vertical lines Word draws (x from the margin, in page order)
bars = [round(x - margin, 2) for _, x in sorted((min(item[1].y, item[2].y), item[1].x) for d in page.get_drawings() for item in d['items']
                                                 if item[0] == 'l' and abs(item[1].x - item[2].x) < 0.1 and abs(item[1].y - item[2].y) > 5)]
print(json.dumps({'paragraphs': result, 'bars': bars}))
