"""Writes scripts/smartart-shaping-fixtures.json (ASCII; non-Latin text as u-escapes) for make-smartart-layouts.ps1 -Fixtures:
SmartArt block lists that check shaping against Word's own drawing part - Indic text in Nirmala UI (Devanagari reph, half
and below-base forms; Kannada below-base forms Nirmala spaces with GPOS 'dist'; Telugu conjuncts) and a family with no bold /
italic face (Gabriola, Sylfaen, Impact: Word synthesizes the style). One long word per node, so a few percent of width moves the size a point.
  python scripts/make-smartart-shaping-fixtures.py
  powershell -File scripts/make-smartart-layouts.ps1 -Fixtures scripts/smartart-shaping-fixtures.json
  python scripts/strip-smartart-drawing.py"""
import json
import pathlib


def text(*code_points):
    return ''.join(chr(c) for c in code_points)


# antarrashtriyakaran, dharmashastra, sarvashreshtha (reph, conjuncts, half forms, pre-base i matra)
DEVANAGARI = [text(0x0905, 0x0928, 0x094D, 0x0924, 0x0930, 0x094D, 0x0930, 0x093E, 0x0937, 0x094D, 0x091F, 0x094D, 0x0930, 0x0940, 0x092F, 0x0915, 0x0930, 0x0923),
              text(0x0927, 0x0930, 0x094D, 0x092E, 0x0936, 0x093E, 0x0938, 0x094D, 0x0924, 0x094D, 0x0930, 0x093F, 0x092F, 0x094B, 0x0902),
              text(0x0938, 0x0930, 0x094D, 0x0935, 0x0936, 0x094D, 0x0930, 0x0947, 0x0937, 0x094D, 0x0920, 0x0924, 0x092E)]
# Kannada words with below-base conjuncts (na-virama-na, ka-virama-ka, ta-virama-ra)
KANNADA = [text(0x0C95, 0x0CA8, 0x0CCD, 0x0CA8, 0x0CA1, 0x0CB8, 0x0C82, 0x0CB8, 0x0CCD, 0x0C95, 0x0CC3, 0x0CA4, 0x0CBF),
           text(0x0CAA, 0x0CCD, 0x0CB0, 0x0CAA, 0x0C82, 0x0C9A, 0x0CAA, 0x0CCD, 0x0CB0, 0x0CB8, 0x0CBF, 0x0CA6, 0x0CCD, 0x0CA7),
           text(0x0CB5, 0x0CBF, 0x0CA6, 0x0CCD, 0x0CAF, 0x0CBE, 0x0CB0, 0x0CCD, 0x0CA5, 0x0CBF, 0x0C97, 0x0CB3, 0x0CC1)]
# Telugu words with conjuncts (ka-ssa-ra, stra)
TELUGU = [text(0x0C15, 0x0C4D, 0x0C37, 0x0C4D, 0x0C30, 0x0C2E, 0x0C3E, 0x0C32, 0x0C15, 0x0C4D, 0x0C37, 0x0C4D, 0x0C2E, 0x0C3F),
          text(0x0C38, 0x0C4D, 0x0C24, 0x0C4D, 0x0C30, 0x0C40, 0x0C38, 0x0C02, 0x0C18, 0x0C2E, 0x0C41, 0x0C32, 0x0C24, 0x0C4B),
          text(0x0C2A, 0x0C4D, 0x0C30, 0x0C2A, 0x0C02, 0x0C1A, 0x0C35, 0x0C3F, 0x0C26, 0x0C4D, 0x0C2F, 0x0C3E, 0x0C32, 0x0C2F, 0x0C02)]
LATIN = ['Extraordinary', 'Internationalization', 'Characteristically']

fixtures = [
    {'id': 'default', 'name': 'devanagari-block-list', 'complexFont': 'Nirmala UI', 'texts': DEVANAGARI},
    {'id': 'default', 'name': 'kannada-block-list', 'complexFont': 'Nirmala UI', 'texts': KANNADA},
    {'id': 'default', 'name': 'telugu-block-list', 'complexFont': 'Nirmala UI', 'texts': TELUGU},
    {'id': 'default', 'name': 'gabriola-bold-block-list', 'bold': True, 'minorLatin': 'Gabriola', 'texts': LATIN},
    {'id': 'default', 'name': 'gabriola-italic-block-list', 'italic': True, 'minorLatin': 'Gabriola', 'texts': LATIN},
    {'id': 'default', 'name': 'gabriola-block-list', 'minorLatin': 'Gabriola', 'texts': LATIN},
    # two more families with no bold face: does Word's synthesized bold measure narrower than the regular face in general?
    {'id': 'default', 'name': 'sylfaen-bold-block-list', 'bold': True, 'minorLatin': 'Sylfaen', 'texts': LATIN},
    {'id': 'default', 'name': 'sylfaen-block-list', 'minorLatin': 'Sylfaen', 'texts': LATIN},
    {'id': 'default', 'name': 'impact-bold-block-list', 'bold': True, 'minorLatin': 'Impact', 'texts': LATIN},
    {'id': 'default', 'name': 'impact-block-list', 'minorLatin': 'Impact', 'texts': LATIN},
]
target = pathlib.Path(__file__).resolve().parent / 'smartart-shaping-fixtures.json'
target.write_text(json.dumps(fixtures, ensure_ascii=True, indent=1) + '\n', encoding='ascii')
print('wrote', target.name, len(fixtures), 'fixtures')
