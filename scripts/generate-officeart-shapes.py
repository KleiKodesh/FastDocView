"""Generates FastDocView/Html/WordOfficeArtShapeDefinitions.cs from LibreOffice's OfficeArt shape tables.

Inputs (LibreOffice core master, downloaded by the caller):
  svx/source/customshapes/EnhancedCustomShapeGeometry.cxx   (the mso_CustomShape tables and GetCustomShapeContent)
  include/svx/msdffdef.hxx                                   (enum MSO_SPT, DFF_Prop_* numbers)
  svx/source/customshapes/EnhancedCustomShapeTypeNames.cxx  (pNameTypeTableArray: the ODF draw:type names)
Usage: python scripts/generate-officeart-shapes.py <EnhancedCustomShapeGeometry.cxx> <msdffdef.hxx> <EnhancedCustomShapeTypeNames.cxx> <output .cs>
Each table is copied as data: vertices (an "n MSO_I" coordinate = equation n), segment commands, calculation records,
default adjust values, text rectangles, the coordinate size and the stretch points (nXRef / nYRef); and the ODF type
names. Nothing is interpreted here; WordOfficeArtShapeTypes evaluates them.
"""
import re
import sys

geometry, definitions, typenames, output = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
geo = open(geometry, encoding='utf-8').read()
defs = open(definitions, encoding='utf-8').read()

constants = {name: int(value) for name, value in re.findall(r'#define\s+(DFF_Prop_\w+)\s+(\d+)', defs)}
enum_body = defs[defs.index('enum MSO_SPT'):]
enum_body = enum_body[:enum_body.index('};')]
shape_types = {name: int(value) for name, value in re.findall(r'(mso_spt\w+)\s*=\s*(\d+)', enum_body)}
geo = re.sub(r'//[^\n]*', '', geo)
geo = re.sub(r'/\*.*?\*/', '', geo, flags=re.S)
constants['MIN_INT32'] = -2147483648


def value(token):
    token = token.strip()
    token = re.sub(r'^(-?(?:0x[0-9a-fA-F]+|\d+))[lLuU]+$', r'\1', token)      # C literal suffixes (4l)
    special = re.fullmatch(r'(-?\w+)\s+MSO_I', token)
    if special:
        return ('eq', value(special.group(1))[1])
    if token in constants:
        return ('n', constants[token])
    if re.fullmatch(r'-?0x[0-9a-fA-F]+', token):
        return ('n', int(token, 16))
    if re.fullmatch(r'-?\d+', token):
        return ('n', int(token))
    # a constant expression (e.g. static_cast<sal_Int32>((21600 - ((M_SQRT2-1)*21600)) / 2)): evaluated as C would
    expression = re.sub(r'static_cast<\w+>', 'int', token).replace('M_SQRT2', 'math.sqrt(2)').replace('M_PI', 'math.pi')
    if re.fullmatch(r'[0-9a-z_.+\-*/() ]+', expression):
        import math
        return ('n', int(eval(expression, {'int': int, 'math': math})))
    raise ValueError('unknown token ' + token)


def array(kind, name):
    match = re.search(r'const\s+' + kind + r'\s+' + re.escape(name) + r'\s*\[\s*\]\s*=\s*\{(.*?)\};', geo, re.S)
    if not match:
        raise ValueError('no array ' + name)
    return match.group(1)


def numbers(text):
    return [value(t)[1] for t in re.split(r',', text.replace('{', ' ').replace('}', ' ')) if t.strip()]


def vertices(name):
    if name is None:
        return []
    result = []
    for a, b in re.findall(r'\{\s*([^,{}]+?)\s*,\s*([^,{}]+?)\s*\}', array('SvxMSDffVertPair', name)):
        for token in (a, b):
            kind, number = value(token)
            result.append(number | 0x40000000 if kind == 'eq' else number)
    return result


def text_rects(name):
    if name is None:
        return []
    result = []
    pair = r'\{\s*([^,{}]+?)\s*,\s*([^,{}]+?)\s*\}'
    for a, b, c, d in re.findall(r'\{\s*' + pair + r'\s*,\s*' + pair + r'\s*\}', array('SvxMSDffTextRectangles', name)):
        for token in (a, b, c, d):
            kind, number = value(token)
            result.append(number | 0x40000000 if kind == 'eq' else number)
    return result


def calculations(name):
    if name is None:
        return []
    result = []
    for flags, p1, p2, p3 in re.findall(r'\{\s*([^,{}]+)\s*,\s*\{\s*([^,{}]+)\s*,\s*([^,{}]+)\s*,\s*([^,{}]+)\s*\}\s*\}', array('SvxMSDffCalculationData', name)):
        result += [value(flags)[1], value(p1)[1], value(p2)[1], value(p3)[1]]
    return result


def span(text):
    match = re.match(r'\s*std::span<const \w+>\(\s*(\w*)\s*\)', text)
    return match.group(1) or None if match else None


cases = re.findall(r'case\s+(mso_spt\w+)\s*:\s*pCustomShape\s*=\s*&(\w+)\s*;', geo)
entries = []
for spt, shape in cases:
    body = re.search(r'const\s+mso_CustomShape\s+' + shape + r'\s*=\s*\{(.*?)\};', geo, re.S).group(1)
    parts = []
    depth = 0
    current = ''
    for character in body:
        if character == '(':
            depth += 1
        if character == ')':
            depth -= 1
        if character == ',' and depth == 0:
            parts.append(current)
            current = ''
        else:
            current += character
    parts.append(current)
    vert, segm, calc, default, rects = (span(p) for p in parts[:5])
    width, height = value(parts[5])[1], value(parts[6])[1]
    xref, yref = value(parts[7])[1], value(parts[8])[1]
    segments = numbers(array('sal_uInt16', segm)) if segm else []
    defaults = numbers(array('sal_Int32', default))[1:] if default else []
    entries.append((shape_types[spt], spt, shape, width, height, vertices(vert), segments, calculations(calc), defaults, text_rects(rects), xref, yref))

with open(output, 'w', encoding='utf-8', newline='\n') as out:
    out.write('// <auto-generated> by scripts/generate-officeart-shapes.py from LibreOffice core master\n')
    out.write('// svx/source/customshapes/EnhancedCustomShapeGeometry.cxx (GetCustomShapeContent and its mso_CustomShape tables)\n')
    out.write('// and include/svx/msdffdef.hxx. Do not edit by hand; regenerate. </auto-generated>\n')
    out.write('namespace FastDocView.Html;\n\n')
    out.write('/// <summary>One OfficeArt shape type\'s outline as LibreOffice defines it: coordinate size, vertices (bit 30 set = the\n')
    out.write('/// result of equation n), segment commands, calculation records (flags, p1, p2, p3), default adjust values, text rectangles\n')
    out.write('/// (left, top, right, bottom per rectangle, marked like vertices) and the stretch points (int.MinValue = none).</summary>\n')
    out.write('internal sealed record WordOfficeArtShapeDefinition(int CoordinateWidth, int CoordinateHeight, int[] Vertices, ushort[] Segments, int[] Calculations, int[] Defaults,\n')
    out.write('                                                    int[] TextRectangles, int StretchX, int StretchY);\n\n')
    out.write('internal static class WordOfficeArtShapeDefinitions\n{\n')
    out.write('    public static WordOfficeArtShapeDefinition? Of(int shapeType) => shapeType switch\n    {\n')
    seen = set()
    for number, spt, shape, width, height, verts, segs, calcs, defaults, rects, xref, yref in entries:
        if number in seen:
            continue
        seen.add(number)
        def ints(values):
            return '[' + ', '.join(str(v) for v in values) + ']'
        out.write(f'        {number} => new({width}, {height}, {ints(verts)}, {ints(segs)}, {ints(calcs)}, {ints(defaults)}, {ints(rects)}, {xref}, {yref}),   // {spt} = {shape}\n')
    out.write('        _ => null,\n    };\n\n')
    table = open(typenames, encoding='utf-8').read().split('pNameTypeTableArray[] =')[1].split('};')[0]
    names = re.findall(r'\{\s*u"([^"]+)"_ustr\s*,\s*(mso_spt\w+)\s*\}', table)
    out.write('    /// <summary>The MSO_SPT number of an ODF draw:type name (EnhancedCustomShapeTypeNames.cxx pNameTypeTableArray); null = mso_sptNil.</summary>\n')
    out.write('    public static int? TypeOfName(string name) => name switch\n    {\n')
    seen_names = set()
    for name, spt in names:
        if name in seen_names:
            continue
        seen_names.add(name)
        out.write(f'        "{name}" => {shape_types[spt]},\n')
    out.write('        _ => null,\n    };\n\n')
    out.write('    /// <summary>The ODF draw:type name of an MSO_SPT number (EnhancedCustomShapeTypeNames::Get(MSO_SPT): the table entry at that index).</summary>\n')
    out.write('    public static string? NameOfType(int type) => type switch\n    {\n')
    for index, (name, spt) in enumerate(names):
        if shape_types[spt] == index:
            out.write(f'        {index} => "{name}",\n')
    out.write('        _ => null,\n    };\n}\n')
print(len(names), 'type names')
print(len(entries), 'shape types')
