import fitz, sys
pdf = sys.argv[1]; out = sys.argv[2]
Z = 8
pix = fitz.open(pdf)[0].get_pixmap(matrix=fitz.Matrix(Z, Z))
W = pix.width; SAMPLES = pix.samples
def px(x, y):
    i = (int(y * Z) * W + int(x * Z)) * pix.n
    s = SAMPLES
    return s[i], s[i + 1], s[i + 2]
def avg(pts):
    c = [px(x, y) for x, y in pts]
    return tuple(round(sum(v[k] for v in c) / len(c)) for k in range(3))
names = ['relaxedInset','circle','slope','cross','angle','softRound','convex','coolSlant','divot','riblet','hardEdge','artDeco']
with open(out, 'w') as f:
    for i, name in enumerate(names):
        x0 = 50 + (i % 3) * 180; y0 = 70 + (i // 3) * 170
        for k in range(48):
            d = (k + 0.5) / 2
            off = [-8, -4, 0, 4, 8]
            L = avg([(x0 + d, y0 + 55 + o) for o in off]); R = avg([(x0 + 150 - d, y0 + 55 + o) for o in off])
            T = avg([(x0 + 75 + o, y0 + d) for o in off]); B = avg([(x0 + 75 + o, y0 + 110 - d) for o in off])
            f.write('%s %.2f %s\n' % (name, d / 24, ' '.join('%02x%02x%02x' % c for c in (L, R, T, B))))
        f.write('%s centre %s\n' % (name, '%02x%02x%02x' % px(x0 + 75, y0 + 55)))
