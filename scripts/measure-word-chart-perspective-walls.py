import fitz, numpy as np, json, sys, glob, os
folder = sys.argv[1]
probes = json.load(open(os.path.join(folder, 'word-charts-perspective-v1.json')))
out = []
def hull(pts):
    pts = sorted(set(map(tuple, pts)))
    def cross(o, a, b): return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])
    lo, up = [], []
    for p in pts:
        while len(lo) >= 2 and cross(lo[-2], lo[-1], p) <= 0: lo.pop()
        lo.append(p)
    for p in reversed(pts):
        while len(up) >= 2 and cross(up[-2], up[-1], p) <= 0: up.pop()
        up.append(p)
    return np.array(lo[:-1] + up[:-1], float)
def quad(mask):
    if mask.sum() < 30: return None
    inner = mask.copy(); inner[1:-1, 1:-1] = mask[1:-1, 1:-1] & mask[:-2, 1:-1] & mask[2:, 1:-1] & mask[1:-1, :-2] & mask[1:-1, 2:]
    ys, xs = np.nonzero(mask & ~inner)
    # pixel corners as the region's outline
    P = np.concatenate([np.stack([xs+dx, ys+dy], 1) for dx in (0, 1) for dy in (0, 1)])
    H = hull(P)
    # drop nearly collinear hull points, then the largest-area quadrilateral on the rest
    keep = True
    while keep and len(H) > 4:
        keep = False
        n = len(H)
        for i in range(n):
            a, b, c = H[i-1], H[i], H[(i+1) % n]
            if abs((b[0]-a[0])*(c[1]-a[1]) - (b[1]-a[1])*(c[0]-a[0])) < 2.0:
                H = np.delete(H, i, 0); keep = True; break
    n = len(H)
    if n <= 4: return [list(map(float, h)) for h in H]
    def area(i, j, k): 
        return np.abs((H[j,0]-H[i,0])*(H[k,1]-H[i,1]) - (H[j,1]-H[i,1])*(H[k,0]-H[i,0])) / 2
    best = (-1, None)
    idx = np.arange(n)
    for i in range(n):
        for k in range(i + 2, n):
            js = idx[i+1:k]; ls = np.concatenate([idx[k+1:], idx[:i]])
            if len(js) == 0 or len(ls) == 0: continue
            a1 = area(i, js, k).max() if True else 0
            a2 = area(k, ls, i).max()
            if a1 + a2 > best[0]: best = (a1 + a2, (i, js[area(i, js, k).argmax()], k, ls[area(k, ls, i).argmax()]))
    H = [H[t] for t in best[1]]
    return [list(map(float, h)) for h in H]
k = 0
for f in sorted(glob.glob(os.path.join(folder, 'word-charts-perspective-v1-*.pdf'))):
    d = fitz.open(f)
    for page in d:
        infos = sorted(page.get_image_info(xrefs=True), key=lambda i: i['bbox'][1])
        for info in infos:
            if k >= len(probes): break
            pm = fitz.Pixmap(d, info['xref'])
            smask = d.extract_image(info['xref']).get('smask', 0)
            a = np.frombuffer(pm.samples, np.uint8).reshape(pm.height, pm.width, pm.n).astype(int)
            alpha = np.frombuffer(fitz.Pixmap(d, smask).samples, np.uint8).reshape(pm.height, pm.width) if smask else np.full((pm.height, pm.width), 255)
            rgb = a[:, :, :3]
            dom = rgb.argmax(2); sat = rgb.max(2) - rgb.min(2)
            x0, y0, x1, y1 = info['bbox']
            sx, sy = (x1 - x0) / pm.width, (y1 - y0) / pm.height
            faces = {}
            for name, c in (('floor', 0), ('back', 1), ('side', 2)):
                m = (alpha > 127) & (dom == c) & (sat > 40)
                q = quad(m)
                if q: faces[name] = [[x0 + px * sx, y0 + py * sy] for px, py in q]
            # scene picture bounds (alpha)
            ys, xs = np.nonzero(alpha > 127)
            p = dict(probes[k]); p['image'] = [x0, y0, x1, y1]; p['faces'] = faces
            p['alphaBox'] = [x0 + xs.min() * sx, y0 + ys.min() * sy, x0 + (xs.max() + 1) * sx, y0 + (ys.max() + 1) * sy] if len(xs) else None
            # chart frame: the drawing rect around the image
            for dr in page.get_drawings():
                r = dr['rect']
                if r.x0 <= x0 and r.y0 <= y0 and r.x1 >= x1 and r.y1 >= y1: p['chart'] = [r.x0, r.y0, r.x1, r.y1]
            out.append(p); k += 1
json.dump(out, open(sys.argv[2], 'w'), indent=0)
for p in out: print(p['name'], {n: [[round(v, 1) for v in c] for c in q] for n, q in p['faces'].items()})
