import numpy as np, json, itertools, math
FACES = {
    'floor': [(0, 0, 0), (1, 0, 0), (1, 0, 1), (0, 0, 1)],
    'back': [(0, 0, 1), (1, 0, 1), (1, 1, 1), (0, 1, 1)],
    'side': [(0, 0, 0), (0, 0, 1), (0, 1, 1), (0, 1, 0)],
}
def walls_for(rotY):
    """Word moves the back wall to the front edge once rotY reaches 90 (the 2026-11-24 grid at rotY 90: back wall at z = 0
    fits 0.3-3.8 pt, at z = 1 25-80 pt); the side wall stays at x = 0 there."""
    if 90 <= rotY % 360 < 270:
        return dict(FACES, back=[(x, y, 0) for x, y, _ in FACES['back']])
    return FACES

def load(path): return json.load(open(path))

def turn(p, rotX, rotY):
    x, y, z = p[..., 0], p[..., 1], -p[..., 2]
    a, b = math.radians(rotX), math.radians(rotY)
    x1 = x * math.cos(b) - z * math.sin(b); z1 = x * math.sin(b) + z * math.cos(b)
    y2 = y * math.cos(a) - z1 * math.sin(a); z2 = y * math.sin(a) + z1 * math.cos(a)
    return np.stack([x1, y2, z2], -1)

def project(unit, W, H, D, rotX, rotY, E, c):
    """unit corners (0/1) -> raw 2-D (y up). c = centre (cx, cy, cz) in scene units; E = eye distance (inf = parallel)."""
    p = np.array(unit, float) * np.array([W, H, D]) - np.array(c)
    t = turn(p, rotX, rotY)
    f = 1.0 if not np.isfinite(E) else E / (E - t[..., 2])
    return np.stack([t[..., 0] * f, t[..., 1] * f], -1)

ORDERS = []
for s in range(4):
    ORDERS.append([(s + i) % 4 for i in range(4)]); ORDERS.append([(s - i) % 4 for i in range(4)])

def match(model, meas):
    """model 4x2, meas 4x2 -> meas reordered to best fit model (similarity not applied; positions in same frame)."""
    best = None
    for o in ORDERS:
        m = meas[o]; e = ((m - model) ** 2).sum()
        if best is None or e < best[0]: best = (e, m)
    return best[1]

def residuals(probe, raw_fn, sim=None, thin=0.0, faces=None):
    """raw_fn(unit corners)-> raw 2-D. sim = (s, tx, ty) maps raw (y up) to page; None = least-squares similarity.
    faces: the unit-corner faces to use (default FACES; see walls_for).
    thin > 0: faces thinner than that (page units) are matched by outline distance instead of corner order."""
    models, meas = [], []
    for name, q in probe['faces'].items():
        models.append(raw_fn((faces or FACES)[name])); meas.append(np.array(q))
    if not models: return None
    # first, a rough similarity for matching: least squares on centroid order-independent scale
    M = np.concatenate(models);
    def fit_sim(M, Q):
        # page = (s*x + tx, -s*y + ty)
        A = np.zeros((2 * len(M), 3)); b = np.zeros(2 * len(M))
        A[0::2, 0] = M[:, 0]; A[0::2, 1] = 1; b[0::2] = Q[:, 0]
        A[1::2, 0] = -M[:, 1]; A[1::2, 2] = 1; b[1::2] = Q[:, 1]
        return np.linalg.lstsq(A, b, rcond=None)[0]
    if sim is None:
        # iterate: match with a bbox-based similarity, then refit
        allq = np.concatenate(meas)
        s0 = (allq[:, 0].max() - allq[:, 0].min()) / max(1e-9, M[:, 0].max() - M[:, 0].min())
        sim0 = (s0, allq[:, 0].mean() - s0 * M[:, 0].mean(), allq[:, 1].mean() + s0 * M[:, 1].mean())
        for _ in range(3):
            P = [np.stack([sim0[0] * m[:, 0] + sim0[1], -sim0[0] * m[:, 1] + sim0[2]], 1) for m in models]
            Q = np.concatenate([match(p, q) for p, q in zip(P, meas)])
            sim0 = fit_sim(M, Q)
        sim = sim0
    P = [np.stack([sim[0] * m[:, 0] + sim[1], -sim[0] * m[:, 1] + sim[2]], 1) for m in models]
    if thin <= 0:
        Q = np.concatenate([match(p, q) for p, q in zip(P, meas)])
        return (np.concatenate(P) - Q).ravel(), sim
    # faces seen nearly edge-on (rotX 0 floor, rotY 90 walls): the measured quadrilateral of a sliver has its corners anywhere
    # along the long edges and can miss a thin tip (the measuring drops nearly collinear outline points), so such a face is
    # compared one way only: each measured corner's distance to the model face's outline
    out = []
    for p, q in zip(P, meas):
        if thickness(q) >= thin:  # by the measured face only, so the residual count stays fixed while fitting
            out.append((p - match(p, q)).ravel()); continue
        d = [outline_distance(b, p) for b in q]
        out.append(np.repeat(np.array(d) / math.sqrt(2), 2))  # two entries per corner, as the corner residuals
    return np.concatenate(out), sim

def thickness(quad):
    """A quadrilateral's area over its longest side: about its width across, in the same units."""
    x, y = quad[:, 0], quad[:, 1]
    area = abs(np.dot(x, np.roll(y, -1)) - np.dot(y, np.roll(x, -1))) / 2
    side = max(np.hypot(*(quad[i] - quad[i - 1])) for i in range(len(quad)))
    return area / max(side, 1e-9)

def outline_distance(a, quad):
    """Distance from point a to the closed outline through the quadrilateral's corners."""
    best = float('inf')
    for i in range(len(quad)):
        b, c = quad[i - 1], quad[i]; e = c - b; L = e @ e
        t = 0.0 if L < 1e-12 else min(1.0, max(0.0, (a - b) @ e / L))
        best = min(best, float(np.hypot(*(a - b - t * e))))
    return best

def lm(f, x0, iters=60, step=1e-5):
    x = np.array(x0, float); lam = 1e-3
    r = f(x); cost = r @ r
    for _ in range(iters):
        J = np.zeros((len(r), len(x)))
        for i in range(len(x)):
            dx = np.zeros_like(x); h = step * max(1, abs(x[i])); dx[i] = h
            J[:, i] = (f(x + dx) - r) / h
        g = J.T @ r; A = J.T @ J
        while True:
            try: d = np.linalg.solve(A + lam * np.diag(np.diag(A) + 1e-12), -g)
            except np.linalg.LinAlgError: d = -g * 1e-3
            xn = x + d; rn = f(xn); cn = rn @ rn
            if np.isfinite(cn) and cn < cost:
                x, r, cost, lam = xn, rn, cn, lam * 0.3; break
            lam *= 10
            if lam > 1e12: return x, r
    return x, r
