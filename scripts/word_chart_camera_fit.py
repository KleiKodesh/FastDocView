import numpy as np, json, itertools, math
FACES = {
    'floor': [(0, 0, 0), (1, 0, 0), (1, 0, 1), (0, 0, 1)],
    'back': [(0, 0, 1), (1, 0, 1), (1, 1, 1), (0, 1, 1)],
    'side': [(0, 0, 0), (0, 0, 1), (0, 1, 1), (0, 1, 0)],
}
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

def residuals(probe, raw_fn, sim=None):
    """raw_fn(unit corners)-> raw 2-D. sim = (s, tx, ty) maps raw (y up) to page; None = least-squares similarity."""
    models, meas = [], []
    for name, q in probe['faces'].items():
        models.append(raw_fn(FACES[name])); meas.append(np.array(q))
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
    Q = np.concatenate([match(p, q) for p, q in zip(P, meas)])
    return (np.concatenate(P) - Q).ravel(), sim

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
