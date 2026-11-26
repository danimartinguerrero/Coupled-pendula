"""Propuestas por defecto: primer máximo estacionario (n0) y picos anómalos.

n0: la cola más larga (≥ kmin máximos) cuya deriva relativa
    δ = (dφ/dt)·(N−1)·T/β cumple |δ| ≤ tol, igual que el criterio del libro.
"""
import numpy as np

def deriva(t, h):
    n = np.arange(len(t), dtype=float)
    T = np.polyfit(n, t, 1)[0]
    b = np.polyfit(t, h, 1)[0]
    return b * (len(t) - 1) * T / h.mean()

def anomalos(h, rel=0.25):
    out = set()
    N = len(h)
    for i in range(1, N):
        ref = (h[i - 1] + h[i + 1]) / 2 if i < N - 1 else (h[i - 1] + h[i - 2]) / 2
        if abs(h[i] - ref) > rel * ref:
            out.add(i + 1)
    return out

def propone(t, h, tol=0.10, kmin=5):
    """Devuelve (n0, excluidos, ok)."""
    N = len(h)
    an = anomalos(h)
    for k in range(N, kmin - 1, -1):
        n0 = N - k + 1
        idx = [i for i in range(n0 - 1, N) if (i + 1) not in an]
        if len(idx) >= kmin and abs(deriva(t[idx], h[idx])) <= tol:
            return n0, {n for n in an if n >= n0}, True
    n0 = N - kmin + 1
    return n0, {n for n in an if n >= n0}, False
