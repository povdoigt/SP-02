#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dv_geo.py — Noyau géométrique partagé (stdlib uniquement).

Utilisé par `analyse_terrain.py` (recherche) et `genere_kml.py` (construction).
Tous les calculs se font dans un repère ENU local en mètres, centré sur le
pas de tir : projection équirectangulaire avec correction cos(latitude),
valable sans erreur significative sur quelques kilomètres.

Conventions
-----------
  * repère local : x = est, y = nord, en mètres
  * angle mathématique phi : radians, 0 = est, sens trigonométrique
  * azimut géographique  az : degrés, 0 = nord, sens horaire
"""

from __future__ import annotations

import math
from typing import List, Optional, Sequence, Tuple, Union

R_EARTH = 6371008.8  # rayon moyen WGS84, m

Pt = Tuple[float, float]
Intervalle = Tuple[float, float]          # (phi_debut, phi_fin), sens trigo
TOUT_LE_TOUR = "TOUT"                     # sentinelle : aucun azimut interdit


# =============================================================================
# Projection locale
# =============================================================================

class ENU:
    """Projection équirectangulaire locale centrée sur (lat0, lon0)."""

    def __init__(self, lat0: float, lon0: float):
        self.lat0 = lat0
        self.lon0 = lon0
        self.k = math.cos(math.radians(lat0))

    def fwd(self, lat: float, lon: float) -> Pt:
        return (math.radians(lon - self.lon0) * self.k * R_EARTH,
                math.radians(lat - self.lat0) * R_EARTH)

    def inv(self, x: float, y: float) -> Tuple[float, float]:
        return (self.lat0 + math.degrees(y / R_EARTH),
                self.lon0 + math.degrees(x / (R_EARTH * self.k)))


# =============================================================================
# Angles
# =============================================================================

def norm_angle(a: float) -> float:
    """Ramène un angle dans [0, 2*pi)."""
    return a % (2.0 * math.pi)


def math_angle(p: Pt) -> float:
    return math.atan2(p[1], p[0])


def az_to_phi(az_deg: float) -> float:
    """Azimut géographique -> angle mathématique."""
    return math.radians(90.0 - az_deg)


def phi_to_az(phi: float) -> float:
    """Angle mathématique -> azimut géographique."""
    return (90.0 - math.degrees(phi)) % 360.0


def largeur(iv: Intervalle) -> float:
    """Ouverture angulaire d'un intervalle, en radians."""
    w = norm_angle(iv[1] - iv[0])
    return 2.0 * math.pi if w == 0.0 else w


def bissectrice(iv: Intervalle) -> float:
    return iv[0] + largeur(iv) / 2.0


def dans_intervalle(phi: float, iv: Intervalle) -> bool:
    return norm_angle(phi - iv[0]) <= largeur(iv) + 1e-12


# =============================================================================
# Primitives polygonales
# =============================================================================

def cross2(ax: float, ay: float, bx: float, by: float) -> float:
    return ax * by - ay * bx


def _close(a: Pt, b: Pt, eps: float = 1e-6) -> bool:
    return abs(a[0] - b[0]) < eps and abs(a[1] - b[1]) < eps


def open_ring(poly: Sequence[Pt]) -> List[Pt]:
    """Retire le point de fermeture éventuel."""
    p = list(poly)
    return p[:-1] if len(p) >= 2 and _close(p[0], p[-1]) else p


def signed_area(poly: Sequence[Pt]) -> float:
    s = 0.0
    n = len(poly)
    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % n]
        s += x1 * y2 - x2 * y1
    return 0.5 * s


def ensure_ccw(poly: Sequence[Pt]) -> List[Pt]:
    p = list(poly)
    return p if signed_area(p) >= 0 else p[::-1]


def dedupe(poly: Sequence[Pt], eps: float = 1e-6) -> List[Pt]:
    out: List[Pt] = []
    for p in poly:
        if not out or not _close(out[-1], p, eps):
            out.append(p)
    if len(out) > 1 and _close(out[0], out[-1], eps):
        out.pop()
    return out


def point_dans_polygone(p: Pt, ring: Sequence[Pt]) -> bool:
    x, y = p
    inside = False
    n = len(ring)
    for i in range(n):
        x1, y1 = ring[i]
        x2, y2 = ring[(i + 1) % n]
        if (y1 > y) != (y2 > y):
            if x < x1 + (y - y1) * (x2 - x1) / (y2 - y1):
                inside = not inside
    return inside


def circle_ring_intersections(ring: Sequence[Pt], radius: float,
                              P: Pt = (0.0, 0.0)) -> List[Pt]:
    """Intersections du contour fermé avec le cercle (P, radius)."""
    out: List[Pt] = []
    n = len(ring)
    for i in range(n):
        x1, y1 = ring[i]
        x2, y2 = ring[(i + 1) % n]
        dx, dy = x2 - x1, y2 - y1
        fx, fy = x1 - P[0], y1 - P[1]
        a = dx * dx + dy * dy
        if a == 0.0:
            continue
        b = 2.0 * (fx * dx + fy * dy)
        c = fx * fx + fy * fy - radius * radius
        disc = b * b - 4.0 * a * c
        if disc < 0.0:
            continue
        sq = math.sqrt(disc)
        for t in ((-b - sq) / (2.0 * a), (-b + sq) / (2.0 * a)):
            if 0.0 <= t < 1.0:            # borne haute exclusive : pas de doublon
                out.append((x1 + t * dx, y1 + t * dy))
    return out


def distance_sortie(ring: Sequence[Pt], phi: float, P: Pt = (0.0, 0.0)) -> float:
    """
    Distance à laquelle la demi-droite issue de P d'angle `phi` quitte le
    polygone. C'est la fonction centrale de tout le raisonnement.
    """
    ux, uy = math.cos(phi), math.sin(phi)
    best = math.inf
    n = len(ring)
    for i in range(n):
        a = ring[i]
        b = ring[(i + 1) % n]
        ex, ey = b[0] - a[0], b[1] - a[1]
        den = cross2(ux, uy, ex, ey)
        if abs(den) < 1e-12:
            continue                       # arête parallèle au rayon
        qx, qy = a[0] - P[0], a[1] - P[1]
        t = cross2(qx, qy, ex, ey) / den   # abscisse le long du rayon
        s = cross2(qx, qy, ux, uy) / den   # abscisse le long de l'arête
        if t > 1e-9 and -1e-9 <= s <= 1.0 + 1e-9:
            best = min(best, t)
    return best


# =============================================================================
# Secteurs admissibles
# =============================================================================

def _angles_candidats(ring: Sequence[Pt], radius: float,
                      P: Pt = (0.0, 0.0)) -> List[float]:
    """
    Azimuts où d(phi) peut franchir le seuil `radius`. C'est un ensemble fini
    et exact : d ne change de côté du seuil qu'en
      * un azimut où le cercle de rayon `radius` coupe le périmètre,
      * l'azimut d'un sommet du terrain plus proche que `radius`
        (d y est discontinue).
    """
    cands: List[float] = []
    for p in circle_ring_intersections(ring, radius, P):
        cands.append(math_angle((p[0] - P[0], p[1] - P[1])))
    for v in ring:
        if math.hypot(v[0] - P[0], v[1] - P[1]) < radius:
            cands.append(math_angle((v[0] - P[0], v[1] - P[1])))
    return sorted(set(round(norm_angle(a), 9) for a in cands))


def intervalles_admissibles(ring: Sequence[Pt], radius: float,
                            P: Pt = (0.0, 0.0)
                            ) -> Union[str, List[Intervalle]]:
    """
    Tous les intervalles d'azimuts phi tels que le segment [P, P + radius*u(phi)]
    reste dans le terrain. Retourne TOUT_LE_TOUR si aucun azimut n'est interdit,
    [] si aucun n'est admissible.
    """
    ring = ensure_ccw(open_ring(ring))
    cands = _angles_candidats(ring, radius, P)
    m = len(cands)

    if m == 0:
        return TOUT_LE_TOUR if distance_sortie(ring, 0.0, P) >= radius else []

    ok: List[bool] = []
    for i in range(m):
        a, b = cands[i], cands[(i + 1) % m]
        w = norm_angle(b - a) or 2.0 * math.pi
        ok.append(distance_sortie(ring, a + w / 2.0, P) >= radius)

    if all(ok):
        return TOUT_LE_TOUR
    if not any(ok):
        return []

    # Fusion cyclique des séquences d'intervalles admissibles
    depart = next(i for i in range(m) if not ok[i])
    res: List[Intervalle] = []
    i, vus = (depart + 1) % m, 0
    while vus < m:
        if ok[i]:
            j = i
            while ok[j]:
                j = (j + 1) % m
                vus += 1
            res.append((cands[i], cands[j]))
            i = j
        else:
            i = (i + 1) % m
            vus += 1
    return res


def sous_intervalles_hors_terrain(ring: Sequence[Pt], radius: float,
                                  phi_a: float, phi_b: float,
                                  P: Pt = (0.0, 0.0),
                                  min_deg: float = 0.01) -> List[Intervalle]:
    """
    Portions de [phi_a, phi_b] (sens trigo) où le tir sortirait du terrain
    avant `radius`. Sert au générateur pour signaler un débordement sans
    refuser de dessiner.

    `min_deg` écarte les slivers dus aux arrondis : un azimut recopié depuis
    un rapport (217.4° pour 217.44°) décale le secteur de quelques centièmes
    de degré et ferait apparaître un débordement de largeur nulle.
    """
    ring = ensure_ccw(open_ring(ring))
    span = norm_angle(phi_b - phi_a) or 2.0 * math.pi

    rel = [0.0, span]
    for a in _angles_candidats(ring, radius, P):
        r = norm_angle(a - phi_a)
        if 0.0 < r < span:
            rel.append(r)
    rel = sorted(set(round(r, 9) for r in rel))

    mauvais: List[Intervalle] = []
    for i in range(len(rel) - 1):
        mid = phi_a + (rel[i] + rel[i + 1]) / 2.0
        if distance_sortie(ring, mid, P) < radius:
            debut, fin = phi_a + rel[i], phi_a + rel[i + 1]
            if mauvais and abs(norm_angle(mauvais[-1][1] - debut)) < 1e-9:
                mauvais[-1] = (mauvais[-1][0], fin)
            else:
                mauvais.append((debut, fin))
    seuil = math.radians(min_deg)
    return [iv for iv in mauvais if largeur(iv) > seuil]


# =============================================================================
# Découpage du terrain par un secteur
# =============================================================================

def clip_halfplane(poly: Sequence[Pt], p0: Pt, d: Pt, sign: float) -> List[Pt]:
    """Sutherland-Hodgman : conserve les X tels que sign*cross(d, X-p0) >= 0."""
    def f(X: Pt) -> float:
        return sign * cross2(d[0], d[1], X[0] - p0[0], X[1] - p0[1])

    out: List[Pt] = []
    n = len(poly)
    for i in range(n):
        cur, nxt = poly[i], poly[(i + 1) % n]
        fc, fn = f(cur), f(nxt)
        if fc >= 0.0:
            out.append(cur)
        if (fc > 0.0 and fn < 0.0) or (fc < 0.0 and fn > 0.0):
            t = fc / (fc - fn)
            out.append((cur[0] + t * (nxt[0] - cur[0]),
                        cur[1] + t * (nxt[1] - cur[1])))
    return out


def _rotate_to(poly: List[Pt], pt: Pt, eps: float = 1e-6) -> List[Pt]:
    for i, p in enumerate(poly):
        if _close(p, pt, eps):
            return poly[i:] + poly[:i]
    raise ValueError("point d'ancrage absent du polygone")


def clip_secteur(ring: Sequence[Pt], P: Pt, phi_a: float, phi_b: float,
                 span_max: float = math.radians(170.0)) -> List[Pt]:
    """
    Terrain découpé par le secteur [phi_a, phi_b]. Sans limite de 180° :
    le secteur est fractionné en sous-secteurs convexes clippés séparément,
    puis recollés le long de leurs rayons de coupe communs.
    """
    ring = ensure_ccw(open_ring(ring))
    span = norm_angle(phi_b - phi_a) or 2.0 * math.pi
    n = max(1, math.ceil(span / span_max))

    morceaux: List[List[Pt]] = []
    for i in range(n):
        p1 = phi_a + span * i / n
        p2 = phi_a + span * (i + 1) / n
        m = dedupe(clip_halfplane(
            clip_halfplane(ring, P, (math.cos(p1), math.sin(p1)), +1.0),
            P, (math.cos(p2), math.sin(p2)), -1.0))
        if len(m) >= 3:
            morceaux.append(m)

    if not morceaux:
        return []

    # Chaque morceau tourné pour démarrer en P s'écrit
    # [P, (sur le rayon de début), ..., (sur le rayon de fin)].
    res = _rotate_to(morceaux[0], P)
    for m in morceaux[1:]:
        m = _rotate_to(m, P)
        suite = m[2:] if len(m) > 2 and _close(m[1], res[-1]) else m[1:]
        res = res + suite
    return dedupe(res)


# =============================================================================
# Polylignes
# =============================================================================

def arc_points(radius: float, phi_a: float, phi_b: float,
               pas_deg: float = 2.0, P: Pt = (0.0, 0.0)) -> List[Pt]:
    span = norm_angle(phi_b - phi_a) or 2.0 * math.pi
    n = max(2, int(math.ceil(math.degrees(span) / pas_deg)))
    return [(P[0] + radius * math.cos(phi_a + span * i / n),
             P[1] + radius * math.sin(phi_a + span * i / n))
            for i in range(n + 1)]


def circle_points(radius: float, n: int = 180, P: Pt = (0.0, 0.0)) -> List[Pt]:
    pts = [(P[0] + radius * math.cos(2.0 * math.pi * i / n),
            P[1] + radius * math.sin(2.0 * math.pi * i / n)) for i in range(n)]
    return pts + [pts[0]]


def dash_polyline(pts: Sequence[Pt], dash_m: float,
                  gap_m: float) -> List[List[Pt]]:
    """
    Découpe une polyligne en tirets. KML n'a pas de motif de trait : on
    matérialise les pointillés en multipliant les LineString.
    """
    if dash_m <= 0.0 or gap_m <= 0.0:
        return [list(pts)]

    period = dash_m + gap_m
    dashes: List[List[Pt]] = []
    cur: List[Pt] = []
    s = 0.0

    for i in range(len(pts) - 1):
        p, q = pts[i], pts[i + 1]
        seg = math.dist(p, q)
        if seg == 0.0:
            continue
        ux, uy = (q[0] - p[0]) / seg, (q[1] - p[1]) / seg
        t = 0.0
        while t < seg:
            u = s + t
            r = u % period
            inside = r < dash_m
            step = (dash_m - r) if inside else (period - r)
            t_end = min(t + step, seg)
            if inside:
                a = (p[0] + ux * t, p[1] + uy * t)
                b = (p[0] + ux * t_end, p[1] + uy * t_end)
                if cur and _close(cur[-1], a):
                    cur.append(b)
                else:
                    if len(cur) >= 2:
                        dashes.append(cur)
                    cur = [a, b]
            t = t_end
        s += seg

    if len(cur) >= 2:
        dashes.append(cur)
    return dashes
