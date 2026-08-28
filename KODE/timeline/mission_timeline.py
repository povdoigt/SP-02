"""
mission_timeline.py — génération graphique de timelines de mission.

Modèle
------
Tout instant est un *intervalle* [lo, hi]. Un événement certain est un
intervalle de largeur nulle. Une référence temporelle peut être absolue ou
ancrée sur un autre noeud (début ou fin), avec un décalage lui-même
éventuellement incertain. La résolution parcourt le graphe de dépendances et
propage l'incertitude : un événement déclaré ponctuel mais ancré sur une
fenêtre devient de fait une fenêtre.

Deux couches indépendantes :
  - le scénario théorique (prévol), toujours affichable seul ;
  - un vol réalisé optionnel, superposé en masque (vert = dans l'enveloppe,
    rouge = hors enveloppe).
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Union, Dict, List, Tuple

import matplotlib.pyplot as plt
import matplotlib.transforms as mtransforms
from matplotlib.patches import Rectangle
from matplotlib.lines import Line2D
from matplotlib.colors import to_rgba, to_hex


# ---------------------------------------------------------------------------
# Algèbre d'intervalles
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Interval:
    """Ensemble des instants possibles [lo, hi]."""
    lo: float
    hi: float

    def __post_init__(self):
        if self.hi < self.lo:
            raise ValueError(f"Intervalle inversé : [{self.lo}, {self.hi}]")

    @property
    def is_point(self) -> bool:
        return self.hi == self.lo

    @property
    def width(self) -> float:
        return self.hi - self.lo

    def shifted(self, lo: float, hi: float) -> "Interval":
        """Translation par un décalage lui-même incertain [lo, hi]."""
        return Interval(self.lo + lo, self.hi + hi)

    def contains(self, t: float) -> bool:
        return self.lo <= t <= self.hi

    def __repr__(self) -> str:
        if self.is_point:
            return f"{self.lo:g}s"
        return f"[{self.lo:g} .. {self.hi:g}]s"


# ---------------------------------------------------------------------------
# Références temporelles
# ---------------------------------------------------------------------------

EDGE_START = "start"
EDGE_END = "end"


@dataclass(frozen=True)
class TimeRef:
    """Référence temporelle.

    - absolue  : anchor is None, l'instant vaut [lo, hi]
    - relative : instant = (bord `edge` du noeud `anchor`) + [lo, hi]

    `edge` est obligatoire dès que l'ancre est une fenêtre : « fin de propulsion
    + 5 s » n'a pas le même sens selon qu'on s'accroche à l'ouverture ou à la
    fermeture de la fenêtre.
    """
    lo: float = 0.0
    hi: Optional[float] = None          # None => décalage certain (hi = lo)
    anchor: Optional[str] = None
    edge: str = EDGE_END

    @property
    def span(self) -> Tuple[float, float]:
        return (self.lo, self.lo if self.hi is None else self.hi)


# Constructeurs lisibles ----------------------------------------------------

def at(t: float) -> TimeRef:
    """Instant absolu certain."""
    return TimeRef(lo=t)


def between(t0: float, t1: float) -> TimeRef:
    """Instant absolu incertain, quelque part dans [t0, t1]."""
    return TimeRef(lo=t0, hi=t1)


def after(anchor: str, d: float = 0.0, *, edge: str = EDGE_END) -> TimeRef:
    """Ancre + décalage fixe."""
    return TimeRef(lo=d, anchor=anchor, edge=edge)


def after_between(anchor: str, d0: float, d1: float,
                  *, edge: str = EDGE_END) -> TimeRef:
    """Ancre + décalage incertain dans [d0, d1]."""
    return TimeRef(lo=d0, hi=d1, anchor=anchor, edge=edge)


# ---------------------------------------------------------------------------
# Noeuds de timeline
# ---------------------------------------------------------------------------

@dataclass
class TimelineNode:
    """Base commune : identifiant, entité porteuse, piste, libellé affiché.

    `track` est la sous-ligne de l'entité. Deux noeuds de la même piste
    partagent une ligne : c'est ce qu'on veut pour des phases de vol qui se
    succèdent. Un actionneur, lui, prend sa propre piste — une ligne
    presque vide est une information en soi (« il ne bouge qu'une fois »).
    """
    id: str
    entity: str
    label: str
    track: Optional[str] = None

    @property
    def track_name(self) -> str:
        return self.track or "Phases"

    def refs(self) -> Tuple[TimeRef, TimeRef]:
        raise NotImplementedError

    def resolve_with(self, eval_ref) -> "Resolved":
        """Cas général : les deux bords s'évaluent indépendamment."""
        r0, r1 = self.refs()
        start = eval_ref(r0)
        end = eval_ref(r1) if r1 is not r0 else start
        return Resolved(start, end)


@dataclass
class Event(TimelineNode):
    """Événement ponctuel (déclaré). Peut devenir une fenêtre par propagation."""
    when: TimeRef = field(default_factory=TimeRef)

    def refs(self) -> Tuple[TimeRef, TimeRef]:
        return (self.when, self.when)


@dataclass
class Window(TimelineNode):
    """Fenêtre d'activation : ouverture et fermeture ancrées indépendamment."""
    opens: TimeRef = field(default_factory=TimeRef)
    closes: TimeRef = field(default_factory=TimeRef)

    def refs(self) -> Tuple[TimeRef, TimeRef]:
        return (self.opens, self.closes)


@dataclass
class Entity:
    """Objet physique portant une piste : fusée complète, étage 1, étage 2...

    Sa durée de vie est elle-même une paire de références : l'étage 2 « naît »
    à la séparation, la fusée complète « meurt » au même instant.
    """
    id: str
    label: str
    born: TimeRef = field(default_factory=lambda: at(0.0))
    dies: Optional[TimeRef] = None
    color: Any = None
    """Couleur de la piste. Trois formes acceptées :

    - ``None``      : palette maison, indexée sur l'ordre de déclaration
    - ``"auto"``    : cycle de couleurs courant de matplotlib (C0, C1, ...)
    - explicite     : ``"#B4762C"``, ``"#B4762C99"``, ``"tab:blue"``,
                      ``[0.7, 0.46, 0.17]`` ou ``[180, 118, 44, 0.6]``

    Le canal alpha, s'il est fourni, n'est pas une opacité de remplissage mais
    une **atténuation** : il multiplie toutes les opacités de la piste. C'est
    ce qu'on veut pour faire reculer une entité secondaire sans changer sa
    teinte. Le libellé d'entité ne descend jamais sous 0.45 pour rester
    lisible.
    """


# ---------------------------------------------------------------------------
# Résultat de résolution
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Resolved:
    start: Interval
    end: Interval

    @property
    def envelope(self) -> Interval:
        """Enveloppe totale : au plus tôt possible .. au plus tard possible."""
        return Interval(self.start.lo, self.end.hi)

    @property
    def core(self) -> Optional[Interval]:
        """Partie garantie active, ou None si l'incertitude la dévore.

        `envelope` dit « ça se passe quelque part là-dedans », `core` dit
        « à cet endroit-là c'est certain ». Quand les deux coïncident, la
        fenêtre est franche ; quand `core` est None, on ne peut plus rien
        garantir et seule l'enveloppe a du sens.
        """
        if self.start.hi <= self.end.lo:
            return Interval(self.start.hi, self.end.lo)
        return None

    @property
    def is_certain(self) -> bool:
        return self.start.is_point and self.end.is_point

    @property
    def is_true_point(self) -> bool:
        return self.is_certain and self.start.lo == self.end.lo


# ---------------------------------------------------------------------------
# Vol réalisé
# ---------------------------------------------------------------------------

Measured = Union[float, Tuple[float, float]]


@dataclass
class Flight:
    """Vol réellement observé : simple table id_de_noeud -> instant mesuré.

    Une valeur scalaire pour un événement, un couple (début, fin) pour une
    fenêtre effectivement observée.
    """
    name: str = "Vol"
    times: Dict[str, Measured] = field(default_factory=dict)
    id: Optional[str] = None
    scenario_ref: Optional[str] = None   # chemin du scénario lié
    meta: Dict[str, Any] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Scénario
# ---------------------------------------------------------------------------

PALETTE = ["#2E6F9E", "#B4762C", "#4B7B52", "#8A4B7D", "#8C6D3F"]
DEAD, DEAD_A = "#7C838C", 0.15          # zone où l'entité n'existe pas
FLIGHT_OK, FLIGHT_KO = "#1E9E4A", "#D62728"


def resolve_color(spec: Any, index: int) -> Tuple[str, float]:
    """Normalise une spécification de couleur en (teinte opaque, atténuation)."""
    if spec is None:
        return PALETTE[index % len(PALETTE)], 1.0
    if isinstance(spec, str):
        if spec.strip().lower() == "auto":
            cycle = plt.rcParams["axes.prop_cycle"].by_key().get("color")
            cycle = cycle or PALETTE
            return cycle[index % len(cycle)], 1.0
        r, g, b, a = to_rgba(spec)
        return to_hex((r, g, b)), a
    if isinstance(spec, (list, tuple)):
        vals = [float(v) for v in spec]
        if len(vals) not in (3, 4):
            raise ValueError(f"couleur = 3 ou 4 composantes, reçu {len(vals)}")
        rgb = vals[:3]
        # Tolère aussi bien 0..1 que 0..255 sur les trois premières valeurs.
        if any(v > 1.0 for v in rgb):
            rgb = [v / 255.0 for v in rgb]
        if any(not 0.0 <= v <= 1.0 for v in rgb):
            raise ValueError(f"composantes RGB hors bornes : {spec}")
        alpha = vals[3] if len(vals) == 4 else 1.0
        if not 0.0 <= alpha <= 1.0:
            raise ValueError(f"atténuation hors [0, 1] : {alpha}")
        return to_hex(tuple(rgb)), alpha
    raise ValueError(f"couleur illisible : {spec!r}")


class Scenario:
    """Un scénario de vol : des entités, des noeuds, un rendu."""

    def __init__(self, name: str, *, id: Optional[str] = None,
                 source: Optional[str] = None):
        self.name = name
        self.id = id
        self.source = source
        self._entities: List[Entity] = []
        self._nodes: Dict[str, TimelineNode] = {}
        self._order: List[str] = []
        self._lifetimes: Dict[str, Window] = {}

    # -- construction -------------------------------------------------------

    def add_entity(self, entity: Entity) -> Entity:
        if any(e.id == entity.id for e in self._entities):
            raise ValueError(f"Entité déjà déclarée : {entity.id}")
        self._entities.append(entity)
        # La durée de vie est un noeud comme un autre : on peut donc ancrer un
        # événement sur la naissance d'un étage.
        self._lifetimes[entity.id] = Window(
            id=entity.id, entity=entity.id, label=entity.label,
            opens=entity.born,
            closes=entity.dies if entity.dies is not None else TimeRef(lo=float("inf")),
        )
        return entity

    def add(self, node: TimelineNode) -> TimelineNode:
        if node.id in self._nodes or node.id in self._lifetimes:
            raise ValueError(f"Identifiant déjà utilisé : {node.id}")
        if not any(e.id == node.entity for e in self._entities):
            raise ValueError(f"Entité inconnue : {node.entity}")
        self._nodes[node.id] = node
        self._order.append(node.id)
        return node

    # -- résolution du graphe ----------------------------------------------

    def _lookup(self, nid: str) -> TimelineNode:
        if nid in self._nodes:
            return self._nodes[nid]
        if nid in self._lifetimes:
            return self._lifetimes[nid]
        raise KeyError(f"Ancre inconnue : {nid}")

    def resolve(self) -> Dict[str, Resolved]:
        """Résout tous les noeuds en intervalles absolus (mémoïsé, anti-cycle)."""
        cache: Dict[str, Resolved] = {}
        visiting: List[str] = []

        def eval_ref(ref: TimeRef) -> Interval:
            lo, hi = ref.span
            if ref.anchor is None:
                return Interval(lo, hi)
            parent = solve(ref.anchor)
            base = parent.start if ref.edge == EDGE_START else parent.end
            return base.shifted(lo, hi)

        def solve(nid: str) -> Resolved:
            if nid in cache:
                return cache[nid]
            if nid in visiting:
                cycle = " -> ".join(visiting + [nid])
                raise ValueError(f"Cycle de dépendances : {cycle}")
            visiting.append(nid)
            node = self._lookup(nid)
            res = node.resolve_with(eval_ref)
            visiting.pop()
            cache[nid] = res
            return res

        for nid in list(self._nodes) + list(self._lifetimes):
            solve(nid)
        return cache

    def node_ids(self) -> List[str]:
        return list(self._order)

    def validate(self) -> None:
        """Résout le graphe pour faire remonter tôt cycles et ancres inconnues."""
        try:
            self.resolve()
        except (KeyError, ValueError) as exc:
            raise MissionFileError(
                f"{self.source or self.name} : scénario incohérent — {exc}"
            ) from exc

    def lint(self) -> List[str]:
        """Avertissements non bloquants sur la cohérence du scénario.

        Le cas qui compte : un noeud porté par une entité qui n'est pas encore
        née (ou déjà morte). C'est presque toujours le signe qu'on a rattaché
        l'action au mauvais corps — typiquement une commande envoyée avant la
        séparation mais déclarée sur un étage qui n'existe pas encore.
        """
        res = self.resolve()
        out: List[str] = []
        for nid in self._order:
            node = self._nodes[nid]
            life, r = res[node.entity], res[nid]
            ent = next(e for e in self._entities if e.id == node.entity)
            if r.start.lo < life.start.lo:
                out.append(
                    f"{node.label} ({nid}) commence à {r.start} alors que "
                    f"« {ent.label} » ne naît qu'à {life.start}")
            if life.end.hi != float("inf") and r.end.hi > life.end.hi:
                out.append(
                    f"{node.label} ({nid}) finit à {r.end} alors que "
                    f"« {ent.label} » meurt à {life.end}")
        return out

    def layout(self) -> List[Tuple[Entity, List[Tuple[str, List[str]]]]]:
        """Entité -> pistes -> noeuds, dans l'ordre de déclaration."""
        out = []
        for ent in self._entities:
            tracks: List[Tuple[str, List[str]]] = []
            index: Dict[str, int] = {}
            for nid in self._order:
                node = self._nodes[nid]
                if node.entity != ent.id:
                    continue
                tname = node.track_name
                if tname not in index:
                    index[tname] = len(tracks)
                    tracks.append((tname, []))
                tracks[index[tname]][1].append(nid)
            if tracks:
                out.append((ent, tracks))
        return out

    # -- vérification d'un vol ---------------------------------------------

    def _confront(self, r: Resolved,
                  measured: Measured) -> Tuple[bool, Optional[Resolved]]:
        """Compare une mesure au prévol, et reconstruit le déroulé observé."""
        if isinstance(measured, tuple):
            t0, t1 = measured
            real = Resolved(Interval(t0, t0), Interval(t1, t1))
            return (r.start.contains(t0) and r.end.contains(t1), real)
        return (r.envelope.contains(measured), None)

    def check(self, flight: Flight) -> List[str]:
        """Retourne la liste des écarts entre le vol réalisé et le prévol."""
        res = self.resolve()
        anomalies: List[str] = []
        for nid, measured in flight.times.items():
            if nid not in res:
                anomalies.append(f"{nid} : noeud inconnu dans le scénario")
                continue
            r = res[nid]
            node = self._lookup(nid)
            ok, _ = self._confront(r, measured)
            if isinstance(measured, tuple):
                shown = f"{measured[0]:g}..{measured[1]:g}s"
                attendu = f"début {r.start}, fin {r.end}"
            else:
                shown = f"{measured:g}s"
                attendu = f"enveloppe prévue {r.envelope}"
            if not ok:
                anomalies.append(
                    f"{node.label} : réalisé à {shown}, {attendu}"
                )
        return anomalies

    # -- rendu --------------------------------------------------------------

    @staticmethod
    def _label_widths(labels, fig_w, x0, x1, left, right, fontsize):
        """Largeur de chaque libellé, en unités de données.

        Mesurée sur une figure jetable de même largeur et mêmes marges : la
        conversion pixels -> données en x ne dépend que de la largeur et de
        l'échelle horizontale, jamais de la hauteur. On peut donc mesurer
        avant de savoir combien de lignes de libellés seront nécessaires,
        et éviter la circularité.
        """
        fig = plt.figure(figsize=(fig_w, 2.0))
        ax = fig.add_axes([left, 0.1, right - left, 0.8])
        ax.set_xlim(x0, x1)
        fig.canvas.draw()
        rend = fig.canvas.get_renderer()
        inv = ax.transData.inverted()
        out = []
        for lab in labels:
            t = ax.text(0, 0, lab, fontsize=fontsize)
            bb = t.get_window_extent(renderer=rend)
            (xa, _), (xb, _) = inv.transform([(0.0, 0.0), (bb.width, 0.0)])
            out.append(abs(xb - xa))
        plt.close(fig)
        return out

    @staticmethod
    def _pack_lanes(spans):
        """Range des segments [xl, xr] sur le moins de lignes possible.

        Chaque libellé descend sur la première ligne où il ne heurte aucun
        libellé déjà posé ; s'il n'en trouve aucune, on en ouvre une de plus.
        Le nombre de lignes obtenu vaut pour toutes les pistes, sinon les
        hauteurs de ligne seraient irrégulières d'une entité à l'autre.
        """
        lanes: List[List[Tuple[float, float]]] = []
        assign = [0] * len(spans)
        for idx in sorted(range(len(spans)), key=lambda i: spans[i][0]):
            xl, xr = spans[idx]
            for li, occ in enumerate(lanes):
                if all(xr <= a or xl >= b for a, b in occ):
                    occ.append((xl, xr))
                    assign[idx] = li
                    break
            else:
                lanes.append([(xl, xr)])
                assign[idx] = len(lanes) - 1
        return assign, max(len(lanes), 1)

    def plot(self, flight: Optional[Flight] = None, *,
             ax=None, figsize=(13, None), margin: float = 0.06):
        res = self.resolve()
        lay = self.layout()
        LEFT, RIGHT, FS = 0.21, 0.97, 8.0

        finite = [res[n].envelope for n in self._nodes]
        t_min = min((i.lo for i in finite), default=0.0)
        t_max = max((i.hi for i in finite), default=1.0)
        span = max(t_max - t_min, 1e-6)
        x0, x1 = t_min - margin * span, t_max + 2.2 * margin * span

        rows: List[Tuple[Entity, str, List[str]]] = []
        for ent, tracks in lay:
            for tname, ids in tracks:
                rows.append((ent, tname, ids))
        n_rows = len(rows)
        row_of = {nid: r for r, (_, _, ids) in enumerate(rows) for nid in ids}

        # -- placement des libellés -------------------------------------------
        order = [nid for _, _, ids in rows for nid in ids]
        widths = dict(zip(order, self._label_widths(
            [self._nodes[n].label for n in order],
            figsize[0], x0, x1, LEFT, RIGHT, FS)))
        pad = 0.012 * span

        span_of: Dict[str, Tuple[float, float]] = {}
        for nid in order:
            env = res[nid].envelope
            w = widths[nid] + 2 * pad
            xl = (env.lo + env.hi) / 2 - w / 2
            xl = min(max(xl, x0), x1 - w)          # jamais hors cadre
            span_of[nid] = (xl, xl + w)

        lane_of: Dict[str, int] = {}
        n_lanes = 1
        for _, _, ids in rows:
            assign, used = self._pack_lanes([span_of[n] for n in ids])
            lane_of.update(zip(ids, assign))
            n_lanes = max(n_lanes, used)

        # -- géométrie verticale ----------------------------------------------
        if flight is not None:
            BAR, DY, FBAR, FDY = 0.24, 0.11, 0.14, -0.17
            below = 0.30
        else:
            BAR, DY, FBAR, FDY = 0.30, 0.0, 0.14, 0.0
            below = BAR / 2 + 0.05
        LANE_H = 0.28
        above = 0.05 + n_lanes * LANE_H
        PITCH = above + below + 0.05

        def yb(nid):
            return -row_of[nid] * PITCH

        if ax is None:
            h = figsize[1] or (0.62 * n_rows * PITCH + 0.32 * len(lay) + 1.9)
            fig, ax = plt.subplots(figsize=(figsize[0], h))
        else:
            fig = ax.figure
        blended = mtransforms.blended_transform_factory(ax.transAxes, ax.transData)

        def fade(xa, xb, bot, top, a_from, a_to, n=72):
            if xb <= xa:
                return
            w = (xb - xa) / n
            for i in range(n):
                a = a_from + (a_to - a_from) * (i + 0.5) / n
                if a > 0.002:
                    ax.add_patch(Rectangle((xa + i * w, bot), w, top - bot,
                                           facecolor=DEAD, alpha=a,
                                           edgecolor="none", zorder=1.2))

        def slab(xa, xb, bot, top, a):
            if xb > xa:
                ax.add_patch(Rectangle((xa, bot), xb - xa, top - bot,
                                       facecolor=DEAD, alpha=a,
                                       edgecolor="none", zorder=1.2))

        # -- fonds d'entité et frontières de vie -------------------------------
        r = 0
        for k, (ent, tracks) in enumerate(lay):
            color, atten = resolve_color(ent.color, k)
            top = -r * PITCH + above + 0.02
            bot = -(r + len(tracks) - 1) * PITCH - below - 0.02

            life = res[ent.id]
            b, d = life.start, life.end
            fx0 = max(b.hi, x0)
            fx1 = min(d.lo, x1) if d.lo != float("inf") else x1
            if fx1 > fx0:
                ax.add_patch(Rectangle((fx0, bot), fx1 - fx0, top - bot,
                                       facecolor=color, alpha=0.11 * atten,
                                       edgecolor="none", zorder=0))
            slab(x0, max(b.lo, x0), bot, top, DEAD_A)
            if b.width > 0:
                fade(b.lo, b.hi, bot, top, DEAD_A, 0.0)
            if d.hi != float("inf"):
                if d.width > 0:
                    fade(d.lo, d.hi, bot, top, 0.0, DEAD_A)
                slab(min(d.hi, x1), x1, bot, top, DEAD_A)

            for iv in (b, d):
                if iv.lo == float("inf") or iv.hi == float("inf"):
                    continue
                if iv.width == 0:
                    if x0 <= iv.lo <= x1:
                        ax.plot([iv.lo, iv.lo], [bot, top], color=color,
                                lw=2.0, alpha=atten, zorder=2.4,
                                solid_capstyle="butt")
                else:
                    for xb_ in (iv.lo, iv.hi):
                        if x0 <= xb_ <= x1:
                            ax.plot([xb_, xb_], [bot, top], color=color,
                                    lw=1.1, ls=(0, (3, 2)),
                                    alpha=0.85 * atten, zorder=2.4)

            ax.axhline(bot, color="0.78", lw=0.9, zorder=0.5)
            ax.add_patch(Rectangle((-0.165, bot + 0.05), 0.008,
                                   (top - bot) - 0.10, transform=blended,
                                   facecolor=color, edgecolor="none",
                                   alpha=atten, clip_on=False, zorder=3))
            ax.text(-0.178, (top + bot) / 2, ent.label, transform=blended,
                    rotation=90, ha="center", va="center", fontweight="bold",
                    color=color, alpha=max(atten, 0.45),
                    fontsize=10 if len(tracks) >= 3 else
                             (9 if len(tracks) == 2 else 8))
            r += len(tracks)

        # -- noeuds -------------------------------------------------------------
        for k, (ent, tracks) in enumerate(lay):
            color, atten = resolve_color(ent.color, k)
            for _, ids in tracks:
                for nid in ids:
                    rr = res[nid]
                    base = yb(nid)
                    yy = base + DY
                    env, core = rr.envelope, rr.core

                    if rr.is_true_point:
                        ax.plot([env.lo, env.lo], [yy - BAR / 2, yy + BAR / 2],
                                color=color, lw=2.0, alpha=atten, zorder=4,
                                solid_capstyle="butt")
                        ax.plot([env.lo], [yy], marker="D", ms=5.0,
                                color=color, alpha=atten, zorder=4.1)
                    else:
                        # Le hachuré ne couvre que ce que le coeur ne garantit
                        # pas : sinon les traits transparaissent sous le coeur,
                        # qui n'est pas opaque.
                        if core is None:
                            gaps = [(env.lo, env.hi)]
                        else:
                            gaps = [(env.lo, core.lo), (core.hi, env.hi)]
                        for ga, gb in gaps:
                            if gb > ga:
                                ax.add_patch(Rectangle(
                                    (ga, yy - BAR / 2), gb - ga, BAR,
                                    facecolor="none", edgecolor=color,
                                    hatch="///", lw=0.9, alpha=0.55 * atten,
                                    zorder=2))
                        if core is not None and core.width > 0:
                            ax.add_patch(Rectangle(
                                (core.lo, yy - BAR / 2), core.width, BAR,
                                facecolor=color, edgecolor=color,
                                alpha=0.85 * atten, zorder=3))

                    lane = lane_of[nid]
                    xl, xr = span_of[nid]
                    ly = base + BAR / 2 + 0.06 + lane * LANE_H
                    if lane > 0:
                        xc = min(max((env.lo + env.hi) / 2, xl + pad), xr - pad)
                        ax.plot([xc, xc], [base + BAR / 2 + 0.02, ly - 0.015],
                                color="0.78", lw=0.6, zorder=1.6)
                    ax.text(xl + pad, ly, self._nodes[nid].label, va="bottom",
                            ha="left", fontsize=FS, color="0.2", zorder=5)

        # -- masque « vol réalisé » ---------------------------------------------
        if flight is not None:
            for nid, measured in flight.times.items():
                if nid not in row_of:
                    continue
                rr = res[nid]
                yy = yb(nid) + FDY
                ok, real = self._confront(rr, measured)
                c = FLIGHT_OK if ok else FLIGHT_KO
                if real is not None:
                    env, core = real.envelope, real.core
                    if core is not None and core.width > 0:
                        ax.add_patch(Rectangle(
                            (core.lo, yy - FBAR / 2), core.width, FBAR,
                            facecolor=c, edgecolor="none", zorder=6))
                    pts = (env.lo, env.hi)
                else:
                    pts = (measured,)
                for t in pts:
                    ax.plot([t, t], [yy - FBAR / 2 - 0.07, yy + FBAR / 2 + 0.07],
                            color=c, lw=2.0, solid_capstyle="butt", zorder=6.1)
                ax.text(pts[-1] + 0.006 * span, yy, f"{pts[-1]:g}", color=c,
                        fontsize=7, ha="left", va="center", zorder=6.1)

        # -- axes ---------------------------------------------------------------
        ax.axvline(0.0, color="0.4", lw=1.0, ls="--", zorder=1)
        ax.set_xlim(x0, x1)
        ax.set_ylim(-(n_rows - 1) * PITCH - below - 0.08, above + 0.08)
        ax.set_yticks([-i * PITCH for i in range(n_rows)])
        ax.set_yticklabels([t for _, t, _ in rows], fontsize=9)
        ax.tick_params(axis="y", length=0)
        ax.set_xlabel("Temps depuis le décollage (s)")
        ax.set_title(self.name + (f"  —  masque : {flight.name}"
                                  if flight is not None else ""),
                     fontweight="bold", loc="left")
        ax.grid(axis="x", color="0.9", lw=0.7)
        ax.set_axisbelow(True)
        for side in ("top", "right", "left"):
            ax.spines[side].set_visible(False)

        handles = [
            Line2D([], [], marker="D", ls="none", color="0.35",
                   label="Instant certain"),
            Rectangle((0, 0), 1, 1, facecolor="0.35", alpha=0.85,
                      label="Coeur (garanti)"),
            Rectangle((0, 0), 1, 1, facecolor="none", edgecolor="0.35",
                      hatch="///", label="Enveloppe (incertain)"),
            Rectangle((0, 0), 1, 1, facecolor=DEAD, alpha=0.38,
                      label="Entité inexistante"),
        ]
        if flight is not None:
            handles += [
                Line2D([], [], color=FLIGHT_OK, lw=2.4,
                       label="Réalisé dans l'enveloppe"),
                Line2D([], [], color=FLIGHT_KO, lw=2.4,
                       label="Réalisé hors enveloppe"),
            ]
        # Marges et décalage de légende exprimés en pouces plutôt qu'en
        # fractions : sinon une figure courte (peu de pistes, beaucoup de
        # lignes de libellés) fait retomber la légende sur l'axe des temps.
        fig_h = fig.get_size_inches()[1]
        bottom = min(0.30, 1.05 / fig_h)
        top = 1.0 - min(0.16, 0.52 / fig_h)
        axes_h = fig_h * (top - bottom)
        ax.legend(handles=handles, loc="upper center", fontsize=8,
                  bbox_to_anchor=(0.5, max(-0.45, -0.62 / axes_h)),
                  frameon=False, ncol=6)
        fig.subplots_adjust(left=LEFT, right=RIGHT, bottom=bottom, top=top)
        return fig, ax

# ---------------------------------------------------------------------------
# Couche JSON : chargement d'un scénario et d'un vol
# ---------------------------------------------------------------------------

class MissionFileError(Exception):
    """Erreur de lecture ou de validation d'un fichier de mission."""


_NUM = r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?"
_ABS_RE = re.compile(rf"^\s*({_NUM})\s*(?:\.\.\s*({_NUM})\s*)?$")
_REL_RE = re.compile(
    rf"^\s*([A-Za-z_:][A-Za-z0-9_:]*)\.(start|end)\s*(?:([+-])\s*(.+?)\s*)?$"
)


def _require(obj: dict, key: str, ctx: str):
    if key not in obj:
        raise MissionFileError(f"{ctx} : champ obligatoire manquant « {key} »")
    return obj[key]


def _parse_bounds(text: str, ctx: str) -> Tuple[float, Optional[float]]:
    m = _ABS_RE.match(text)
    if not m:
        raise MissionFileError(f"{ctx} : valeur temporelle illisible « {text} »")
    lo = float(m.group(1))
    hi = float(m.group(2)) if m.group(2) is not None else None
    return lo, hi


def parse_timeref(spec: Any, ctx: str = "?") -> TimeRef:
    """Convertit une référence temporelle JSON en `TimeRef`.

    Formes acceptées :
      6.2                        instant absolu certain
      [20, 26]                   instant absolu incertain
      "6.2"                      idem, sous forme texte
      "5.8..6.4"                 intervalle absolu
      "prop1.end"                ancré sur la fermeture de prop1
      "prop1.end + 0.5"          ancre + décalage fixe
      "prop1.end + 0.5..1.2"     ancre + décalage incertain
      "apo.start - 2.0"          décalage négatif
      {"anchor": "prop1", "edge": "end", "offset": [0.5, 1.2]}   forme explicite

    Le bord (`.start` / `.end`) est obligatoire dès qu'il y a une ancre : c'est
    lui qui lève l'ambiguïté « fin de propulsion + 5 s ».
    """
    if isinstance(spec, bool):
        raise MissionFileError(f"{ctx} : booléen inattendu")

    if isinstance(spec, (int, float)):
        return at(float(spec))

    if isinstance(spec, list):
        if len(spec) != 2:
            raise MissionFileError(f"{ctx} : un intervalle attend 2 valeurs, "
                                   f"{len(spec)} reçue(s)")
        return between(float(spec[0]), float(spec[1]))

    if isinstance(spec, dict):
        anchor = spec.get("anchor")
        if anchor is None:
            if "t" in spec:
                return at(float(spec["t"]))
            if "range" in spec:
                r = spec["range"]
                return between(float(r[0]), float(r[1]))
            raise MissionFileError(
                f"{ctx} : forme objet sans « anchor », « t » ni « range »")
        edge = spec.get("edge", EDGE_END)
        if edge not in (EDGE_START, EDGE_END):
            raise MissionFileError(
                f"{ctx} : bord invalide « {edge} » (attendu start ou end)")
        off = spec.get("offset", 0.0)
        if isinstance(off, list):
            return after_between(anchor, float(off[0]), float(off[1]), edge=edge)
        return after(anchor, float(off), edge=edge)

    if isinstance(spec, str):
        text = spec.strip()
        m = _REL_RE.match(text)
        if m:
            anchor, edge, sign, off = m.groups()
            if off is None:
                return after(anchor, 0.0, edge=edge)
            lo, hi = _parse_bounds(off, ctx)
            if hi is None:
                hi = lo
            if sign == "-":
                lo, hi = -hi, -lo
            return TimeRef(lo=lo, hi=hi, anchor=anchor, edge=edge)
        lo, hi = _parse_bounds(text, ctx)
        return TimeRef(lo=lo, hi=hi)

    raise MissionFileError(
        f"{ctx} : type de référence non supporté ({type(spec).__name__})")


def _collect_nodes(data: dict, entities: List[dict], source: str):
    """Énumère (contexte, entité, piste, dict de noeud).

    La forme canonique reproduit exactement la hiérarchie du rendu, donc ni
    « entity » ni « track » n'ont à être répétés sur les noeuds :

        {"entities": [{"id": "s1",
                       "tracks": [{"name": "Vol balistique",
                                   "nodes": [...]}]}]}

    Deux écritures antérieures restent lues : les noeuds directement sous
    l'entité (piste portée par le champ « track »), et le tableau « nodes »
    au niveau racine (entité portée par le champ « entity »).
    """
    known = {e.get("id") for e in entities}

    def guard(ctx: str, n: dict, field: str, expected: str):
        got = n.get(field)
        if got is not None and got != expected:
            raise MissionFileError(
                f"{ctx} : imbriqué sous « {expected} » mais déclarant "
                f"« {got} » — retirer le champ « {field} »")

    for e in entities:
        eid = e.get("id")
        for ti, tr in enumerate(e.get("tracks", [])):
            if not isinstance(tr, dict):
                raise MissionFileError(
                    f"{source} entities[{eid}].tracks[{ti}] : objet attendu")
            tname = tr.get("name") or tr.get("label")
            if not tname:
                raise MissionFileError(
                    f"{source} entities[{eid}].tracks[{ti}] : « name » manquant")
            for i, n in enumerate(tr.get("nodes", [])):
                ctx = f"{source} entities[{eid}].tracks[{tname}].nodes[{i}]"
                guard(ctx, n, "entity", eid)
                guard(ctx, n, "track", tname)
                yield ctx, eid, tname, n

        for i, n in enumerate(e.get("nodes", [])):
            ctx = f"{source} entities[{eid}].nodes[{i}]"
            guard(ctx, n, "entity", eid)
            yield ctx, eid, n.get("track"), n

    for i, n in enumerate(data.get("nodes", [])):
        ctx = f"{source} nodes[{i}]"
        eid = _require(n, "entity", ctx)
        if eid not in known:
            raise MissionFileError(f"{ctx} : entité inconnue « {eid} »")
        yield ctx, eid, n.get("track"), n


def _scenario_from_dict(data: dict, source: str) -> "Scenario":
    meta = data.get("scenario", {})
    if not isinstance(meta, dict):
        raise MissionFileError(f"{source} : « scenario » doit être un objet")
    sid = meta.get("id")
    name = meta.get("name") or sid or "Scénario"
    sc = Scenario(name, id=sid, source=source)

    entities = data.get("entities")
    if not entities:
        raise MissionFileError(f"{source} : aucune entité déclarée")
    for i, e in enumerate(entities):
        ctx = f"{source} entities[{i}]"
        eid = _require(e, "id", ctx)
        try:
            resolve_color(e.get("color"), 0)
        except ValueError as exc:
            raise MissionFileError(f"{ctx}.color : {exc}") from exc
        sc.add_entity(Entity(
            id=eid,
            label=e.get("label", eid),
            color=e.get("color"),
            born=parse_timeref(e.get("born", 0.0), f"{ctx}.born"),
            dies=(parse_timeref(e["dies"], f"{ctx}.dies")
                  if e.get("dies") is not None else None),
        ))

    for ctx, entity, track, n in _collect_nodes(data, entities, source):
        nid = _require(n, "id", ctx)
        ntype = n.get("type")
        label = n.get("label", nid)
        if ntype == "event":
            node = Event(id=nid, entity=entity, label=label, track=track,
                         when=parse_timeref(_require(n, "when", ctx),
                                            f"{ctx}.when"))
        elif ntype == "window":
            node = Window(id=nid, entity=entity, label=label, track=track,
                          opens=parse_timeref(_require(n, "opens", ctx),
                                              f"{ctx}.opens"),
                          closes=parse_timeref(_require(n, "closes", ctx),
                                               f"{ctx}.closes"))
        else:
            raise MissionFileError(
                f"{ctx} : type inconnu « {ntype} » (attendu event ou window)")
        try:
            sc.add(node)
        except ValueError as exc:
            raise MissionFileError(f"{ctx} : {exc}") from exc

    sc.validate()
    return sc


def load_scenario(path: Union[str, Path]) -> "Scenario":
    """Charge un scénario théorique depuis un fichier JSON."""
    p = Path(path)
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise MissionFileError(f"{p} : JSON invalide — {exc}") from exc
    except OSError as exc:
        raise MissionFileError(f"{p} : lecture impossible — {exc}") from exc
    return _scenario_from_dict(data, source=p.name)


def _parse_measurement(value: Any, ctx: str) -> Measured:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    if isinstance(value, list):
        if len(value) != 2:
            raise MissionFileError(f"{ctx} : mesure de fenêtre = [début, fin]")
        return (float(value[0]), float(value[1]))
    if isinstance(value, dict):
        if "t" in value:
            return float(value["t"])
        if "start" in value and "end" in value:
            return (float(value["start"]), float(value["end"]))
        raise MissionFileError(f"{ctx} : objet sans « t » ni « start »/« end »")
    raise MissionFileError(f"{ctx} : mesure illisible ({type(value).__name__})")


def load_flight(path: Union[str, Path],
                scenario: Optional["Scenario"] = None
                ) -> Tuple["Scenario", "Flight"]:
    """Charge un vol réalisé, et avec lui le scénario auquel il est rattaché.

    Le fichier de vol référence son scénario par `flight.scenario` (chemin
    relatif au fichier de vol). Si `flight.scenario_id` est présent, il est
    vérifié contre l'identifiant déclaré dans le scénario.
    """
    p = Path(path)
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise MissionFileError(f"{p} : JSON invalide — {exc}") from exc
    except OSError as exc:
        raise MissionFileError(f"{p} : lecture impossible — {exc}") from exc

    meta = data.get("flight", {})
    ref = meta.get("scenario")

    if scenario is None:
        if not ref:
            raise MissionFileError(
                f"{p.name} : champ « flight.scenario » manquant — impossible de "
                f"savoir à quel scénario rattacher ce vol")
        sc_path = (p.parent / ref).resolve()
        if not sc_path.exists():
            raise MissionFileError(
                f"{p.name} : scénario référencé introuvable ({sc_path})")
        scenario = load_scenario(sc_path)

    want = meta.get("scenario_id")
    if want and scenario.id and want != scenario.id:
        raise MissionFileError(
            f"{p.name} : ce vol vise le scénario « {want} » mais le fichier "
            f"chargé déclare « {scenario.id} »")

    times: Dict[str, Measured] = {}
    for nid, value in (data.get("measurements") or {}).items():
        times[nid] = _parse_measurement(value, f"{p.name} measurements.{nid}")

    known = set(scenario.node_ids())
    unknown = sorted(set(times) - known)
    if unknown:
        raise MissionFileError(
            f"{p.name} : mesures sur des noeuds absents du scénario "
            f"« {scenario.name} » : {', '.join(unknown)}")

    flight = Flight(
        name=meta.get("name") or meta.get("id") or p.stem,
        times=times,
        id=meta.get("id"),
        scenario_ref=str(ref) if ref else None,
        meta={k: v for k, v in meta.items()
              if k not in ("name", "id", "scenario", "scenario_id")},
    )
    return scenario, flight


# ---------------------------------------------------------------------------
# Ligne de commande
# ---------------------------------------------------------------------------

def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="Génère la timeline d'un scénario de mission, "
                    "éventuellement masquée par un vol réalisé.")
    ap.add_argument("scenario", nargs="?",
                    help="fichier JSON de scénario (déduit du vol si omis)")
    ap.add_argument("-f", "--flight", help="fichier JSON de vol réalisé")
    ap.add_argument("-o", "--out", help="image de sortie (.png, .pdf, .svg)")
    ap.add_argument("--check", action="store_true",
                    help="n'affiche que le rapport d'écarts, sans tracé")
    ap.add_argument("--dpi", type=int, default=150)
    args = ap.parse_args(argv)

    if not args.scenario and not args.flight:
        ap.error("il faut au moins un fichier de scénario ou un fichier de vol")

    try:
        scenario = load_scenario(args.scenario) if args.scenario else None
        flight = None
        if args.flight:
            scenario, flight = load_flight(args.flight, scenario)
    except MissionFileError as exc:
        print(f"Erreur : {exc}")
        return 2

    if flight is not None:
        anomalies = scenario.check(flight)
        print(f"{flight.name} vs {scenario.name} : "
              f"{len(flight.times)} mesure(s), {len(anomalies)} écart(s)")
        for a in anomalies:
            print(f"  HORS ENVELOPPE  {a}")
        if not anomalies:
            print("  tous les instants mesurés sont dans leur enveloppe")

    if args.check:
        return 0

    fig, _ = scenario.plot(flight=flight)
    if args.out:
        fig.savefig(args.out, dpi=args.dpi, bbox_inches="tight")
        print(f"Écrit : {args.out}")
    else:
        plt.show()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
