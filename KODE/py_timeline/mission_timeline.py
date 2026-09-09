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
import matplotlib.patheffects as patheffects
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

# Une fenêtre résolue possède quatre bornes : son ouverture peut flotter dans
# [start.lo, start.hi] et sa fermeture dans [end.lo, end.hi]. On peut s'ancrer
# soit sur un bord entier (l'incertitude se propage), soit sur une borne
# précise (l'ancre devient un instant net).
EDGES = (EDGE_START, EDGE_END,
         "start.lo", "start.hi", "end.lo", "end.hi")


def edge_interval(r: "Resolved", edge: str) -> "Interval":
    """Extrait d'un noeud résolu l'intervalle désigné par `edge`."""
    if edge == EDGE_START:
        return r.start
    if edge == EDGE_END:
        return r.end
    side, bound = edge.split(".")
    iv = r.start if side == EDGE_START else r.end
    v = iv.lo if bound == "lo" else iv.hi
    return Interval(v, v)


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
        self.t_min: Optional[float] = None
        self.t_max: Optional[float] = None
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
            base = edge_interval(parent, ref.edge)
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
            r = res[nid]
            # Une fenêtre doit rester ordonnée bord à bord : si l'ouverture au
            # plus tard dépasse la fermeture au plus tard (ou l'ouverture au
            # plus tôt la fermeture au plus tôt), il existe des tirages où
            # elle se referme avant de s'ouvrir. L'enveloppe masque le
            # problème puisqu'elle ne retient que [start.lo, end.hi].
            if r.start.lo > r.end.lo or r.start.hi > r.end.hi:
                out.append(
                    f"{node.label} ({nid}) : ouverture {r.start} et fermeture "
                    f"{r.end} se croisent — la fenêtre peut se refermer avant "
                    f"de s'ouvrir")
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

    def anchor_instants(self) -> List[Tuple[str, str]]:
        """Couples (noeud, bord) sur lesquels au moins un autre élément s'ancre.

        Ce sont les seuls instants qui *propagent* de l'information : tracer
        une ligne pour chaque bord de chaque noeud noierait le graphique,
        alors qu'un bord dont rien ne dépend n'explique rien.
        """
        used: List[Tuple[str, str]] = []
        seen = set()
        sources: List[TimeRef] = []
        for nid in self._order:
            sources.extend(self._nodes[nid].refs())
        for ent in self._entities:
            sources.append(ent.born)
            if ent.dies is not None:
                sources.append(ent.dies)
        for ref in sources:
            if ref.anchor is None:
                continue
            key = (ref.anchor, ref.edge)
            if key not in seen:
                seen.add(key)
                used.append(key)
        return used

    def owner_of(self, nid: str) -> str:
        """Entité portant un noeud — un noeud de durée de vie se porte lui-même."""
        if nid in self._nodes:
            return self._nodes[nid].entity
        return nid

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

    @staticmethod
    def _pack_labels(items):
        """Range des libellés en tenant compte de leurs traits de rappel.

        `items` est une liste de (x_trait, boîte, texte) où « boîte » porte la
        marge anti-chevauchement et « texte » est l'emprise réelle des glyphes.
        La distinction est essentielle : un trait qui passe dans la marge de
        respiration d'un libellé ne gêne personne, alors qu'un trait qui passe
        dans ses lettres le coupe en deux. Tester la traversée sur la boîte
        padée rendait la contrainte insatisfiable dès que deux noeuds étaient
        distants de moins d'un `pad`.

        Le trait de rappel part du début de la fenêtre et monte jusqu'à la
        ligne du libellé : il traverse donc toutes les lignes en dessous.
        Trois conditions doivent tenir pour poser le libellé i sur la ligne L :
          a) sa boîte ne heurte aucune boîte déjà posée sur L ;
          b) son propre trait ne coupe le texte d'aucune ligne 0..L-1 ;
          c) aucun trait déjà posé plus haut ne coupe son texte.
        """
        n = len(items)
        lanes: List[List[Tuple[float, float]]] = []
        conns: List[Tuple[float, int]] = []
        assign = [0] * n

        texts: List[List[Tuple[float, float]]] = []

        def lane(li):
            return lanes[li] if li < len(lanes) else []

        def glyphs(li):
            return texts[li] if li < len(texts) else []

        # L'ordre de traitement n'est pas un détail : il décide si la
        # contrainte (b) est seulement satisfiable. De gauche à droite, le
        # premier libellé posé s'étend vers la droite et recouvre l'abscisse
        # de tous les traits suivants ; monter d'une ligne n'aide pas, on en
        # traverse simplement davantage, et le glouton part en butée. De
        # droite à gauche au contraire, tout libellé déjà posé commence
        # strictement à droite du trait courant, donc (b) est acquise
        # d'office. Effet de bord agréable : dans une grappe serrée, le noeud
        # le plus précoce finit en haut, et la lecture de haut en bas suit
        # l'ordre chronologique.
        for i in sorted(range(n), key=lambda j: items[j][1][0], reverse=True):
            cx, (xl, xr), (tl, tr) = items[i]
            L, ok = 0, False
            # Comparaisons strictes : un trait qui tombe pile sur le bord d'un
            # texte le touche sans le couper. C'est le cas courant de deux
            # fenêtres démarrant au même instant, et le traiter comme une
            # collision rendait la contrainte insatisfiable — aucune ligne
            # plus haute n'y échappe, puisqu'on passe devant le même bord.
            while L <= len(lanes) + len(conns):
                ok = all(xr <= a or xl >= b for a, b in lane(L))
                if ok:
                    ok = not any(a < cx < b
                                 for li in range(L) for a, b in glyphs(li))
                if ok:
                    ok = not any(L2 > L and tl < cx2 < tr
                                 for cx2, L2 in conns)
                if ok:
                    break
                L += 1
            if not ok:
                # Filet de sécurité : une ligne neuve au-dessus de tout, plutôt
                # qu'un indice arbitraire qui laisserait un trou dans la pile.
                L = len(lanes)
            while len(lanes) <= L:
                lanes.append([])
                texts.append([])
            lanes[L].append((xl, xr))
            texts[L].append((tl, tr))
            conns.append((cx, L))
            assign[i] = L
        return assign, max(len(lanes), 1)

    def plot(self, flight: Optional[Flight] = None, *,
             ax=None, figsize=(13, None), margin: float = 0.06,
             t_min: Optional[float] = None, t_max: Optional[float] = None):
        """Trace le scénario.

        `t_min` / `t_max` fixent la fenêtre temporelle affichée. Non fournis,
        ils sont déduits des noeuds avec une marge. Fournis, ils sont
        respectés au pouce près : aucune marge n'est ajoutée de ce côté-là, et
        les noeuds entièrement hors cadre sont écartés du tracé plutôt que
        rognés — un libellé isolé désignant une barre invisible serait pire
        que son absence. À défaut d'argument, les valeurs du scénario
        (bloc « view » du JSON) servent de repli.
        """
        res = self.resolve()
        lay = self.layout()
        LEFT, RIGHT, FS = 0.21, 0.97, 8.0

        if t_min is None:
            t_min = self.t_min
        if t_max is None:
            t_max = self.t_max

        finite = [res[n].envelope for n in self._nodes]
        auto_lo = min((i.lo for i in finite), default=0.0)
        auto_hi = max((i.hi for i in finite), default=1.0)
        span0 = max(auto_hi - auto_lo, 1e-6)
        x0 = auto_lo - margin * span0 if t_min is None else float(t_min)
        x1 = auto_hi + 2.2 * margin * span0 if t_max is None else float(t_max)
        if x1 <= x0:
            raise ValueError(
                f"Fenêtre temporelle vide : t_min={x0} >= t_max={x1}")

        def visible(nid: str) -> bool:
            env = res[nid].envelope
            return not (env.hi < x0 or env.lo > x1)

        rows: List[Tuple[Entity, str, List[str]]] = []
        for ent, tracks in lay:
            for tname, ids in tracks:
                rows.append((ent, tname, [n for n in ids if visible(n)]))
        n_rows = len(rows)
        row_of = {nid: r for r, (_, _, ids) in enumerate(rows) for nid in ids}

        # -- placement horizontal des libellés (empilement en k lignes) -------
        order = [nid for _, _, ids in rows for nid in ids]

        # Le cadre est élargi jusqu'à ce que le libellé le plus à droite tienne
        # sans être rabattu. Un rabattement décroche le texte de son trait, et
        # surtout il peut empiler deux libellés au même endroit : leurs traits
        # tombent alors dans le texte partagé, et aucune ligne ne permet plus
        # d'y échapper. Mieux vaut un peu de marge à droite qu'une contrainte
        # devenue insatisfiable. La largeur d'un texte en unités de données
        # dépend de l'échelle, donc on itère — ça converge vite, un libellé
        # étant bien plus étroit que l'axe entier.
        for _ in range(6):
            pad = 0.012 * (x1 - x0)
            widths = dict(zip(order, self._label_widths(
                [self._nodes[n].label for n in order],
                figsize[0], x0, x1, LEFT, RIGHT, FS)))
            if t_max is not None:
                break          # borne imposée : on ne l'élargit pas
            need = max((res[n].envelope.lo + widths[n] + 2 * pad
                        for n in order), default=x1)
            if need <= x1 + 1e-9:
                break
            x1 = need
        pad = 0.012 * (x1 - x0)
        span = x1 - x0

        span_of: Dict[str, Tuple[float, float]] = {}
        conn_of: Dict[str, float] = {}
        text_of: Dict[str, float] = {}
        glyph_of: Dict[str, Tuple[float, float]] = {}
        for nid in order:
            env = res[nid].envelope
            # Le trait de rappel est imposé au début de la fenêtre, donc le
            # libellé s'aligne là aussi : trait et texte partagent la même
            # abscisse et la liaison se lit sans ambiguïté. Si le cadre force
            # un décalage à droite, un petit segment horizontal rattrape
            # l'écart (voir le tracé plus bas).
            conn_of[nid] = env.lo
            # Position du texte et boîte de collision sont deux choses
            # distinctes : le texte s'aligne exactement sur le trait, et la
            # marge anti-chevauchement se met *autour* de la boîte. Les
            # confondre décalait le libellé d'un `pad` vers la droite, si
            # bien que le trait ne tombait plus sous sa première lettre.
            tx = min(max(env.lo, x0 + pad), x1 - widths[nid] - pad)
            text_of[nid] = tx
            span_of[nid] = (tx - pad, tx + widths[nid] + pad)
            glyph_of[nid] = (tx, tx + widths[nid])

        lane_of: Dict[str, int] = {}
        label_lanes: List[int] = []
        for _, _, ids in rows:
            assign, used = self._pack_labels(
                [(conn_of[n], span_of[n], glyph_of[n]) for n in ids])
            lane_of.update(zip(ids, assign))
            label_lanes.append(used)

        # -- empilement vertical des barres qui se chevauchent dans le temps --
        # Deux fenêtres de la même piste peuvent avoir une intersection non
        # nulle (retard de mise à feu, marge de recouvrement...). Plutôt que
        # de les dessiner l'une sur l'autre, on les répartit sur des
        # sous-lignes locales à la piste, sur le même principe glouton que
        # pour les libellés — mais packées sur le chevauchement temporel réel
        # (l'enveloppe), pas sur la largeur du texte.
        #
        # Les instants certains sont exclus du paquetage : ils occupent toute
        # la hauteur de la piste et ne consomment donc pas de sous-ligne. Un
        # événement traverse ainsi les fenêtres qui l'englobent, ce qui est
        # exactement la lecture voulue (un Max Q se lit « pendant » la
        # propulsion, pas « à côté »).
        # Le découpage se fait par grappe de chevauchement, pas par piste
        # entière : deux fenêtres qui ne se croisent jamais dans le temps
        # n'ont aucune raison de se partager la hauteur. Une piste peut donc
        # avoir une zone serrée découpée en trois, et ailleurs une fenêtre
        # isolée qui reprend toute la hauteur disponible.
        #
        # Les instants certains sont exclus du paquetage : ils occupent toute
        # la hauteur de la piste et ne consomment donc pas de sous-ligne. Un
        # événement traverse ainsi les fenêtres qui l'englobent, ce qui est
        # exactement la lecture voulue (un Max Q se lit « pendant » la
        # propulsion, pas « à côté »).
        bar_lane_of: Dict[str, int] = {}
        depth_of: Dict[str, int] = {}
        for _, _, ids in rows:
            spans_ids = [n for n in ids if not res[n].is_true_point]
            spans_ids.sort(key=lambda n: res[n].envelope.lo)
            cluster: List[str] = []
            reach = float("-inf")

            def flush(group):
                if not group:
                    return
                assign, used = self._pack_lanes(
                    [(res[n].envelope.lo, res[n].envelope.hi) for n in group])
                bar_lane_of.update(zip(group, assign))
                depth_of.update({n: max(used, 1) for n in group})

            for nid in spans_ids:
                env = res[nid].envelope
                # Nouvelle grappe dès qu'on ne touche plus rien de la
                # précédente : la portée cumulée est le maximum des fins vues.
                if cluster and env.lo >= reach:
                    flush(cluster)
                    cluster, reach = [], float("-inf")
                cluster.append(nid)
                reach = max(reach, env.hi)
            flush(cluster)

        # -- géométrie verticale ------------------------------------------------
        # Hauteur de piste constante H. À l'intérieur, chaque grappe partage
        # cette hauteur entre ses n sous-lignes : une barre fait H/n moins une
        # marge. La densité est donc mesurée là où elle existe.
        H, VMARGIN, LANE_H = 0.66, 0.07, 0.28
        above_of: List[float] = [0.05 + lanes * LANE_H for lanes in label_lanes]
        slot_of: Dict[str, float] = {n: H / depth_of[n] for n in depth_of}
        barh_of: Dict[str, float] = {n: max(0.05, slot_of[n] - VMARGIN)
                                     for n in depth_of}

        # Sommet de chaque piste. L'écart avant la piste suivante doit absorber
        # la marge « above » de CETTE piste suivante (sa propre zone de
        # libellés), et seulement la sienne : une piste qui n'a besoin que
        # d'une ligne de libellé ne doit pas hériter de l'espace nécessaire à
        # une piste voisine qui, elle, en empile cinq.
        row_top: List[float] = [0.0]
        row_bot: List[float] = []
        for r in range(n_rows):
            row_bot.append(row_top[r] - H)
            if r < n_rows - 1:
                row_top.append(row_bot[r] - above_of[r + 1] - 0.05)

        # Centre vertical de chaque noeud : milieu de sa sous-ligne pour une
        # fenêtre, milieu de la piste entière pour un instant certain.
        center_of: Dict[str, float] = {}
        for nid in order:
            r_i = row_of[nid]
            if res[nid].is_true_point:
                center_of[nid] = row_top[r_i] - H / 2
            else:
                center_of[nid] = (row_top[r_i]
                                  - (bar_lane_of[nid] + 0.5) * slot_of[nid])

        def split(bh):
            """Répartit l'épaisseur disponible entre prévu et réalisé."""
            if flight is None:
                return bh, 0.0
            return bh * 0.56, bh * 0.32

        if ax is None:
            total = row_top[0] - row_bot[-1] + above_of[0]
            h = figsize[1] or (0.62 * total + 0.32 * len(lay) + 1.5)
            fig, ax = plt.subplots(figsize=(figsize[0], h))
        else:
            fig = ax.figure
        # Marges et limites d'axes fixées tout de suite, avant tout dessin :
        # le calibrage des libellés d'entité (ci-après) mesure des tailles en
        # pixels, qui ne sont justes que si la mise en page finale de la figure
        # est déjà en place (subplots_adjust déplacé ici depuis la fin).
        fig_h = fig.get_size_inches()[1]
        bottom = min(0.30, 1.05 / fig_h)
        top = 1.0 - min(0.16, 0.52 / fig_h)
        fig.subplots_adjust(left=LEFT, right=RIGHT, bottom=bottom, top=top)
        ax.set_xlim(x0, x1)
        ax.set_ylim(row_bot[-1] - 0.08, row_top[0] + above_of[0] + 0.08)
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
            r0, r1 = r, r + len(tracks) - 1
            box_top = row_top[r0] + above_of[r0] + 0.02
            box_bot = row_bot[r1] - 0.02

            life = res[ent.id]
            b, d = life.start, life.end
            fx0 = max(b.hi, x0)
            fx1 = min(d.lo, x1) if d.lo != float("inf") else x1
            if fx1 > fx0:
                ax.add_patch(Rectangle((fx0, box_bot), fx1 - fx0, box_top - box_bot,
                                       facecolor=color, alpha=0.11 * atten,
                                       edgecolor="none", zorder=0))
            slab(x0, max(b.lo, x0), box_bot, box_top, DEAD_A)
            if b.width > 0:
                fade(b.lo, b.hi, box_bot, box_top, DEAD_A, 0.0)
            if d.hi != float("inf"):
                if d.width > 0:
                    fade(d.lo, d.hi, box_bot, box_top, 0.0, DEAD_A)
                slab(min(d.hi, x1), x1, box_bot, box_top, DEAD_A)

            for iv in (b, d):
                if iv.lo == float("inf") or iv.hi == float("inf"):
                    continue
                if iv.width == 0:
                    if x0 <= iv.lo <= x1:
                        ax.plot([iv.lo, iv.lo], [box_bot, box_top], color=color,
                                lw=2.0, alpha=atten, zorder=2.4,
                                solid_capstyle="butt")
                else:
                    for xb_ in (iv.lo, iv.hi):
                        if x0 <= xb_ <= x1:
                            ax.plot([xb_, xb_], [box_bot, box_top], color=color,
                                    lw=1.1, ls=(0, (3, 2)),
                                    alpha=0.85 * atten, zorder=2.4)

            ax.axhline(box_bot, color="0.78", lw=0.9, zorder=0.5)
            ax.add_patch(Rectangle((-0.165, box_bot + 0.05), 0.008,
                                   (box_top - box_bot) - 0.10, transform=blended,
                                   facecolor=color, edgecolor="none",
                                   alpha=atten, clip_on=False, zorder=3))
            box_px = abs(ax.transData.transform((0, box_top))[1]
                        - ax.transData.transform((0, box_bot))[1])
            start_fs = 10 if len(tracks) >= 3 else (9 if len(tracks) == 2 else 8)
            fs = start_fs
            fig.canvas.draw()
            renderer = fig.canvas.get_renderer()
            probe = ax.text(0, 0, ent.label, fontsize=fs)
            need_px = probe.get_window_extent(renderer=renderer).width
            probe.remove()
            if need_px > 0.72 * box_px:
                fs = max(6.0, fs * 0.72 * box_px / need_px)
            ax.text(-0.178, (box_top + box_bot) / 2, ent.label, transform=blended,
                    rotation=90, ha="center", va="center", fontweight="bold",
                    color=color, alpha=max(atten, 0.45), fontsize=fs)
            r += len(tracks)

        # -- noeuds -------------------------------------------------------------
        for k, (ent, tracks) in enumerate(lay):
            color, atten = resolve_color(ent.color, k)
            for _, ids in tracks:
                for nid in ids:
                    if nid not in row_of:
                        continue          # hors de la fenêtre imposée
                    rr = res[nid]
                    r_i = row_of[nid]
                    c_y = center_of[nid]
                    env, core = rr.envelope, rr.core

                    if rr.is_true_point:
                        # Instant certain : toute la hauteur de la piste.
                        top_y, bot_y = row_top[r_i], row_bot[r_i]
                        # Liseré noir : un trait fin de la couleur de l'entité
                        # se perd sur le fond teinté de cette même entité.
                        ax.plot([env.lo, env.lo], [bot_y, top_y],
                                color=color, lw=1.9, alpha=atten, zorder=4,
                                solid_capstyle="butt",
                                path_effects=[patheffects.withStroke(
                                    linewidth=3.1, foreground="#111111")])
                        ax.plot([env.lo], [c_y], marker="D", ms=5.2,
                                color=color, alpha=atten, zorder=4.1,
                                markeredgecolor="#111111", markeredgewidth=0.7)
                        conn_top = top_y
                    else:
                        bh_i = barh_of[nid]
                        th, _fh = split(bh_i)
                        yy = c_y + (bh_i - th) / 2
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
                                    (ga, yy - th / 2), gb - ga, th,
                                    facecolor="none", edgecolor=color,
                                    hatch="///", lw=0.9, alpha=0.55 * atten,
                                    zorder=2))
                        if core is not None and core.width > 0:
                            ax.add_patch(Rectangle(
                                (core.lo, yy - th / 2), core.width, th,
                                facecolor=color, edgecolor=color,
                                alpha=0.85 * atten, zorder=3))
                        conn_top = yy + th / 2

                    # Le libellé se range dans la zone partagée au-dessus de la
                    # piste, jamais juste au-dessus de sa propre barre : sinon
                    # le texte d'une barre en sous-ligne 1+ tomberait sur la
                    # barre qui la précède dans l'empilement.
                    lane = lane_of[nid]
                    tx = text_of[nid]
                    ly = row_top[r_i] + 0.06 + lane * LANE_H
                    cx = conn_of[nid]
                    # Trait de rappel : vertical, au début de la fenêtre, et
                    # aligné sur la première lettre du libellé.
                    ax.plot([cx, cx], [conn_top + 0.02, ly - 0.012],
                            color="0.62", lw=1.0, zorder=1.6,
                            solid_capstyle="butt")
                    if abs(tx - cx) > 1e-9:
                        # Le cadre a repoussé le texte : petit segment de
                        # liaison pour que le trait reste rattaché au libellé.
                        ax.plot([cx, tx], [ly - 0.012, ly - 0.012],
                                color="0.62", lw=1.0, zorder=1.6,
                                solid_capstyle="butt")
                    ax.text(tx, ly, self._nodes[nid].label, va="bottom",
                            ha="left", fontsize=FS, color="0.2", zorder=5)

        # -- masque « vol réalisé » ---------------------------------------------
        if flight is not None:
            for nid, measured in flight.times.items():
                if nid not in row_of:
                    continue
                rr = res[nid]
                r_i = row_of[nid]
                bh = H if rr.is_true_point else barh_of[nid]
                th, fh = split(bh)
                fh = max(fh, 0.05)
                yy = center_of[nid] - (bh - fh) / 2
                ok, real = self._confront(rr, measured)
                c = FLIGHT_OK if ok else FLIGHT_KO
                if real is not None:
                    env, core = real.envelope, real.core
                    if core is not None and core.width > 0:
                        ax.add_patch(Rectangle(
                            (core.lo, yy - fh / 2), core.width, fh,
                            facecolor=c, edgecolor="none", zorder=6))
                    pts = (env.lo, env.hi)
                else:
                    pts = (measured,)
                for t in pts:
                    ax.plot([t, t], [yy - fh / 2 - 0.05, yy + fh / 2 + 0.05],
                            color=c, lw=2.0, solid_capstyle="butt", zorder=6.1)
                ax.text(pts[-1] + 0.006 * span, yy, f"{pts[-1]:g}", color=c,
                        fontsize=7, ha="left", va="center", zorder=6.1)

        # -- axes ---------------------------------------------------------------
        ax.axvline(0.0, color="0.4", lw=1.0, ls="--", zorder=1)
        ticks, labels = [], []
        for r, (_, tname, _) in enumerate(rows):
            ticks.append(row_top[r] - H / 2)
            labels.append(tname)
        ax.set_yticks(ticks)
        ax.set_yticklabels(labels, fontsize=9)
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
                   markeredgecolor="#111111", markeredgewidth=0.7,
                   label="Instant certain"),
            Rectangle((0, 0), 1, 1, facecolor="0.35", alpha=0.85,
                      label="Coeur (garanti)"),
            Rectangle((0, 0), 1, 1, facecolor="none", edgecolor="0.35",
                      hatch="///", label="Enveloppe (incertain)"),
        ]
        if flight is not None:
            handles += [
                Line2D([], [], color=FLIGHT_OK, lw=2.4,
                       label="Réalisé dans l'enveloppe"),
                Line2D([], [], color=FLIGHT_KO, lw=2.4,
                       label="Réalisé hors enveloppe"),
            ]
        # Décalage de légende exprimé en pouces : sinon une figure courte
        # (peu de pistes, beaucoup de lignes de libellés) fait retomber la
        # légende sur l'axe des temps. Marges déjà posées plus haut.
        axes_h = fig_h * (top - bottom)
        ax.legend(handles=handles, loc="upper center", fontsize=8,
                  bbox_to_anchor=(0.5, max(-0.45, -0.62 / axes_h)),
                  frameon=False, ncol=min(len(handles), 5))
        return fig, ax

# ---------------------------------------------------------------------------
# Couche JSON : chargement d'un scénario et d'un vol
# ---------------------------------------------------------------------------

class MissionFileError(Exception):
    """Erreur de lecture ou de validation d'un fichier de mission."""


_NUM = r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?"
_ABS_RE = re.compile(rf"^\s*({_NUM})\s*(?:\.\.\s*({_NUM})\s*)?$")
# Un identifiant admet « : » et « - », très courants pour hiérarchiser des ids
# (« s1:seq:airbrakes »). Le point reste réservé au séparateur de bord : un id
# contenant un point serait indistinguable de « <ancre>.<bord> ».
_IDENT = r"[A-Za-z_][A-Za-z0-9_:\-]*"
_REL_RE = re.compile(
    rf"^\s*({_IDENT})"
    rf"\.(start|end)(?:\.(lo|hi))?"
    rf"\s*(?:([+-])\s*(.+?)\s*)?$"
)

GRAMMAIRE_REF = """Une référence temporelle s'écrit sous l'une de ces formes.

  ABSOLUE — un instant lu directement sur l'axe
    6.2                    instant certain
    [5.8, 6.4]             instant incertain (liste JSON)
    "5.8..6.4"             idem, en texte

  RELATIVE — <ancre>.<bord>[ ±<décalage>]
    "prop1.end"            fermeture de prop1, incertitude comprise
    "prop1.end + 0.5"      décalage fixe
    "prop1.end + 0.5..1.2" décalage lui-même incertain
    "apo.start - 2.0"      décalage négatif

  <ancre>  identifiant d'un noeud, ou d'une entité pour sa durée de vie.
           Lettres, chiffres, « _ », « : » et « - ». Pas de point : il est
           réservé au séparateur de bord.
  <bord>   OBLIGATOIRE. « fin de propulsion + 5 s » est ambigu quand
           l'ancre est une fenêtre, donc le bord doit être nommé :
             start      ouverture, incertitude comprise
             end        fermeture, incertitude comprise
             start.lo   au plus tôt de l'ouverture   (instant net)
             start.hi   au plus tard de l'ouverture  (instant net)
             end.lo     au plus tôt de la fermeture  (instant net)
             end.hi     au plus tard de la fermeture (instant net)
           Les deux premiers propagent l'incertitude de l'ancre ; les
           quatre autres la coupent en visant une borne précise.

  OBJET — même chose en explicite, utile pour générer du JSON
    {"t": 6.2}
    {"range": [5.8, 6.4]}
    {"anchor": "prop1", "edge": "end", "offset": [0.5, 1.2]}
"""


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

    La grammaire acceptée est celle de `GRAMMAIRE_REF`, reproduite dans les
    messages d'erreur pour qu'on n'ait pas à la chercher.
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
        if edge not in EDGES:
            raise MissionFileError(
                f"{ctx} : bord invalide « {edge} ». Bords acceptés : "
                + ", ".join(EDGES))
        off = spec.get("offset", 0.0)
        if isinstance(off, list):
            return after_between(anchor, float(off[0]), float(off[1]), edge=edge)
        return after(anchor, float(off), edge=edge)

    if isinstance(spec, str):
        text = spec.strip()
        m = _REL_RE.match(text)
        if m:
            anchor, side, bound, sign, off = m.groups()
            edge = f"{side}.{bound}" if bound else side
            if off is None:
                return after(anchor, 0.0, edge=edge)
            lo, hi = _parse_bounds(off, ctx)
            if hi is None:
                hi = lo
            if sign == "-":
                lo, hi = -hi, -lo
            return TimeRef(lo=lo, hi=hi, anchor=anchor, edge=edge)
        if re.fullmatch(_IDENT, text):
            raise MissionFileError(
                f"{ctx} : « {text} » désigne une ancre sans préciser le bord. "
                f"Écrire par exemple « {text}.end » ou « {text}.start.lo ».\n"
                + GRAMMAIRE_REF)
        if "." in text and not _ABS_RE.match(text):
            # Ressemble à une référence relative mais ne colle pas au motif :
            # c'est presque toujours un bord oublié ou mal orthographié.
            raise MissionFileError(
                f"{ctx} : référence relative illisible « {text} ».\n"
                + GRAMMAIRE_REF)
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
    view = data.get("view") or {}
    if not isinstance(view, dict):
        raise MissionFileError(f"{source} view : objet attendu")
    for key in ("t_min", "t_max"):
        if view.get(key) is not None:
            try:
                setattr(sc, key, float(view[key]))
            except (TypeError, ValueError):
                raise MissionFileError(
                    f"{source} view.{key} : nombre attendu, reçu "
                    f"{view[key]!r}") from None
    if (sc.t_min is not None and sc.t_max is not None
            and sc.t_max <= sc.t_min):
        raise MissionFileError(
            f"{source} view : t_max ({sc.t_max}) doit dépasser "
            f"t_min ({sc.t_min})")

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
    ap.add_argument("--t-min", type=float, default=None,
                    dest="t_min", metavar="T",
                    help="borne gauche imposée de l'axe des temps (s)")
    ap.add_argument("--t-max", type=float, default=None,
                    dest="t_max", metavar="T",
                    help="borne droite imposée de l'axe des temps (s)")
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

    for w in scenario.lint():
        print(f"  ATTENTION  {w}")

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

    fig, _ = scenario.plot(flight=flight,
                           t_min=args.t_min, t_max=args.t_max)
    if args.out:
        fig.savefig(args.out, dpi=args.dpi, bbox_inches="tight")
        print(f"Écrit : {args.out}")
    else:
        plt.show()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())