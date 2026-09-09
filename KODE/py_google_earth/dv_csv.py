#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dv_csv.py — Lecture des CSV de vol (stdlib uniquement).

Aucune détection automatique de format : les colonnes sont DÉCLARÉES par
l'utilisateur dans le JSON, en numérotation 1 = première colonne. Les
exports d'outils de simulation changent d'ordre et de langue d'une version
à l'autre ; deviner est une source d'erreurs silencieuses.

Toute ligne commençant par le préfixe de commentaire est ignorée — cela
couvre à la fois l'entête et les marqueurs d'événements (« # Event LIFTOFF
occurred at t=0.06 seconds »). Les événements ne sont pas exploités ici.
"""

from __future__ import annotations

import bisect
import csv
import math
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

COLONNES_DEFAUT = {"temps": 1, "altitude": 2, "distance_laterale": 7}


@dataclass
class Vol:
    """Un vol lu depuis un CSV, en unités SI, altitude relative au sol."""
    nom: str
    chemin: str
    temps: List[float] = field(default_factory=list)
    altitude: List[float] = field(default_factory=list)
    distance: List[float] = field(default_factory=list)
    lignes_lues: int = 0
    lignes_ignorees: int = 0
    lignes_invalides: int = 0

    def __len__(self) -> int:
        return len(self.altitude)

    @property
    def apogee_m(self) -> float:
        return max(self.altitude) if self.altitude else 0.0

    @property
    def i_apogee(self) -> int:
        return max(range(len(self.altitude)), key=self.altitude.__getitem__)

    @property
    def portee_m(self) -> float:
        """Distance latérale maximale atteinte (pas forcément le dernier point)."""
        return max(self.distance) if self.distance else 0.0

    @property
    def duree_s(self) -> Optional[float]:
        return (self.temps[-1] - self.temps[0]) if len(self.temps) >= 2 else None


class ErreurCSV(Exception):
    pass


def _idx(colonnes: Dict[str, int], cle: str, defaut: Optional[int]) -> Optional[int]:
    v = colonnes.get(cle, defaut)
    if v is None:
        return None
    v = int(v)
    if v < 1:
        raise ErreurCSV(f"colonne '{cle}' = {v} : la numérotation commence à 1")
    return v - 1                      # -> index 0-based


def lire_vol(chemin: str, nom: Optional[str] = None,
             colonnes: Optional[Dict[str, int]] = None,
             separateur: str = ",", commentaire: str = "#",
             tolerance_invalides: int = 20) -> Vol:
    """
    Lit un CSV de vol. `colonnes` est un dict en numérotation 1-based :
        {"temps": 1, "altitude": 2, "distance_laterale": 7}
    'temps' est facultatif.
    """
    colonnes = dict(colonnes or COLONNES_DEFAUT)
    i_t = _idx(colonnes, "temps", None)
    i_alt = _idx(colonnes, "altitude", COLONNES_DEFAUT["altitude"])
    i_dst = _idx(colonnes, "distance_laterale",
                 COLONNES_DEFAUT["distance_laterale"])
    besoin = max(x for x in (i_t, i_alt, i_dst) if x is not None) + 1

    if not os.path.isfile(chemin):
        raise ErreurCSV(f"fichier introuvable : {chemin}")

    vol = Vol(nom=nom or os.path.splitext(os.path.basename(chemin))[0],
              chemin=chemin)
    largeur_max = 0

    # newline='' + utf-8-sig : gère CRLF et un éventuel BOM
    with open(chemin, "r", encoding="utf-8-sig", newline="") as f:
        for no, brut in enumerate(f, start=1):
            s = brut.strip()
            if not s:
                continue
            if commentaire and s.startswith(commentaire):
                vol.lignes_ignorees += 1      # entête ET marqueurs d'événements
                continue

            champs = next(csv.reader([s], delimiter=separateur))
            largeur_max = max(largeur_max, len(champs))
            if len(champs) < besoin:
                vol.lignes_invalides += 1
                continue
            try:
                alt = float(champs[i_alt])
                dst = float(champs[i_dst])
                tps = float(champs[i_t]) if i_t is not None else float(no)
            except ValueError:
                vol.lignes_invalides += 1
                continue

            vol.temps.append(tps)
            vol.altitude.append(alt)
            vol.distance.append(dst)
            vol.lignes_lues += 1

    if not vol.lignes_lues:
        raise ErreurCSV(
            f"{chemin} : aucune ligne de données exploitable.\n"
            f"    {vol.lignes_ignorees} ligne(s) de commentaire, "
            f"{vol.lignes_invalides} ligne(s) rejetée(s), "
            f"{largeur_max} colonne(s) détectée(s).\n"
            f"    -> vérifie 'colonnes' (numérotation 1-based) et 'separateur'."
        )
    if vol.lignes_invalides > tolerance_invalides:
        raise ErreurCSV(
            f"{chemin} : {vol.lignes_invalides} lignes rejetées sur "
            f"{vol.lignes_invalides + vol.lignes_lues}.\n"
            f"    {largeur_max} colonne(s) détectée(s) — 'colonnes' est "
            f"probablement mal renseigné."
        )
    return vol


# =============================================================================
# Décimation
# =============================================================================

def decimer(vol: Vol, nb_points: Optional[int],
            garder_apogee: bool = True) -> Tuple[List[int], Optional[str]]:
    """
    Choisit les indices à afficher. Retourne (indices, avertissement).

      * nb_points None      -> tous les points
      * nb_points >= total  -> tous les points, avec avertissement (clippé)
      * nb_points <= 0      -> aucun point (le vol n'est pas tracé)
      * sinon               -> répartition régulière, premier et dernier
                               toujours conservés

    Si `garder_apogee`, l'indice d'altitude maximale est forcé dans la
    sélection en remplaçant l'indice retenu le plus proche : le nombre de
    points demandé est respecté, et le sommet de la trajectoire n'est pas
    perdu par l'échantillonnage.
    """
    total = len(vol)

    if nb_points is None:
        return list(range(total)), None

    n = int(nb_points)
    if n <= 0:
        return [], (f"{vol.nom} : nb_points = {n} <= 0, vol non tracé")
    if n >= total:
        avert = None if n == total else (
            f"{vol.nom} : nb_points = {n} > {total} points disponibles, "
            f"limité à {total}")
        return list(range(total)), avert
    if n == 1:
        return [vol.i_apogee if garder_apogee else 0], None

    idx = sorted({round(i * (total - 1) / (n - 1)) for i in range(n)})

    if garder_apogee:
        apo = vol.i_apogee
        if apo not in idx:
            # remplace le voisin le plus proche, sans toucher aux extrémités
            interieur = [j for j in idx if j not in (idx[0], idx[-1])]
            if interieur:
                proche = min(interieur, key=lambda j: abs(j - apo))
                idx = sorted(set(idx) - {proche} | {apo})
            else:
                idx = sorted(set(idx) | {apo})

    return idx, None


@dataclass
class Groupe:
    """Un groupe de vols — typiquement une configuration de fusée."""
    id: str
    nom: str
    reference: bool = False
    # Réglages portés par le groupe lui-même, et non hérités depuis ses vols :
    # un vol peut redéfinir sa couleur sans changer celle du groupe.
    couleur: Optional[str] = None
    pas_de_tir: Optional[str] = None
    zone: Optional[str] = None
    plafond_m: float = 0.0
    azimut_deg: Optional[float] = None      # cône imposé par l'utilisateur
    demi_angle_deg: Optional[float] = None
    domaine: Optional[dict] = None          # résolu par l'analyseur, lu par le générateur
    vols: List[Tuple[Vol, dict]] = field(default_factory=list)

    @property
    def portee_m(self) -> float:
        """
        Portée du groupe : la plus grande distance latérale atteinte parmi ses
        vols balistiques (tous ses vols si aucun n'est marqué). C'est ce qui
        dimensionne le domaine pour cette configuration de fusée.
        """
        refs = [v for v, o in self.vols if o.get("balistique")] or \
               [v for v, _ in self.vols]
        return max((v.portee_m for v in refs), default=0.0)

    @property
    def vol_dimensionnant(self) -> Optional[Vol]:
        refs = [v for v, o in self.vols if o.get("balistique")] or \
               [v for v, _ in self.vols]
        return max(refs, key=lambda v: v.portee_m, default=None)


# Réglages hérités en cascade : bloc 'vols' -> groupe -> vol.
# Un groupe peut venir d'un autre export que le précédent, et un vol d'un
# format différent du reste de son groupe : chaque niveau peut tout redéfinir.
HERITABLES = ("colonnes", "separateur", "commentaire", "nb_points",
              "couleur", "azimut_deg", "pas_de_tir", "balistique",
              "affichage")


def charger_vols(bloc: Optional[dict], base_dir: str = ".",
                 normaliser_rideau=None) -> List[Groupe]:
    """
    Charge le bloc JSON `vols` et retourne la liste des groupes.

    `groupes` est la forme normale. Une `liste` de vols à la racine du bloc est
    acceptée et produit un groupe unique implicite.
    """
    if not bloc:
        return []

    groupes_bruts = bloc.get("groupes")
    if groupes_bruts is None:
        if not bloc.get("liste"):
            return []
        groupes_bruts = [{"id": "vols", "nom": "Vols",
                          "reference": True, "liste": bloc["liste"]}]

    dossier = bloc.get("dossier", ".")
    if not os.path.isabs(dossier):
        dossier = os.path.join(base_dir, dossier)

    racine_groupe = {
        "couleur": bloc.get("couleur", "#ff3020"),
        "pas_de_tir": bloc.get("pas_de_tir"),
        "zone": bloc.get("zone"),
        "plafond_m": bloc.get("plafond_m"),
        "azimut_deg": bloc.get("azimut_deg"),
        "demi_angle_deg": bloc.get("demi_angle_deg"),
    }
    racine = {
        "colonnes": bloc.get("colonnes", COLONNES_DEFAUT),
        "separateur": bloc.get("separateur", ","),
        "commentaire": bloc.get("commentaire", "#"),
        "nb_points": bloc.get("nb_points"),        # None -> tous les points
        "couleur": bloc.get("couleur", "#ff3020"),
        "azimut_deg": bloc.get("azimut_deg"),      # None -> azimut de la vue
        "pas_de_tir": bloc.get("pas_de_tir"),      # None -> celui du groupe
        "balistique": False,
        "affichage": bloc.get("affichage"),
    }
    rideau_racine = (normaliser_rideau(bloc.get("rideau"))
                     if normaliser_rideau else bloc.get("rideau"))

    groupes: List[Groupe] = []
    for ig, gb in enumerate(groupes_bruts):
        gid = str(gb.get("id", f"groupe_{ig + 1}"))
        herite = dict(racine)
        herite.update({k: v for k, v in gb.items() if k in HERITABLES})
        rideau_groupe = (normaliser_rideau(gb.get("rideau"), rideau_racine)
                         if normaliser_rideau else rideau_racine)

        g_ = dict(racine_groupe)
        g_.update({k: v for k, v in gb.items() if k in racine_groupe})
        pl = g_["plafond_m"]
        groupe = Groupe(
            id=gid, nom=str(gb.get("nom", gid)),
            reference=bool(gb.get("reference", False)),
            couleur=g_["couleur"], pas_de_tir=g_["pas_de_tir"],
            zone=g_["zone"], plafond_m=float(pl) if pl is not None else 0.0,
            azimut_deg=(None if g_["azimut_deg"] is None
                        else float(g_["azimut_deg"])),
            demi_angle_deg=(None if g_["demi_angle_deg"] is None
                            else float(g_["demi_angle_deg"])),
            domaine=gb.get("domaine"))

        for entree in gb.get("liste", []):
            if "fichier" not in entree:
                raise ErreurCSV(
                    f"vols.groupes[{ig}].liste : entrée sans champ 'fichier'")
            opts = dict(herite)
            opts.update({k: v for k, v in entree.items()
                         if k not in ("fichier", "rideau", "domaine")})
            opts["rideau"] = (normaliser_rideau(entree.get("rideau"),
                                                rideau_groupe)
                              if normaliser_rideau else rideau_groupe)
            opts["groupe"] = gid

            chemin = entree["fichier"]
            if not os.path.isabs(chemin):
                sous = entree.get("dossier", gb.get("dossier"))
                chemin = os.path.join(dossier, sous or "", chemin)
            vol = lire_vol(chemin, nom=opts.get("nom"),
                           colonnes=opts["colonnes"],
                           separateur=opts["separateur"],
                           commentaire=opts["commentaire"])
            opts["nom"] = vol.nom
            groupe.vols.append((vol, opts))

        groupes.append(groupe)

    ids = [g.id for g in groupes]
    if len(set(ids)) != len(ids):
        raise ErreurCSV(f"identifiants de groupe en double : {ids}")
    return groupes


def choisir_groupe(groupes: List[Groupe], ident=None) -> Optional[Groupe]:
    """
    Groupe dimensionnant : identifiant explicite, puis le groupe marqué
    'reference', puis — à défaut — celui de plus grande portée. Ce dernier
    choix est le cas le plus défavorable, seul défendable pour un domaine de
    sécurité couvrant plusieurs fusées.
    """
    if not groupes:
        return None
    if ident is not None:
        for g in groupes:
            if g.id == ident:
                return g
        raise ErreurCSV(
            f"groupe '{ident}' introuvable — disponibles : "
            + ", ".join(f"'{g.id}'" for g in groupes))
    marques = [g for g in groupes if g.reference]
    if len(marques) == 1:
        return marques[0]
    if len(marques) > 1:
        raise ErreurCSV(
            "plusieurs groupes marqués 'reference' : "
            + ", ".join(f"'{g.id}'" for g in marques))
    return max(groupes, key=lambda g: g.portee_m)


# =============================================================================
# Motifs de rideau
# =============================================================================
#
# Le rideau est la surface verticale sous la trajectoire. Comme la trajectoire
# est projetée en ligne droite le long d'un azimut, cette surface se décrit
# entièrement dans le plan (d, h) : d = distance latérale, h = altitude. Tous
# les motifs sont donc construits en 2D dans ce plan, puis relevés en 3D.
#
# Angle d'un motif, en degrés dans le plan (d, h) :
#     0   = horizontal (lignes de niveau)
#     90  = vertical
#     45  = diagonale montante
#     135 = diagonale descendante
#
# L'angle est métrique : à 45°, un mètre horizontal vaut un mètre vertical,
# ce qui donne bien 45° à l'écran dans Google Earth sans exagération verticale.

MOTIFS: Dict[str, List[float]] = {
    "aucun":           [],
    "plein":           [],          # traité par <extrude>, pas par un motif
    "hachures":        [90.0],      # alias historique
    "hachures_90":     [90.0],
    "hachures_45":     [45.0],
    "hachures_135":    [135.0],
    "hachures_0":      [0.0],
    "quadrillage_0":   [0.0, 90.0],
    "quadrillage_45":  [45.0, 135.0],
}


def mode_rideau(valeur) -> str:
    """Normalise le champ 'rideau'. Accepte les booléens historiques."""
    if valeur is True:
        return "plein"
    if valeur is False or valeur is None:
        return "aucun"
    v = str(valeur).strip().lower()
    if v not in MOTIFS:
        raise ErreurCSV(
            f"rideau = '{valeur}' inconnu — attendu : "
            + ", ".join(f"'{m}'" for m in MOTIFS))
    return v


def _profil(vol: Vol, indices: Sequence[int]) -> Tuple[List[float], List[float]]:
    """Profil (distance latérale, altitude) trié et dédoublonné en d."""
    couples = sorted((vol.distance[i], vol.altitude[i]) for i in indices)
    ds: List[float] = []
    hs: List[float] = []
    for d, h in couples:
        if ds and abs(d - ds[-1]) < 1e-9:
            hs[-1] = max(hs[-1], h)     # même abscisse : on garde le plus haut
        else:
            ds.append(d)
            hs.append(h)
    return ds, hs


def _alt_en(ds: List[float], hs: List[float], d: float) -> float:
    """Altitude du profil à l'abscisse d, par interpolation linéaire."""
    if d < ds[0] or d > ds[-1]:
        return -1.0                      # hors du rideau
    i = bisect.bisect_left(ds, d)
    if i == 0:
        return hs[0]
    d0, d1 = ds[i - 1], ds[i]
    if d1 == d0:
        return max(hs[i - 1], hs[i])
    t = (d - d0) / (d1 - d0)
    return hs[i - 1] + t * (hs[i] - hs[i - 1])


def motif_rideau(vol: Vol, indices: Sequence[int], mode: str,
                 pas_m: float) -> List[List[Tuple[float, float]]]:
    """
    Segments 2D (d, h) du motif demandé, découpés sur la silhouette du rideau.

    Familles de droites parallèles espacées de `pas_m` **perpendiculairement**,
    clippées sur la région {0 <= h <= altitude(d)}. Le découpage se fait par
    échantillonnage puis raffinement par dichotomie aux bords : la silhouette
    d'une trajectoire n'est ni convexe ni monotone, un clipping analytique
    serait fragile pour un gain nul à l'affichage.

    L'espacement est mesuré dans le plan (d, h), donc en distance latérale pour
    les hachures verticales : c'est l'écart que l'œil perçoit sur la carte.
    """
    angles = MOTIFS.get(mode, [])
    if not angles or pas_m <= 0.0 or len(indices) < 2:
        return []

    ds, hs = _profil(vol, indices)
    if len(ds) < 2:
        return []
    dmax, hmax = ds[-1], max(hs)
    if dmax <= 0.0 or hmax <= 0.0:
        return []

    coins = [(0.0, 0.0), (dmax, 0.0), (0.0, hmax), (dmax, hmax)]
    pas_ech = max(0.25, pas_m / 25.0)
    segments: List[List[Tuple[float, float]]] = []

    # Tolérance : au pied exact d'une hachure verticale, h vaut -1e-14 (produit
    # par cos(90°)), et un test strict h >= 0 rejetterait le point, décalant le
    # pied au-dessus du sol.
    EPS = 1e-6

    def dedans(d: float, h: float) -> bool:
        return -EPS <= h <= _alt_en(ds, hs, d) + EPS

    for angle in angles:
        th = math.radians(angle)
        ux, uy = math.cos(th), math.sin(th)          # direction des lignes
        nx, ny = -uy, ux                             # normale : sens d'espacement

        cs = [p[0] * nx + p[1] * ny for p in coins]
        k0 = math.ceil(min(cs) / pas_m)
        k1 = math.floor(max(cs) / pas_m)

        for k in range(k0, k1 + 1):
            c = k * pas_m
            bx, by = c * nx, c * ny
            ts = [(p[0] - bx) * ux + (p[1] - by) * uy for p in coins]
            t, tfin = min(ts), max(ts)

            def pt(t: float) -> Tuple[float, float]:
                return (bx + t * ux, by + t * uy)

            def bord(t_in: float, t_out: float) -> float:
                """Dichotomie sur la frontière entre un t dedans et un t dehors."""
                for _ in range(14):
                    tm = 0.5 * (t_in + t_out)
                    if dedans(*pt(tm)):
                        t_in = tm
                    else:
                        t_out = tm
                return t_in

            debut: Optional[float] = None
            prec = t
            prec_in = dedans(*pt(t))
            if prec_in:
                debut = t
            while prec < tfin:
                cur = min(prec + pas_ech, tfin)
                cur_in = dedans(*pt(cur))
                if cur_in and not prec_in:
                    debut = bord(cur, prec)
                elif prec_in and not cur_in:
                    if debut is not None:
                        segments.append([pt(debut), pt(bord(prec, cur))])
                    debut = None
                prec, prec_in = cur, cur_in
            if prec_in and debut is not None:
                segments.append([pt(debut), pt(tfin)])

    # Aux tangences, une droite peut effleurer la silhouette et produire un
    # segment de longueur nulle : invisible, mais il alourdit le KML et fausse
    # le décompte.
    mini = max(1.0, pas_m / 50.0)
    return [s for s in segments if math.dist(s[0], s[1]) >= mini]


def segments_3d(segments2d: Sequence[Sequence[Tuple[float, float]]],
                phi: float, penetration_m: float = 0.0
                ) -> List[List[Tuple[float, float, float]]]:
    """
    Relève les segments (d, h) dans le repère local ENU le long de l'azimut phi.

    `penetration_m` enfonce sous le sol les extrémités posées à h = 0. Le
    rideau est dessiné en altitude absolue depuis l'altitude déclarée du pas de
    tir ; sur un terrain en relief, un bord pile à h = 0 flotte ou disparaît
    selon la pente. Enfoncer de quelques dizaines de mètres masque le défaut
    sans fausser les altitudes utiles.
    """
    ux, uy = math.cos(phi), math.sin(phi)
    out = []
    for seg in segments2d:
        out.append([(d * ux, d * uy, (h if h > 1e-3 else -penetration_m))
                    for d, h in seg])
    return out


def trace_sol(vol: Vol, indices: Sequence[int],
              phi: float) -> List[Tuple[float, float, float]]:
    """Projection au sol de la trajectoire : borne basse du rideau."""
    ux, uy = math.cos(phi), math.sin(phi)
    return [(vol.distance[i] * ux, vol.distance[i] * uy, 0.0) for i in indices]


def points_3d(vol: Vol, indices: Sequence[int],
              phi: float) -> List[Tuple[float, float, float]]:
    """
    Projette la trajectoire (altitude, distance latérale) dans le repère local
    ENU, le long de l'azimut d'angle mathématique `phi`.

    Les CSV ne fournissent qu'un profil dans le plan de tir : on suppose donc
    une trajectoire rectiligne en plan, ce qui est l'hypothèse usuelle sans
    vent ni données de cap.
    """
    ux, uy = math.cos(phi), math.sin(phi)
    return [(vol.distance[i] * ux, vol.distance[i] * uy, vol.altitude[i])
            for i in indices]
