#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dv_conf.py — Chargement et normalisation de la configuration (stdlib seule).

Deux fichiers, deux rôles :

  * `etude.json` — ce qui est étudié : pas de tir, zones au sol, vols.
    Écrit à la main, stable.
  * `vue.json`   — ce qui est dessiné : la même chose plus un azimut, un
    demi-angle et une portée résolus. Produit par l'analyseur, éditable.

Ce module ne fait aucune géométrie : il lit, valide, applique les valeurs par
défaut et résout les références par identifiant.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


class ErreurConf(Exception):
    pass


# =============================================================================
# Pas de tir et zones
# =============================================================================

@dataclass
class PasDeTir:
    id: str
    nom: str
    lat: float
    lon: float
    alt_m: float = 0.0
    reference: bool = False
    affichage: Dict[str, Any] = field(default_factory=dict)


@dataclass
class Zone:
    id: str
    nom: str
    points: List[List[float]]          # [[lat, lon], ...]
    contrainte: bool = False           # borne les domaines de vol
    couleur: Dict[str, Any] = field(default_factory=dict)
    plafond_m: float = 0.0             # une zone peut avoir son propre plafond
    volume: Dict[str, Any] = field(default_factory=dict)
    affichage: Dict[str, Any] = field(default_factory=dict)


def _liste(valeur, champ: str) -> List[dict]:
    """Accepte un objet seul là où une liste est attendue."""
    if valeur is None:
        raise ErreurConf(f"'{champ}' manquant")
    if isinstance(valeur, dict):
        return [valeur]
    if isinstance(valeur, list):
        return valeur
    raise ErreurConf(f"'{champ}' doit être un objet ou une liste")


def charger_pas_de_tir(cfg: dict) -> List[PasDeTir]:
    brut = _liste(cfg.get("pas_de_tir"), "pas_de_tir")
    pads: List[PasDeTir] = []
    for i, p in enumerate(brut):
        for c in ("lat", "lon"):
            if c not in p:
                raise ErreurConf(f"pas_de_tir[{i}] : '{c}' manquant")
        pid = str(p.get("id", f"pdt_{i + 1}"))
        pads.append(PasDeTir(
            id=pid, nom=str(p.get("nom", pid)),
            lat=float(p["lat"]), lon=float(p["lon"]),
            alt_m=float(p.get("alt_m", 0.0)),
            reference=bool(p.get("reference", False)),
            affichage=p.get("affichage")))
    ids = [p.id for p in pads]
    if len(set(ids)) != len(ids):
        raise ErreurConf(f"identifiants de pas de tir en double : {ids}")
    return pads


def charger_zones(cfg: dict) -> List[Zone]:
    # 'terrain' : forme historique à une seule zone
    if "zones" not in cfg and "terrain" in cfg:
        brut = [{"id": "terrain", "nom": "Terrain",
                 "points": cfg["terrain"], "contrainte": True}]
    else:
        brut = _liste(cfg.get("zones"), "zones")

    zones: List[Zone] = []
    for i, z in enumerate(brut):
        pts = z.get("points")
        if not pts or len(pts) < 3:
            raise ErreurConf(f"zones[{i}] : 'points' doit contenir >= 3 sommets")
        zid = str(z.get("id", f"zone_{i + 1}"))
        zones.append(Zone(
            id=zid, nom=str(z.get("nom", zid)), points=pts,
            contrainte=bool(z.get("contrainte", False)),
            couleur=dict(z.get("couleur", {})),
            plafond_m=float(z.get("plafond_m") or 0.0),
            volume=dict(z.get("volume", {})),
            affichage=z.get("affichage")))
    ids = [z.id for z in zones]
    if len(set(ids)) != len(ids):
        raise ErreurConf(f"identifiants de zone en double : {ids}")
    return zones


def _choisir(objets, ident, drapeau: str, quoi: str):
    """Résout une référence : identifiant explicite, puis drapeau, puis unique."""
    if ident is not None:
        for o in objets:
            if o.id == ident:
                return o
        raise ErreurConf(
            f"{quoi} '{ident}' introuvable — disponibles : "
            + ", ".join(f"'{o.id}'" for o in objets))
    marques = [o for o in objets if getattr(o, drapeau)]
    if len(marques) == 1:
        return marques[0]
    if len(marques) > 1:
        raise ErreurConf(
            f"plusieurs {quoi} marqués '{drapeau}' : "
            + ", ".join(f"'{o.id}'" for o in marques)
            + " — un seul est permis, ou désigne-le explicitement")
    if len(objets) == 1:
        return objets[0]
    raise ErreurConf(
        f"{len(objets)} {quoi} déclarés et aucun marqué '{drapeau}' — "
        "précise lequel utiliser")


def choisir_pas_de_tir(pads: List[PasDeTir], ident=None) -> PasDeTir:
    return _choisir(pads, ident, "reference", "pas de tir")


def choisir_zone(zones: List[Zone], ident=None) -> Zone:
    return _choisir(zones, ident, "contrainte", "zone")


@dataclass
class Domaine:
    """Un domaine de vol résolu : de quel pas de tir, dans quelle zone, pour
    quel groupe de fusées, avec quel cône."""
    id: str
    nom: str
    azimut_deg: float
    demi_angle_deg: float
    portee_max_m: float
    pas_de_tir: Optional[str] = None
    zone: Optional[str] = None
    groupe: Optional[str] = None
    plafond_m: float = 0.0
    couleur: Optional[str] = None
    volume: Dict[str, Any] = field(default_factory=dict)
    affichage: Dict[str, Any] = field(default_factory=dict)
    force: bool = False                # cône imposé, non validé par l'analyse


def domaine_depuis(d: Optional[dict], groupe_id: str, groupe_nom: str,
                   couleur=None) -> Optional[Domaine]:
    """
    Construit un domaine depuis l'attribut `domaine` d'un groupe de vols.

    Un domaine est toujours rattaché à un groupe : il en tire sa portée, son
    pas de tir et sa zone. Le représenter ailleurs qu'à l'intérieur du groupe
    obligerait à maintenir une correspondance par identifiant pour rien.
    """
    if not d:
        return None
    for c in ("azimut_deg", "demi_angle_deg", "portee_max_m"):
        if d.get(c) is None:
            raise ErreurConf(f"groupe '{groupe_id}', domaine : '{c}' manquant")
    return Domaine(
        id=groupe_id, nom=str(d.get("nom", groupe_nom)),
        azimut_deg=float(d["azimut_deg"]),
        demi_angle_deg=float(d["demi_angle_deg"]),
        portee_max_m=float(d["portee_max_m"]),
        pas_de_tir=d.get("pas_de_tir"), zone=d.get("zone"),
        groupe=groupe_id,
        plafond_m=float(d.get("plafond_m") or 0.0),
        couleur=d.get("couleur", couleur),
        volume=dict(d.get("volume", {})),
        affichage=d.get("affichage"),
        force=bool(d.get("force", False)))


# =============================================================================
# Rideau
# =============================================================================

RIDEAU_DEFAUT: Dict[str, Any] = {
    "motif": "aucun",
    "pas_m": 100.0,
    "epaisseur": 1.5,
    "alpha": 0.25,
    "penetration_sol_m": 30.0,
    "trace_sol": False,
}

_ALIAS_RIDEAU = {True: "plein", False: "aucun"}


def normaliser_rideau(valeur, herite: Optional[dict] = None) -> Dict[str, Any]:
    """
    Tous les réglages du rideau tiennent dans un objet. Une chaîne ou un
    booléen restent acceptés comme raccourci pour `{"motif": ...}`, ce qui
    évite d'imposer un objet complet quand seul le motif change.
    """
    base = dict(herite or RIDEAU_DEFAUT)
    if valeur is None:
        return base
    if isinstance(valeur, bool):
        base["motif"] = _ALIAS_RIDEAU[valeur]
        return base
    if isinstance(valeur, str):
        base["motif"] = valeur
        return base
    if isinstance(valeur, dict):
        inconnus = set(valeur) - set(RIDEAU_DEFAUT)
        if inconnus:
            raise ErreurConf(
                "rideau : clé(s) inconnue(s) " + ", ".join(sorted(inconnus))
                + " — attendu : " + ", ".join(sorted(RIDEAU_DEFAUT)))
        base.update(valeur)
        return base
    raise ErreurConf("rideau doit être une chaîne, un booléen ou un objet")


# =============================================================================
# Affichage
# =============================================================================
#
# Chaque élément graphique peut être allumé ou éteint indépendamment. Les
# valeurs par défaut sont dans `rendu.affichage`, et chaque élément (zone,
# domaine, vol, pas de tir) peut les redéfinir.

AFFICHAGE_DEFAUT: Dict[str, Dict[str, bool]] = {
    "zone": {
        "polygone": True,       # empreinte au sol
        "volume": True,         # murs ou arêtes jusqu'au plafond
        "plafond": True,        # contour à l'altitude du plafond
    },
    "domaine": {
        "secteur": True,        # polygone rempli, découpé par la zone
        "limites": True,        # segments A et B
        "reperes": True,        # repères ponctuels A et B
        "contour_secteur": True,   # secteur demandé, non découpé
        "cercle_portee": True,
        "volume": True,
        "plafond": True,
    },
    "vol": {
        "trajectoire": True,
        "rideau": True,         # motif du rideau
        "trace_sol": True,      # soumis aussi à rideau.trace_sol
    },
    "pas_de_tir": {
        "repere": True,
    },
}


def normaliser_affichage(valeur, genre: str,
                         herite: Optional[dict] = None) -> Dict[str, bool]:
    """
    Normalise un bloc `affichage`. Un booléen seul allume ou éteint tout
    l'élément — c'est le geste le plus fréquent, il ne doit pas obliger à
    énumérer les clés.
    """
    base = dict(herite or AFFICHAGE_DEFAUT[genre])
    if valeur is None:
        return base
    if isinstance(valeur, bool):
        return {k: valeur for k in base}
    if isinstance(valeur, dict):
        inconnus = set(valeur) - set(AFFICHAGE_DEFAUT[genre])
        if inconnus:
            raise ErreurConf(
                f"affichage ({genre}) : clé(s) inconnue(s) "
                + ", ".join(sorted(inconnus)) + " — attendu : "
                + ", ".join(sorted(AFFICHAGE_DEFAUT[genre])))
        base.update({k: bool(v) for k, v in valeur.items()})
        return base
    raise ErreurConf(f"affichage ({genre}) doit être un booléen ou un objet")


# =============================================================================
# Rendu
# =============================================================================

RENDU_DEFAUT: Dict[str, Any] = {
    "volume": {
        # Extruder une empreinte de plusieurs centaines d'hectares sur des
        # kilomètres de haut produit un pavé opaque qui masque ce qu'il contient :
        # le fil de fer est le défaut. Ce qui s'affiche ou non se règle par
        # `affichage.volume` et `affichage.plafond`, élément par élément.
        "style": "aretes",      # "aretes" | "plein"
        "alpha_mur": 0.15,
    },
    "affichage": AFFICHAGE_DEFAUT,
    "segments_pointilles": True,
    "tiret_m": 25.0,
    "espace_m": 15.0,
    "couleurs": {
        "terrain": {"trait": "#ffd200", "fond": "#ffd200",
                    "alpha_fond": 0.12, "epaisseur": 2},
        "domaine": {"trait": "#ff3020", "fond": "#ff3020",
                    "alpha_fond": 0.28, "epaisseur": 2},
        "cercle":  {"trait": "#22a0ff", "epaisseur": 2},
        "secteur": {"trait": "#ff8800", "epaisseur": 2},
        "limites": {"trait": "#ffffff", "epaisseur": 3},
        "plafond": {"trait": "#ff3020", "epaisseur": 2},
    },
}


def fusion(base: dict, sur: Optional[dict]) -> dict:
    """Fusion récursive : `sur` écrase `base`, dictionnaire par dictionnaire."""
    out = dict(base)
    for k, v in (sur or {}).items():
        out[k] = fusion(out[k], v) if isinstance(v, dict) and \
            isinstance(out.get(k), dict) else v
    return out


def charger_rendu(cfg: dict) -> dict:
    r = fusion(RENDU_DEFAUT, cfg.get("rendu"))
    r["affichage"] = fusion(AFFICHAGE_DEFAUT, (cfg.get("rendu") or {}).get("affichage"))
    return r


def volume_de(rendu: dict, surcharge: Optional[dict]) -> dict:
    """Réglages de volume d'un élément : le global, redéfini par l'élément."""
    return fusion(rendu["volume"], surcharge or {})


def dom_dict(d: Domaine) -> dict:
    out = {"nom": d.nom}
    for c in ("pas_de_tir", "zone", "couleur"):
        if getattr(d, c) is not None:
            out[c] = getattr(d, c)
    out.update({"azimut_deg": round(d.azimut_deg, 3),
                "demi_angle_deg": round(d.demi_angle_deg, 3),
                "portee_max_m": round(d.portee_max_m, 1)})
    if d.plafond_m:
        out["plafond_m"] = round(d.plafond_m, 1)
    if d.force:
        out["force"] = True
    if d.volume:
        out["volume"] = d.volume
    if d.affichage:
        out["affichage"] = d.affichage
    return out


def lire_json(chemin: str) -> dict:
    if not os.path.isfile(chemin):
        raise ErreurConf(f"fichier introuvable : {chemin}")
    with open(chemin, "r", encoding="utf-8-sig") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError as e:
            raise ErreurConf(f"{chemin} : JSON invalide ligne {e.lineno} — {e.msg}")


def ecrire_json(chemin: str, donnees: dict) -> None:
    with open(chemin, "w", encoding="utf-8") as f:
        json.dump(donnees, f, indent=2, ensure_ascii=False)
        f.write("\n")


def pad_dict(p: PasDeTir) -> dict:
    d = {"id": p.id, "nom": p.nom, "lat": p.lat, "lon": p.lon,
         "alt_m": p.alt_m, "reference": p.reference}
    if p.affichage is not None:
        d["affichage"] = p.affichage
    return d


def zone_dict(z: Zone) -> dict:
    d = {"id": z.id, "nom": z.nom, "points": z.points,
         "contrainte": z.contrainte}
    if z.couleur:
        d["couleur"] = z.couleur
    if z.plafond_m:
        d["plafond_m"] = z.plafond_m
    if z.volume:
        d["volume"] = z.volume
    if z.affichage is not None:
        d["affichage"] = z.affichage
    return d
