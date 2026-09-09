#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analyse_terrain.py — Analyseur terrain / domaine de vol.

Rôle : RECHERCHE. Ne dessine rien.

Une analyse est faite PAR GROUPE DE VOLS, depuis le pas de tir et dans la zone
que le groupe désigne. Le domaine obtenu est écrit comme attribut du groupe.

Quand aucun azimut n'est tenable, l'analyseur n'interrompt pas tout : il le
signale et passe au groupe suivant. Si le groupe impose déjà son cône
(`azimut_deg` et `demi_angle_deg`), le domaine est produit tel quel, marqué
`force` — c'est la seule façon de visualiser une configuration qui ne tient
pas, ce qui est justement ce qu'on veut voir.

Entrée  : etude.json
Sortie  : vue.json

Usage
-----
    python3 analyse_terrain.py etude.json -o vue.json
    python3 analyse_terrain.py etude.json --groupe beta --azimut 230
"""

from __future__ import annotations

import argparse
import copy
import math
import os
import sys
from typing import List, Optional, Tuple

import dv_geo as G
import dv_csv as C
import dv_conf as K


# =============================================================================
# Analyse
# =============================================================================

class PasDeSecteur(Exception):
    """Aucun cône n'est déductible : à l'utilisateur d'en imposer un."""


def analyser(ring, portee, P=(0.0, 0.0)):
    """(secteurs, optimal). Lève PasDeSecteur si rien n'est déductible."""
    if not G.point_dans_polygone(P, ring):
        raise PasDeSecteur(
            "le pas de tir est hors de la zone "
            "(vérifie l'ordre lat/lon, ou la zone choisie)")

    secteurs = G.intervalles_admissibles(ring, portee, P)

    if secteurs == G.TOUT_LE_TOUR:
        raise PasDeSecteur(
            f"le disque de rayon {portee:.0f} m tient entièrement dans la "
            "zone : aucun azimut n'est interdit, la zone ne contraint rien")
    if not secteurs:
        d = min(G.distance_sortie(ring, math.radians(a), P) for a in range(360))
        raise PasDeSecteur(
            f"aucun azimut ne permet {portee:.0f} m sans sortir de la zone "
            f"(sortie la plus lointaine : {d:.0f} m environ)")

    return secteurs, max(secteurs, key=G.largeur)


def marges(secteurs, phi):
    """Marges angulaires disponibles de part et d'autre d'un azimut imposé."""
    for iv in secteurs:
        if G.dans_intervalle(phi, iv):
            return (G.norm_angle(phi - iv[0]), G.norm_angle(iv[1] - phi), iv)
    return None


def cone_retenu(secteurs, optimal, az_impose, demi_impose
                ) -> Tuple[float, float, bool, List[str]]:
    """
    Azimut et demi-angle finalement retenus, à partir de l'analyse et des
    valeurs éventuellement imposées. Retourne aussi les avertissements.
    """
    notes: List[str] = []
    if az_impose is None:
        az = G.phi_to_az(G.bissectrice(optimal))
        demi = math.degrees(G.largeur(optimal)) / 2.0
    else:
        az = float(az_impose)
        m = marges(secteurs, G.az_to_phi(az))
        if m is None:
            demi = math.degrees(G.largeur(optimal)) / 2.0
            notes.append(f"azimut imposé {az:.1f}° hors des secteurs "
                         "admissibles : demi-angle repris du secteur optimal")
        else:
            demi = math.degrees(min(m[0], m[1]))

    if demi_impose is not None:
        d0 = float(demi_impose)
        if d0 > demi + 1e-9:
            notes.append(f"demi-angle imposé ± {d0:.1f}° > ± {demi:.1f}° "
                         "exploitable : le secteur débordera de la zone")
        demi = d0
    return az, demi, bool(notes), notes


# =============================================================================
# Rapport
# =============================================================================

def entete(nom, pads, zones, groupes, L):
    add = L.append
    add("=" * 70)
    add(f"  ANALYSE DU DOMAINE DE VOL — {nom}")
    add("=" * 70)
    add("")
    add("  Pas de tir")
    add("  " + "-" * 66)
    for p in pads:
        add(f"      {p.id:<12s} {p.nom:<20s} "
            f"{p.lat:11.7f}, {p.lon:11.7f}   alt {p.alt_m:5.0f} m")
    add("")
    add("  Zones")
    add("  " + "-" * 66)
    for z in zones:
        r = "contrainte" if z.contrainte else "décor"
        pl = f"   plafond {z.plafond_m:.0f} m" if z.plafond_m else ""
        add(f"      {z.id:<12s} {z.nom:<20s} {len(z.points):3d} sommets  "
            f"({r}){pl}")
    add("")
    if groupes:
        add("  Groupes de vols")
        add("  " + "-" * 66)
        for g in groupes:
            vd = g.vol_dimensionnant
            add(f"      {g.id:<12s} {g.nom:<20s} {len(g.vols)} vol(s)   "
                f"portée {g.portee_m:7.0f} m"
                + (f"   ({vd.nom})" if vd else ""))
        add("")


def section_domaine(L, groupe, pad, zone, portee, origine, plafond,
                    ring, secteurs, optimal, azimuts_imposes, P=(0.0, 0.0)):
    add = L.append
    add("")
    add("  " + "=" * 66)
    add(f"  DOMAINE — {groupe.nom}")
    add("  " + "=" * 66)
    add(f"    Pas de tir      : {pad.nom} ({pad.id})")
    add(f"    Zone            : {zone.nom} ({zone.id})")
    add(f"    Portée retenue  : {portee:.0f} m   ({origine})")
    add(f"    Surface de zone : {abs(G.signed_area(ring)) / 1e4:.1f} ha")
    add("")

    if plafond > 0.0:
        add(f"    Plafond de vol : {plafond:.0f} m au-dessus du sol du pas de tir")
        add("    " + "-" * 62)
        depasse = False
        for vol, _ in groupe.vols:
            marge = plafond - vol.apogee_m
            depasse = depasse or marge < 0
            add(f"      {vol.nom:<24s} apogée {vol.apogee_m:7.0f} m   "
                f"marge {marge:+8.0f} m   "
                f"[{'OK ' if marge >= 0 else 'DÉPASSEMENT'}]")
        if depasse:
            add("      [!] au moins un vol dépasse le plafond déclaré.")
        add("")

    add("    Distances de sortie de la zone depuis le pas de tir")
    add("    " + "-" * 62)
    for az in range(0, 360, 30):
        d = G.distance_sortie(ring, G.az_to_phi(az), P)
        add(f"      azimut {az:3d}°   sortie à {d:8.0f} m   "
            f"[{'OK ' if d >= portee else 'NON'}]")
    add("")

    if secteurs is None:
        add("    Aucun secteur admissible pour cette portée.")
        add("")
        return

    add(f"    Secteurs admissibles pour R = {portee:.0f} m")
    add("    " + "-" * 62)
    for iv in sorted(secteurs, key=G.largeur, reverse=True):
        mark = " <== le plus large" if iv is optimal else ""
        add(f"      {G.phi_to_az(iv[1]):6.1f}° -> {G.phi_to_az(iv[0]):6.1f}°"
            f"   ouverture {math.degrees(G.largeur(iv)):6.1f}°"
            f"   bissectrice {G.phi_to_az(G.bissectrice(iv)):6.1f}°{mark}")
    add("")

    w = math.degrees(G.largeur(optimal))
    bis = G.bissectrice(optimal)
    sortie = G.distance_sortie(ring, bis, P)
    add("    RECOMMANDATION")
    add("    " + "-" * 62)
    add(f"      Azimut de tir   : {G.phi_to_az(bis):.1f}°")
    add(f"      Demi-angle      : ± {w / 2:.1f}°")
    add(f"      Sortie sur axe  : {sortie:.0f} m  (marge {sortie - portee:.0f} m)")
    add("")

    if azimuts_imposes:
        add("    AZIMUTS IMPOSÉS")
        add("    " + "-" * 62)
        for az in azimuts_imposes:
            phi = G.az_to_phi(az)
            m = marges(secteurs, phi)
            if m is None:
                add(f"      {az:6.1f}°  HORS DOMAINE — sortie à "
                    f"{G.distance_sortie(ring, phi, P):.0f} m "
                    f"pour R = {portee:.0f} m")
            else:
                g_, d_, iv = m
                add(f"      {az:6.1f}°  admissible — marge "
                    f"{math.degrees(g_):.1f}° vers {G.phi_to_az(iv[0]):.1f}°, "
                    f"{math.degrees(d_):.1f}° vers {G.phi_to_az(iv[1]):.1f}°")
                add(f"                demi-angle symétrique exploitable : "
                    f"± {math.degrees(min(g_, d_)):.1f}°")
        add("")


# =============================================================================
# Programme principal
# =============================================================================

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Analyse le domaine de vol, un par groupe de vols.")
    ap.add_argument("etude", help="JSON d'étude (pas de tir, zones, vols)")
    ap.add_argument("--groupe", default=None, metavar="ID",
                    help="n'analyser que ce groupe (défaut : tous)")
    ap.add_argument("--pas-de-tir", default=None, metavar="ID",
                    help="force le pas de tir de tous les groupes")
    ap.add_argument("--zone", default=None, metavar="ID",
                    help="force la zone de contrainte de tous les groupes")
    ap.add_argument("--portee", type=float, default=None,
                    help="portée horizontale max en m (surcharge tout)")
    ap.add_argument("--plafond", type=float, default=None,
                    help="plafond de vol en m au-dessus du sol du pas de tir")
    ap.add_argument("--azimut", type=float, action="append", default=None,
                    metavar="DEG", help="azimut imposé à évaluer (répétable)")
    ap.add_argument("-o", "--vue", default=None,
                    help="écrit la vue JSON pour genere_kml.py")
    ap.add_argument("--rapport", default=None,
                    help="écrit aussi le rapport dans ce fichier texte")
    args = ap.parse_args(argv)

    try:
        cfg = K.lire_json(args.etude)
        pads = K.charger_pas_de_tir(cfg)
        zones = K.charger_zones(cfg)
    except K.ErreurConf as e:
        raise SystemExit(f"Configuration : {e}")

    base = os.path.dirname(os.path.abspath(args.etude))
    try:
        groupes = C.charger_vols(cfg.get("vols"), base, K.normaliser_rideau)
    except (C.ErreurCSV, K.ErreurConf) as e:
        raise SystemExit(f"Lecture des vols : {e}")

    if args.groupe:
        groupes = [C.choisir_groupe(groupes, args.groupe)]
    if not groupes:
        raise SystemExit(
            "Aucun groupe de vols : rien à analyser.\n"
            "  -> un domaine est toujours rattaché à un groupe ; déclare au "
            "moins un groupe dans 'vols.groupes'."
        )

    nom = cfg.get("nom", "sans nom")
    L: List[str] = []
    entete(nom, pads, zones, groupes, L)

    alertes: List[str] = []
    resolus: dict = {}

    for groupe in groupes:
        try:
            pad = K.choisir_pas_de_tir(pads, args.pas_de_tir or groupe.pas_de_tir)
            zone = K.choisir_zone(zones, args.zone or groupe.zone)
        except K.ErreurConf as e:
            raise SystemExit(f"Configuration : groupe « {groupe.id} » : {e}")

        portee, origine = args.portee, "option --portee"
        if portee is None:
            portee = groupe.portee_m
            vd = groupe.vol_dimensionnant
            origine = f"vol « {vd.nom} »" if vd else "groupe"
        if not portee:
            alertes.append(f"groupe « {groupe.id} » : portée nulle, ignoré")
            continue
        portee = float(portee)

        plafond = (args.plafond if args.plafond is not None
                   else groupe.plafond_m)

        enu = G.ENU(pad.lat, pad.lon)
        ring = G.ensure_ccw(G.open_ring([enu.fwd(a, b) for a, b in zone.points]))

        secteurs = optimal = None
        echec = None
        try:
            secteurs, optimal = analyser(ring, portee)
        except PasDeSecteur as e:
            echec = str(e)

        section_domaine(L, groupe, pad, zone, portee, origine, plafond,
                        ring, secteurs, optimal, args.azimut or [])

        az_i, demi_i = groupe.azimut_deg, groupe.demi_angle_deg

        if echec is not None:
            # Analyse impossible : on ne fabrique un domaine que si le groupe
            # impose entièrement son cône.
            if az_i is None:
                alertes.append(
                    f"groupe « {groupe.id} » : {echec}.\n"
                    f"      Aucun domaine produit. Pour en forcer un, fixe "
                    f"'azimut_deg' et 'demi_angle_deg' sur le groupe.")
                L.append(f"    [!] {echec} — aucun domaine produit.")
                L.append("")
                continue
            if demi_i is None:
                alertes.append(
                    f"groupe « {groupe.id} » : {echec}.\n"
                    f"      'azimut_deg' est fixé à {az_i:.1f}° mais pas "
                    f"'demi_angle_deg' : impossible de déduire l'ouverture.\n"
                    f"      Aucun domaine produit.")
                L.append(f"    [!] {echec} — 'demi_angle_deg' manquant, "
                         "aucun domaine produit.")
                L.append("")
                continue
            az, demi, force = float(az_i), float(demi_i), True
            alertes.append(
                f"groupe « {groupe.id} » : {echec}.\n"
                f"      Domaine forcé à {az:.1f}° ± {demi:.1f}° "
                f"— il sortira de la zone.")
            L.append(f"    [!] {echec}")
            L.append(f"    Domaine FORCÉ : {az:.1f}° ± {demi:.1f}°")
            L.append("")
        else:
            az, demi, force, notes = cone_retenu(secteurs, optimal, az_i, demi_i)
            for n in notes:
                alertes.append(f"groupe « {groupe.id} » : {n}")
            if az_i is not None or demi_i is not None:
                L.append(f"    CÔNE RETENU (imposé) : {az:.1f}° ± {demi:.1f}°")
                L.append("")

        resolus[groupe.id] = K.Domaine(
            id=groupe.id, nom=groupe.nom, azimut_deg=az, demi_angle_deg=demi,
            portee_max_m=portee, pas_de_tir=pad.id, zone=zone.id,
            groupe=groupe.id, plafond_m=plafond, couleur=groupe.couleur,
            force=force)

    L.append("=" * 70)
    txt = "\n".join(L)
    print(txt)
    if args.rapport:
        with open(args.rapport, "w", encoding="utf-8") as f:
            f.write(txt + "\n")

    for a in alertes:
        print(f"[!] {a}", file=sys.stderr)

    if args.vue:
        # La vue est faite pour être relue et éditée : tous les réglages y sont
        # écrits en clair, y compris ceux qui n'étaient qu'implicites dans
        # l'étude. Sans cela, on ne peut pas basculer un affichage sans aller
        # chercher le nom de la clé dans la documentation.
        rendu = K.charger_rendu(cfg)
        vue = {
            "$schema": "./vue.schema.json",
            "_genere_par": "analyse_terrain.py — éditable avant génération du KML",
            "nom": nom,
            "pas_de_tir": [],
            "zones": [],
            "rendu": rendu,
        }
        for p in pads:
            d = K.pad_dict(p)
            d["affichage"] = K.normaliser_affichage(
                p.affichage, "pas_de_tir", rendu["affichage"]["pas_de_tir"])
            vue["pas_de_tir"].append(d)
        for z in zones:
            d = K.zone_dict(z)
            d["volume"] = K.volume_de(rendu, z.volume)
            d["affichage"] = K.normaliser_affichage(
                z.affichage, "zone", rendu["affichage"]["zone"])
            vue["zones"].append(d)

        if cfg.get("vols"):
            bloc = copy.deepcopy(cfg["vols"])
            gs = bloc.get("groupes")
            if gs is None:
                gs = [{"id": "vols", "nom": "Vols", "liste": bloc.pop("liste")}]
                bloc["groupes"] = gs
            par_id = {g.id: g for g in groupes}
            for gb in gs:
                gid = gb.get("id")
                if gid in resolus:
                    dom = resolus[gid]
                    dd = K.dom_dict(dom)
                    dd["volume"] = K.volume_de(rendu, dom.volume)
                    dd["affichage"] = K.normaliser_affichage(
                        dom.affichage, "domaine", rendu["affichage"]["domaine"])
                    gb["domaine"] = dd
                else:
                    gb.pop("domaine", None)
                gr = par_id.get(gid)
                if gr is None:
                    continue
                for entree, (_, opts) in zip(gb.get("liste", []), gr.vols):
                    entree["rideau"] = opts["rideau"]
                    entree["affichage"] = K.normaliser_affichage(
                        opts.get("affichage"), "vol", rendu["affichage"]["vol"])
            vue["vols"] = bloc
        K.ecrire_json(args.vue, vue)
        print(f"Vue écrite : {args.vue}   ({len(resolus)} domaine(s) sur "
              f"{len(groupes)} groupe(s))")
        print(f"  -> édite-la si besoin, puis : python3 genere_kml.py {args.vue}")

    return 0 if resolus else 1


if __name__ == "__main__":
    # Les erreurs de configuration remontent de partout (affichage, rideau,
    # références) : un seul filet, pour ne jamais afficher de traceback.
    try:
        sys.exit(main())
    except (K.ErreurConf, C.ErreurCSV) as err:
        raise SystemExit(f"Configuration : {err}")
