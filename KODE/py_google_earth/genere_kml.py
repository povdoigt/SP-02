#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
genere_kml.py — Générateur KML à partir d'une vue.

Rôle : CONSTRUCTION. Ne cherche rien, ne recommande rien.

La vue impose un azimut et un demi-angle ; le générateur dessine le secteur
correspondant, découpé par la zone de contrainte. Si le secteur demandé sort
de la zone, il le DESSINE QUAND MÊME et le signale : c'est volontaire, on veut
pouvoir visualiser une configuration non conforme.

Usage
-----
    python3 genere_kml.py vue.json
    python3 genere_kml.py vue.json --azimut 230 -o essai_230.kml
"""

from __future__ import annotations

import argparse
import math
import os
import sys
from typing import List, Sequence

import dv_geo as G
import dv_csv as C
import dv_conf as K

Pt = G.Pt


# =============================================================================
# Construction KML
# =============================================================================

def kml_color(rgb: str, alpha: float) -> str:
    """'#RRGGBB' + alpha [0..1] -> 'aabbggrr' (ordre KML, inversé)."""
    s = rgb.lstrip("#")
    a = format(max(0, min(255, int(round(alpha * 255)))), "02x")
    return f"{a}{s[4:6]}{s[2:4]}{s[0:2]}".lower()


def esc(s: str) -> str:
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


class KmlBuilder:
    def __init__(self, enu: G.ENU, alt_sol: float):
        self.enu = enu
        self.alt_sol = alt_sol
        self.styles: List[str] = []
        self.body: List[str] = []
        self._vus = set()

    # -- styles -------------------------------------------------------------
    def style_ligne(self, sid, rgb, width, alpha=1.0):
        if sid in self._vus:
            return
        self._vus.add(sid)
        self.styles.append(
            f'<Style id="{sid}"><LineStyle>'
            f'<color>{kml_color(rgb, alpha)}</color><width>{width}</width>'
            f'</LineStyle><LabelStyle><scale>0</scale></LabelStyle></Style>')

    def style_poly(self, sid, rgb_line, rgb_fill, width, alpha_fill,
                   alpha_line=1.0):
        if sid in self._vus:
            return
        self._vus.add(sid)
        self.styles.append(
            f'<Style id="{sid}">'
            f'<LineStyle><color>{kml_color(rgb_line, alpha_line)}</color>'
            f'<width>{width}</width></LineStyle>'
            f'<PolyStyle><color>{kml_color(rgb_fill, alpha_fill)}</color>'
            f'<fill>1</fill><outline>1</outline></PolyStyle>'
            f'<LabelStyle><scale>0</scale></LabelStyle></Style>')

    def style_traj(self, sid, rgb, width, alpha_mur):
        """PolyStyle sert au mur d'un LineString extrudé : sans lui, Google
        Earth remplit le rideau plein avec une couleur par défaut."""
        if sid in self._vus:
            return
        self._vus.add(sid)
        self.styles.append(
            f'<Style id="{sid}">'
            f'<LineStyle><color>{kml_color(rgb, 1.0)}</color>'
            f'<width>{width}</width></LineStyle>'
            f'<PolyStyle><color>{kml_color(rgb, alpha_mur)}</color>'
            f'<fill>1</fill><outline>0</outline></PolyStyle>'
            f'<LabelStyle><scale>0</scale></LabelStyle></Style>')

    # -- coordonnées --------------------------------------------------------
    def coords(self, pts: Sequence[Pt], alt_rel: float) -> str:
        out = []
        for x, y in pts:
            lat, lon = self.enu.inv(x, y)
            out.append(f"{lon:.8f},{lat:.8f},{self.alt_sol + alt_rel:.2f}")
        return " ".join(out)

    def coords3(self, pts3) -> str:
        """pts3 = [(x, y, alt_relative_sol)] : altitude propre à chaque point."""
        out = []
        for x, y, h in pts3:
            lat, lon = self.enu.inv(x, y)
            out.append(f"{lon:.8f},{lat:.8f},{self.alt_sol + h:.2f}")
        return " ".join(out)

    # -- éléments -----------------------------------------------------------
    def polygone(self, nom, sid, pts, hauteur_m=0.0, description=""):
        ring = list(pts)
        if not G._close(ring[0], ring[-1]):
            ring.append(ring[0])          # KML exige un LinearRing fermé
        mode = "absolute" if hauteur_m > 0.0 else "clampToGround"
        self.body.append(
            f'<Placemark><name>{esc(nom)}</name>'
            f'<description>{esc(description)}</description>'
            f'<styleUrl>#{sid}</styleUrl>'
            f'<Polygon><extrude>{1 if hauteur_m > 0 else 0}</extrude>'
            f'<altitudeMode>{mode}</altitudeMode>'
            f'<outerBoundaryIs><LinearRing><coordinates>'
            f'{self.coords(ring, hauteur_m)}'
            f'</coordinates></LinearRing></outerBoundaryIs></Polygon></Placemark>')

    def ligne(self, nom, sid, pts, hauteur_m=0.0, description=""):
        self.lignes_multi(nom, sid, [list(pts)], hauteur_m, description)

    def lignes_multi(self, nom, sid, groupes, hauteur_m=0.0, description=""):
        """Plusieurs LineString dans UN Placemark : évite de polluer le panneau
        latéral de Google Earth avec des dizaines d'entrées."""
        mode = "absolute" if hauteur_m > 0.0 else "clampToGround"
        geoms = "".join(
            f'<LineString><tessellate>1</tessellate>'
            f'<altitudeMode>{mode}</altitudeMode>'
            f'<coordinates>{self.coords(g, hauteur_m)}</coordinates></LineString>'
            for g in groupes)
        self.body.append(
            f'<Placemark><name>{esc(nom)}</name>'
            f'<description>{esc(description)}</description>'
            f'<styleUrl>#{sid}</styleUrl>'
            f'<MultiGeometry>{geoms}</MultiGeometry></Placemark>')

    def trajectoire(self, nom, sid, pts3, rideau=False, description=""):
        self.body.append(
            f'<Placemark><name>{esc(nom)}</name>'
            f'<description>{esc(description)}</description>'
            f'<styleUrl>#{sid}</styleUrl>'
            f'<LineString><extrude>{1 if rideau else 0}</extrude>'
            f'<altitudeMode>absolute</altitudeMode>'
            f'<coordinates>{self.coords3(pts3)}</coordinates>'
            f'</LineString></Placemark>')

    def trajectoires_multi(self, nom, sid, groupes3, description=""):
        """Plusieurs polylignes 3D dans UN Placemark (motif d'un rideau)."""
        geoms = "".join(
            f'<LineString><altitudeMode>absolute</altitudeMode>'
            f'<coordinates>{self.coords3(g)}</coordinates></LineString>'
            for g in groupes3)
        self.body.append(
            f'<Placemark><name>{esc(nom)}</name>'
            f'<description>{esc(description)}</description>'
            f'<styleUrl>#{sid}</styleUrl>'
            f'<MultiGeometry>{geoms}</MultiGeometry></Placemark>')

    def contour_altitude(self, nom, sid, pts, hauteur_m, description=""):
        """Contour fermé tracé à une altitude donnée : matérialise le plafond."""
        ring = list(pts)
        if not G._close(ring[0], ring[-1]):
            ring.append(ring[0])
        self.body.append(
            f'<Placemark><name>{esc(nom)}</name>'
            f'<description>{esc(description)}</description>'
            f'<styleUrl>#{sid}</styleUrl>'
            f'<LineString><altitudeMode>absolute</altitudeMode>'
            f'<coordinates>{self.coords(ring, hauteur_m)}</coordinates>'
            f'</LineString></Placemark>')

    def aretes_verticales(self, nom, sid, pts, hauteur_m, description=""):
        """Une arête verticale par sommet, du sol au plafond.

        Alternative aux murs pleins : un pavé extrudé de plusieurs kilomètres
        de haut masque les trajectoires qu'il est censé contenir, alors qu'un
        volume en fil de fer se lit sans rien cacher."""
        groupes = [[(x, y, 0.0), (x, y, hauteur_m)] for x, y in pts]
        self.trajectoires_multi(nom, sid, groupes, description)

    def ligne_sol(self, nom, sid, pts3, description=""):
        """Polyligne plaquée sur le relief : clampToGround ignore l'altitude et
        suit le terrain, ce qui évite qu'elle s'enterre si l'altitude déclarée
        du pas de tir diffère du relief réel."""
        pts = [(x, y) for x, y, _ in pts3]
        self.body.append(
            f'<Placemark><name>{esc(nom)}</name>'
            f'<description>{esc(description)}</description>'
            f'<styleUrl>#{sid}</styleUrl>'
            f'<LineString><tessellate>1</tessellate>'
            f'<altitudeMode>clampToGround</altitudeMode>'
            f'<coordinates>{self.coords(pts, 0.0)}</coordinates>'
            f'</LineString></Placemark>')

    def point(self, nom, p, description=""):
        self.body.append(
            f'<Placemark><name>{esc(nom)}</name>'
            f'<description>{esc(description)}</description>'
            f'<Point><altitudeMode>clampToGround</altitudeMode>'
            f'<coordinates>{self.coords([p], 0.0)}</coordinates>'
            f'</Point></Placemark>')

    def ouvrir_dossier(self, nom, ouvert=True):
        self.body.append(f'<Folder><name>{esc(nom)}</name>'
                         f'<open>{1 if ouvert else 0}</open>')

    def fermer_dossier(self):
        self.body.append("</Folder>")

    def render(self, nom_doc: str) -> str:
        return ('<?xml version="1.0" encoding="UTF-8"?>\n'
                '<kml xmlns="http://www.opengis.net/kml/2.2" '
                'xmlns:gx="http://www.google.com/kml/ext/2.2">\n'
                f'<Document><name>{esc(nom_doc)}</name><open>1</open>\n'
                + "\n".join(self.styles) + "\n"
                + "\n".join(self.body) + "\n</Document></kml>\n")


def dessiner_volume(k, nom, sid_mur, sid_ligne, pts, plafond, style, desc="",
                    murs=True, plafond_visible=True):
    """
    Volume au-dessus d'une empreinte au sol, jusqu'au plafond.

      * "plein"  : polygone extrudé — murs et couvercle d'une seule couleur,
        KML ne permettant pas de les distinguer dans un même Placemark.
      * "aretes" : une verticale par sommet, plus le contour du plafond.

    Dans les deux cas le contour du plafond est tracé : c'est la ligne que
    l'on vient lire, et un mur translucide seul ne la donne pas nettement.
    """
    if murs:
        if style == "plein":
            k.polygone(f"{nom} — murs", sid_mur, pts, plafond, desc)
        else:
            k.aretes_verticales(f"{nom} — arêtes", sid_ligne, pts, plafond, desc)
    if plafond_visible:
        k.contour_altitude(f"{nom} — plafond", sid_ligne, pts, plafond,
                           f"Plafond à {plafond:.0f} m au-dessus du sol")


class Surcharge(argparse.Action):
    """
    Conserve l'ORDRE des options sur la ligne de commande.

    argparse écrase par défaut : `--domaine a --azimut 1 --domaine b --azimut 2`
    ne retiendrait que `b` et `2`, en perdant silencieusement la surcharge sur
    `a`. On enregistre donc la séquence telle qu'elle est tapée, et chaque
    `--domaine` ouvre une portée pour les options qui le suivent.
    """

    def __call__(self, parser, ns, valeur, option=None):
        seq = getattr(ns, "sequence", None)
        if seq is None:
            seq = []
            setattr(ns, "sequence", seq)
        seq.append((self.dest, valeur))


def resoudre_surcharges(sequence, doms):
    """
    Répartit les surcharges par domaine.

    Une option placée AVANT tout `--domaine` s'applique à tous les domaines ;
    placée après, elle ne vaut que pour le domaine courant.
    """
    ids = [d.id for d in doms]
    out = {i: {} for i in ids}
    cible = None
    for cle, valeur in sequence or []:
        if cle == "domaine":
            if valeur not in ids:
                raise K.ErreurConf(
                    f"domaine '{valeur}' introuvable — "
                    + ("cette vue n'en contient aucun"
                       if not ids else
                       "disponibles : " + ", ".join(f"'{i}'" for i in ids)))
            cible = valeur
        else:
            for i in ([cible] if cible else ids):
                out[i][cle] = valeur
    return out


# =============================================================================
# Programme principal
# =============================================================================

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Génère le KML d'une vue de domaine de vol.")
    ap.add_argument("vue", help="vue JSON (issue de analyse_terrain.py)")
    ap.add_argument("-o", "--out", default=None, help="fichier KML de sortie")
    ap.add_argument("--domaine", action=Surcharge, metavar="ID",
                    help="ouvre une portée : les surcharges qui suivent ne "
                         "valent que pour ce domaine. Répétable")
    ap.add_argument("--azimut", action=Surcharge, type=float, metavar="DEG",
                    help="surcharge l'azimut (degrés)")
    ap.add_argument("--demi-angle", action=Surcharge, type=float, metavar="DEG",
                    help="surcharge le demi-angle (degrés)")
    ap.add_argument("--portee", action=Surcharge, type=float, metavar="M",
                    help="surcharge la portée (m)")
    ap.add_argument("--plafond", action=Surcharge, type=float, metavar="M",
                    help="surcharge le plafond (m au-dessus du sol)")
    args = ap.parse_args(argv)

    try:
        vue = K.lire_json(args.vue)
        pads = K.charger_pas_de_tir(vue)
        zones = K.charger_zones(vue)
    except K.ErreurConf as e:
        raise SystemExit(f"Configuration : {e}")

    try:
        groupes = C.charger_vols(vue.get("vols"), os.path.dirname(
            os.path.abspath(args.vue)), K.normaliser_rideau)
        doms = []
        for gr in groupes:
            d = K.domaine_depuis(gr.domaine, gr.id, gr.nom, gr.couleur)
            if d is not None:
                doms.append(d)
        surcharges = resoudre_surcharges(getattr(args, "sequence", None), doms)
    except (C.ErreurCSV, K.ErreurConf) as e:
        raise SystemExit(f"Lecture des vols : {e}")

    rendu = K.charger_rendu(vue)
    coul = rendu["couleurs"]

    # Repère local unique pour tout le KML, centré sur le pas de tir du premier
    # domaine ; les autres pas de tir s'y expriment normalement.
    # Origine du repère local, et surtout altitude de référence de tout ce qui
    # est dessiné en absolu. Sans domaine, on prend le pas de tir du premier
    # groupe plutôt qu'un pas de tir arbitraire : les trajectoires en partent.
    ref = None
    if doms:
        ref = doms[0].pas_de_tir
    elif groupes:
        ref = next((g_.pas_de_tir for g_ in groupes if g_.pas_de_tir), None)
    try:
        pad0 = K.choisir_pas_de_tir(pads, ref)
    except K.ErreurConf:
        pad0 = pads[0]
    enu = G.ENU(pad0.lat, pad0.lon)
    alt_sol = pad0.alt_m

    k = KmlBuilder(enu, alt_sol)
    nom = vue.get("nom", "Base de lancement")
    avertissements: List[str] = []

    c = coul["cercle"]
    k.style_ligne("s_cercle", c["trait"], c["epaisseur"], 0.9)
    c = coul["secteur"]
    k.style_ligne("s_secteur", c["trait"], c["epaisseur"], 0.95)
    c = coul["limites"]
    k.style_ligne("s_limites", c["trait"], c["epaisseur"], 0.95)

    # -- zones --------------------------------------------------------------
    k.ouvrir_dossier("Zones")
    for i, z in enumerate(zones):
        aff = K.normaliser_affichage(z.affichage, "zone",
                                     rendu["affichage"]["zone"])
        vz = K.volume_de(rendu, z.volume)
        cz = K.fusion(coul["terrain"], z.couleur)
        sid = f"s_zone{i}"
        k.style_poly(sid, cz["trait"], cz["fond"], cz["epaisseur"],
                     cz["alpha_fond"])
        k.style_poly(f"{sid}_v", cz["trait"], cz["fond"], cz["epaisseur"],
                     float(vz["alpha_mur"]))
        k.style_ligne(f"{sid}_l", cz["trait"], cz["epaisseur"], 0.9)
        pts = G.open_ring([enu.fwd(a, b) for a, b in z.points])
        desc = (f"Zone « {z.id} »<br/>"
                f"{'Contrainte des domaines' if z.contrainte else 'Décor'}<br/>"
                f"Surface : {abs(G.signed_area(pts)) / 1e4:.1f} ha"
                + (f"<br/>Plafond : {z.plafond_m:.0f} m" if z.plafond_m else ""))
        if aff["polygone"]:
            k.polygone(z.nom, sid, pts, 0.0, desc)
        # Chaque zone a son propre plafond, avec les mêmes réglages de volume
        # qu'un domaine.
        if z.plafond_m > 0.0:
            dessiner_volume(k, z.nom, f"{sid}_v", f"{sid}_l", pts, z.plafond_m,
                            vz["style"], desc,
                            murs=aff["volume"],
                            plafond_visible=aff["plafond"])
    k.fermer_dossier()

    # -- pas de tir ---------------------------------------------------------
    utilises = {d.pas_de_tir for d in doms}
    k.ouvrir_dossier("Pas de tir")
    for p in pads:
        if not K.normaliser_affichage(p.affichage, "pas_de_tir",
                                      rendu["affichage"]["pas_de_tir"])["repere"]:
            continue
        k.point(p.nom, enu.fwd(p.lat, p.lon),
                f"Identifiant : {p.id}<br/>Altitude sol : {p.alt_m:.0f} m"
                + ("<br/><b>Origine d'un domaine</b>" if p.id in utilises else ""))
    k.fermer_dossier()

    # -- domaines -----------------------------------------------------------
    resume_dom = []
    if not doms:
        # Cas légitime : l'analyseur n'a pas pu conclure pour ce groupe. On
        # dessine tout le reste plutôt que de refuser le fichier entier.
        avertissements.append(
            "aucun domaine dans la vue : seuls les zones, les pas de tir et "
            "les trajectoires sont dessinés.\n"
            "      Pour en obtenir un, fixe 'azimut_deg' et 'demi_angle_deg' "
            "sur le groupe, puis relance l'analyse.")
    for idom, dom in enumerate(doms):
        sur = surcharges[dom.id]
        az = float(sur.get("azimut", dom.azimut_deg))
        demi = float(sur.get("demi_angle", dom.demi_angle_deg))
        portee = float(sur.get("portee", dom.portee_max_m))
        plafond = float(sur.get("plafond", dom.plafond_m))

        try:
            padd = K.choisir_pas_de_tir(pads, dom.pas_de_tir)
            zoned = K.choisir_zone(zones, dom.zone)
        except K.ErreurConf as e:
            raise SystemExit(f"domaine '{dom.id}' : {e}")

        P = enu.fwd(padd.lat, padd.lon)
        dz = padd.alt_m - alt_sol
        ring = G.ensure_ccw(G.open_ring(
            [enu.fwd(a, b) for a, b in zoned.points]))

        phi0 = G.az_to_phi(az)
        phi_a, phi_b = phi0 - math.radians(demi), phi0 + math.radians(demi)
        A = (P[0] + portee * math.cos(phi_a), P[1] + portee * math.sin(phi_a))
        B = (P[0] + portee * math.cos(phi_b), P[1] + portee * math.sin(phi_b))

        domaine = G.clip_secteur(ring, P, phi_a, phi_b)
        if len(domaine) < 3:
            raise SystemExit(
                f"domaine '{dom.id}' : le découpage a produit un polygone vide.")
        hors = G.sous_intervalles_hors_terrain(ring, portee, phi_a, phi_b, P)

        aff = K.normaliser_affichage(dom.affichage, "domaine",
                                     rendu["affichage"]["domaine"])
        vd = K.volume_de(rendu, dom.volume)
        rgb = dom.couleur or coul["domaine"]["trait"]
        cd = dict(coul["domaine"], trait=rgb, fond=rgb)
        sid = f"s_dom{idom}"
        k.style_poly(sid, cd["trait"], cd["fond"], cd["epaisseur"],
                     cd["alpha_fond"])
        k.style_poly(f"{sid}_v", cd["trait"], cd["fond"], cd["epaisseur"],
                     float(vd["alpha_mur"]))
        k.style_ligne(f"{sid}_l", cd["trait"], cd["epaisseur"], 0.95)

        k.ouvrir_dossier(f"{dom.nom} — {az:.1f}° ± {demi:.1f}°"
                         + (" [FORCÉ]" if dom.force else ""),
                         ouvert=(len(doms) == 1))
        desc = (f"Pas de tir : {padd.nom}<br/>Zone : {zoned.nom}<br/>"
                + (f"Groupe : {dom.groupe}<br/>" if dom.groupe else "")
                + f"Azimut : {az:.1f}°<br/>Demi-angle : ± {demi:.1f}°<br/>"
                  f"Portée : {portee:.0f} m<br/>"
                  f"Surface : {abs(G.signed_area(domaine)) / 1e4:.1f} ha"
                + (f"<br/>Plafond : {plafond:.0f} m" if plafond > 0 else ""))
        if aff["secteur"]:
            k.polygone(f"Domaine (R = {portee:.0f} m)", sid, domaine, 0.0, desc)

        for label, pt, phi in (("A", A, phi_a), ("B", B, phi_b)):
            seg = [P, pt]
            d = f"Azimut {G.phi_to_az(phi):.1f}° — {portee:.0f} m"
            if aff["limites"]:
                if rendu["segments_pointilles"]:
                    k.lignes_multi(f"Limite {label}", "s_limites",
                                   G.dash_polyline(seg, float(rendu["tiret_m"]),
                                                   float(rendu["espace_m"])),
                                   description=d)
                else:
                    k.ligne(f"Limite {label}", "s_limites", seg, description=d)
            if aff["reperes"]:
                k.point(f"{dom.id} — {label}", pt, d)

        if aff["contour_secteur"]:
            k.ligne("Contour du secteur", "s_secteur",
                    [P] + G.arc_points(portee, phi_a, phi_b, P=P) + [P],
                    description="Secteur demandé, non découpé par la zone")
        if aff["cercle_portee"]:
            k.ligne(f"Cercle de portée ({portee:.0f} m)", "s_cercle",
                    G.circle_points(portee, P=P))
        if plafond > 0.0:
            dessiner_volume(k, f"Volume (plafond {plafond:.0f} m)",
                            f"{sid}_v", f"{sid}_l", domaine, plafond + dz,
                            vd["style"], desc,
                            murs=aff["volume"], plafond_visible=aff["plafond"])
        k.fermer_dossier()

        resume_dom.append((dom, padd, zoned, az, demi, portee, plafond,
                           abs(G.signed_area(domaine)) / 1e4, hors, sur))

    # -- trajectoires -------------------------------------------------------
    resume: List[tuple] = []
    if groupes:
        k.ouvrir_dossier("Trajectoires")
        for ig, groupe in enumerate(groupes):
            k.ouvrir_dossier(groupe.nom)
            for iv, (vol, opts) in enumerate(groupe.vols):
                indices, avert = C.decimer(vol, opts["nb_points"])
                if avert:
                    avertissements.append(avert)
                if not indices:
                    resume.append((groupe, vol, 0, "aucun", 0))
                    continue

                # Le groupe hérite du domaine analysé pour lui : même pas de
                # tir et même azimut, sauf redéfinition explicite sur le vol.
                dom_g = next((d for d in doms if d.groupe == groupe.id), None)
                pid = opts.get("pas_de_tir") or (
                    dom_g.pas_de_tir if dom_g else None)
                pv = pad0
                if pid:
                    try:
                        pv = K.choisir_pas_de_tir(pads, pid)
                    except K.ErreurConf as e:
                        raise SystemExit(f"vols, groupe '{groupe.id}' : {e}")
                origine = enu.fwd(pv.lat, pv.lon)
                dz = pv.alt_m - alt_sol

                if opts.get("azimut_deg") is not None:
                    az_v = float(opts["azimut_deg"])
                elif dom_g is not None:
                    az_v = dom_g.azimut_deg
                elif doms:
                    az_v = doms[0].azimut_deg
                else:
                    az_v = 0.0
                    avertissements.append(
                        f"{vol.nom} : aucun azimut disponible (ni domaine, ni "
                        f"'azimut_deg'), trajectoire projetée plein nord")
                phi_v = G.az_to_phi(az_v)
                affv = K.normaliser_affichage(opts.get("affichage"), "vol",
                                              rendu["affichage"]["vol"])
                rid = opts["rideau"]
                mode = C.mode_rideau(rid["motif"]) if affv["rideau"] else "aucun"
                sid = f"s_v{ig}_{iv}"
                k.style_traj(sid, opts["couleur"], 3, float(rid["alpha"]))

                pts3 = [(x + origine[0], y + origine[1], h + dz)
                        for x, y, h in C.points_3d(vol, indices, phi_v)]
                if affv["trajectoire"] or mode == "plein":
                    k.trajectoire(
                        vol.nom, sid, pts3, rideau=(mode == "plein"),
                        description=(f"Groupe : {groupe.nom}<br/>"
                                     f"Fichier : {os.path.basename(vol.chemin)}<br/>"
                                     f"Pas de tir : {pv.nom}<br/>"
                                     f"Azimut : {G.phi_to_az(phi_v):.1f}°<br/>"
                                     f"Apogée : {vol.apogee_m:.0f} m<br/>"
                                     f"Portée : {vol.portee_m:.0f} m<br/>"
                                     f"Points : {len(indices)} / {len(vol)}<br/>"
                                     f"Rideau : {mode}"))

                n_seg = 0
                angles = C.MOTIFS.get(mode) or []
                if angles:
                    pas = float(rid["pas_m"])
                    segs = [[(x + origine[0], y + origine[1], h + dz)
                             for x, y, h in s]
                            for s in C.segments_3d(
                                C.motif_rideau(vol, indices, mode, pas),
                                phi_v, float(rid["penetration_sol_m"]))]
                    n_seg = len(segs)
                    if segs:
                        # L'opacité porte ici sur le trait ; en mode "plein"
                        # elle porte sur le PolyStyle. Deux objets KML distincts.
                        sidm = f"{sid}_m"
                        k.style_ligne(sidm, opts["couleur"],
                                      float(rid["epaisseur"]), float(rid["alpha"]))
                        k.trajectoires_multi(
                            f"{vol.nom} — {mode}", sidm, segs,
                            description=(f"Motif : {mode} ("
                                         + ", ".join(f"{a:g}°" for a in angles)
                                         + f")<br/>Pas : {pas:.0f} m"
                                           f"<br/>Segments : {len(segs)}"))
                    else:
                        avertissements.append(
                            f"{vol.nom} : motif '{mode}' vide (pas_m = {pas:g})")

                if rid["trace_sol"] and affv["trace_sol"]:
                    k.ligne_sol(f"{vol.nom} — trace au sol", sid,
                                [(x + origine[0], y + origine[1], 0.0)
                                 for x, y, _ in C.trace_sol(vol, indices, phi_v)],
                                description="Projection au sol")

                resume.append((groupe, vol, len(indices), mode, n_seg))
            k.fermer_dossier()
        k.fermer_dossier()

    # Le titre ne peut pas reprendre l'azimut d'un domaine : il peut n'y en
    # avoir aucun, ou plusieurs avec des azimuts différents.
    if len(doms) == 1:
        d0 = resume_dom[0]
        titre_doc = f"Domaine de vol — {nom} — {d0[3]:.1f}°"
    elif doms:
        titre_doc = f"Domaines de vol — {nom} ({len(doms)})"
    else:
        titre_doc = f"{nom} — zones et trajectoires"

    out = args.out or os.path.splitext(args.vue)[0] + ".kml"
    with open(out, "w", encoding="utf-8") as f:
        f.write(k.render(titre_doc))

    # --- compte rendu ------------------------------------------------------
    print(f"Zones           : {len(zones)}   Pas de tir : {len(pads)}   "
          f"Domaines : {len(doms)}"
          + ("   (aucun domaine à dessiner)" if not doms else ""))
    for dom, padd, zoned, az, demi, portee, plafond, surf, hors, sur in resume_dom:
        print(f"  [{dom.id}] {dom.nom}"
              + (f"   (surchargé : "
                 + ", ".join(f"{k}={v:g}" for k, v in sorted(sur.items())) + ")"
                 if sur else ""))
        print(f"      Pas de tir  : {padd.nom} ({padd.id})   "
              f"zone {zoned.nom} ({zoned.id})")
        print(f"      Cône        : {az:.1f}° ± {demi:.1f}°   "
              f"portée {portee:.0f} m   surface {surf:.1f} ha"
              + (f"   plafond {plafond:.0f} m" if plafond > 0 else ""))
        if hors:
            tot = sum(math.degrees(G.largeur(iv)) for iv in hors)
            print(f"      [!] DÉBORDEMENT : {tot:.1f}° hors de la zone")
            for iv in hors:
                print(f"            {G.phi_to_az(iv[1]):6.1f}° -> "
                      f"{G.phi_to_az(iv[0]):6.1f}°"
                      f"   ({math.degrees(G.largeur(iv)):.1f}°)")
        else:
            print("      Conformité  : le secteur reste dans la zone")
    if resume:
        print("Trajectoires    :")
        gcourant = None
        for groupe, vol, n, mode, n_seg in resume:
            if groupe is not gcourant:
                print(f"  [{groupe.id}] {groupe.nom}"
                      f"   portée groupe {groupe.portee_m:.0f} m")
                gcourant = groupe
            etat = f"{n:4d} / {len(vol):4d} pts" if n else "  non tracé "
            sfx = f"   rideau {mode}" + (f" ({n_seg} segments)" if n_seg else "")
            print(f"      {vol.nom:<22s} {etat}   apogée {vol.apogee_m:7.0f} m"
                  f"   portée {vol.portee_m:7.0f} m{sfx}")
    for a in avertissements:
        print(f"[!] {a}")
    print(f"KML écrit       : {out}")
    return 0


if __name__ == "__main__":
    # Les erreurs de configuration remontent de partout (affichage, rideau,
    # références) : un seul filet, pour ne jamais afficher de traceback.
    try:
        sys.exit(main())
    except (K.ErreurConf, C.ErreurCSV) as err:
        raise SystemExit(f"Configuration : {err}")
