# Domaine de vol — analyse et visualisation Google Earth

Deux outils en chaîne, reliés par un fichier JSON intermédiaire.

```
etude.json  ──▶  analyse_terrain.py  ──▶  vue.json  ──▶  genere_kml.py  ──▶  domaine.kml
(écrit à          (RECHERCHE)            (éditable)      (CONSTRUCTION)      (Google Earth)
 la main)
```

---

## Pourquoi deux JSON

Les deux fichiers ne jouent pas le même rôle :

| Fichier | Qui l'écrit | Contient | Change quand ? |
|---|---|---|---|
| `etude.json` | **toi, une fois** | pas de tir, zones, vols | jamais, ou presque |
| `vue.json` | **l'analyseur** | l'étude + azimut, demi-angle, portée résolus | à chaque analyse, ou à la main |

`etude.json` décrit **ce qui est étudié**. `vue.json` décrit **ce qui est
dessiné** : fichier jetable, régénérable, et modifiable entre les deux étapes.
C'est là qu'on remplace 217,6° par 230°.

La vue recopie l'étude à l'intérieur d'elle-même : elle est autonome,
archivable et transmissible seule.

> Les noms sont une convention, pas une contrainte : les deux CLI acceptent
> n'importe quel chemin.

---

## Installation

Python 3.8+, aucune dépendance. Les modules `dv_*.py` doivent être dans le
même dossier que les deux exécutables.

```
domaine_vol/
├── dv_geo.py             # géométrie      — ne s'exécute pas seul
├── dv_csv.py             # lecture CSV    — ne s'exécute pas seul
├── dv_conf.py            # configuration  — ne s'exécute pas seul
├── analyse_terrain.py    # étape 1
├── genere_kml.py         # étape 2
├── etude.schema.json     # JSON Schema de l'étude
├── vue.schema.json       # JSON Schema de la vue
├── etude.json            # tes données
└── vols/
    ├── alpha/*.csv
    └── beta/*.csv
```

---

## Utilisation en 30 secondes

```bash
python3 analyse_terrain.py etude.json -o vue.json   # 1. analyser
                                                    # 2. éditer vue.json si besoin
python3 genere_kml.py vue.json -o domaine.kml       # 3. générer
```

Puis ouvrir `domaine.kml` dans **Google Earth Pro (desktop)**. La version web
gère mal l'extrusion et les polygones en altitude absolue.

---

## Structure : zones, pas de tir, domaines

`pas_de_tir` et `zones` sont des listes. Tout est dessiné.

```json
"pas_de_tir": [
  { "id": "rampe_a", "nom": "Rampe A", "lat": 43.2184361, "lon": -0.0473333, "alt_m": 350.0 },
  { "id": "rampe_b", "nom": "Rampe B", "lat": 43.2172000, "lon": -0.0492000, "alt_m": 348.0 }
],
"zones": [
  { "id": "perimetre", "nom": "Périmètre de la base", "contrainte": true, "points": [ ... ] },
  { "id": "public", "nom": "Zone public", "plafond_m": 200.0,
    "volume": { "style": "plein", "alpha_mur": 0.10 },
    "couleur": { "trait": "#00c060", "fond": "#00c060", "alpha_fond": 0.20 },
    "points": [ ... ] }
]
```

**Un domaine appartient toujours à un groupe de vols.** Il en tire sa portée,
son pas de tir et sa zone — le représenter ailleurs obligerait à maintenir une
correspondance par identifiant pour rien. Le groupe porte donc tout :

```json
"vols": {
  "zone": "perimetre",          // défaut pour tous les groupes
  "plafond_m": 3000.0,
  "groupes": [
    { "id": "beta",  "pas_de_tir": "rampe_a", ... },
    { "id": "alpha", "pas_de_tir": "rampe_b", "zone": "perimetre", ... }
  ]
}
```

L'analyseur écrit le résultat **dans le groupe**, sous la clé `domaine` :

```json
"groupes": [
  { "id": "beta", "pas_de_tir": "rampe_a",
    "domaine": { "nom": "Fusée beta", "pas_de_tir": "rampe_a", "zone": "perimetre",
                 "couleur": "#ff3020", "azimut_deg": 217.6, "demi_angle_deg": 24.8,
                 "portee_max_m": 2486.7, "plafond_m": 3000.0 },
    "liste": [ ... ] }
]
```

> Il n'y a plus d'objet `tir`. Ce qu'il portait est réparti : `zone` et
> `plafond_m` sur le bloc `vols` ou sur chaque groupe, `pas_de_tir` sur le
> groupe.

Sans aucun groupe de vols, il n'y a pas de domaine — le KML ne contient que
les zones et les pas de tir.

### Résolution des références

Pour le pas de tir d'un groupe, dans l'ordre :

1. `--pas-de-tir` (force tous les groupes) ;
2. le champ `pas_de_tir` du groupe ;
3. le champ `pas_de_tir` du bloc `vols` ;
4. l'élément marqué `"reference"` ;
5. s'il n'y en a qu'un, celui-là.

Même logique pour la zone (drapeau `contrainte`). Sinon l'analyseur s'arrête
en listant les identifiants disponibles. Deux éléments marqués est aussi une
erreur : mieux vaut refuser que choisir au hasard.

Les zones non marquées `contrainte` sont décoratives — zone public, bâtiments,
exclusions.

---

## Imposer un cône, et le cas où rien ne tient

Un groupe peut fixer son propre cône :

```json
{ "id": "beta", "azimut_deg": 217.0, "demi_angle_deg": 20.0 }
```

| Analyse | `azimut_deg` | `demi_angle_deg` | Résultat |
|---|---|---|---|
| réussie | — | — | recommandation : bissectrice du plus large secteur |
| réussie | fixé | — | azimut imposé, demi-angle exploitable calculé autour |
| réussie | — | fixé | azimut recommandé, ouverture imposée |
| réussie | fixé | fixé | les deux imposés ; débordement signalé le cas échéant |
| **échouée** | fixé | fixé | **domaine forcé**, marqué `"force": true`, avertissement |
| **échouée** | fixé | — | pas de domaine — l'ouverture n'est pas déductible |
| **échouée** | — | — | pas de domaine — l'azimut n'est pas déductible |

**Un échec n'interrompt plus l'exécution.** Les autres groupes sont analysés
normalement, et l'avertissement part sur `stderr` :

```
[!] groupe « beta » : aucun azimut ne permet 4282 m sans sortir de la zone
      (sortie la plus lointaine : 146 m environ).
      Aucun domaine produit. Pour en forcer un, fixe 'azimut_deg' et
      'demi_angle_deg' sur le groupe.
```

Avec les deux valeurs fixées, le domaine est produit et le générateur le
signale — c'est la seule façon de *voir* une configuration qui ne tient pas,
ce qui est précisément ce qu'on cherche à regarder :

```
  [beta] Fusée beta
      Cône : 217.0° ± 20.0°   portée 4282 m
      [!] DÉBORDEMENT : 40.0° hors de la zone
```

Le dossier KML est alors suffixé `[FORCÉ]`.

L'analyseur sort en code 1 si aucun domaine n'a pu être produit.

**Le générateur accepte une vue sans aucun domaine.** Il dessine les zones,
les pas de tir et les trajectoires, et le signale :

```
Zones           : 2   Pas de tir : 2   Domaines : 0   (aucun domaine à dessiner)
[!] aucun domaine dans la vue : seuls les zones, les pas de tir et les
      trajectoires sont dessinés.
      Pour en obtenir un, fixe 'azimut_deg' et 'demi_angle_deg' sur le groupe,
      puis relance l'analyse.
```

Les trajectoires sont alors projetées sur l'`azimut_deg` du groupe s'il en a
un, et l'altitude de référence est celle du pas de tir du premier groupe.
Sans azimut nulle part, elles partent plein nord, avec un avertissement.

---

## Affichage : tout est débrayable

Chaque élément graphique s'allume ou s'éteint indépendamment, par un bloc
`affichage`. Un **booléen seul** allume ou éteint tout l'élément — c'est le
geste le plus fréquent, il ne doit pas obliger à énumérer les clés.

| Élément | Clés |
|---|---|
| zone | `polygone`, `volume`, `plafond` |
| domaine | `secteur`, `limites`, `reperes`, `contour_secteur`, `cercle_portee`, `volume`, `plafond` |
| vol | `trajectoire`, `rideau`, `trace_sol` |
| pas de tir | `repere` |

```json
"domaine": { "affichage": { "cercle_portee": false, "contour_secteur": false } },
"zones":   [ { "affichage": { "volume": false } } ],
"liste":   [ { "fichier": "beta_80.csv", "affichage": { "rideau": false } },
             { "fichier": "alpha_45.csv", "affichage": false } ]
```

**Tout est écrit en clair dans `vue.json`.** L'analyseur y matérialise les
valeurs par défaut au lieu de les laisser implicites : sans cela, il faudrait
aller chercher le nom d'une clé dans cette documentation pour basculer un
affichage. Le fichier est fait pour être relu et modifié, donc il porte tous
ses réglages.

```json
"zones": [ { "id": "public", "plafond_m": 200.0,
             "volume": { "style": "plein", "alpha_mur": 0.1 },
             "affichage": { "polygone": true, "volume": true, "plafond": true } } ],
"groupes": [ { "id": "beta",
    "domaine": { "azimut_deg": 217.618, "demi_angle_deg": 24.829,
                 "portee_max_m": 2486.7, "plafond_m": 3000.0,
                 "volume": { "style": "aretes", "alpha_mur": 0.15 },
                 "affichage": { "secteur": true, "limites": true, "reperes": true,
                                "contour_secteur": true, "cercle_portee": true,
                                "volume": true, "plafond": true } },
    "liste": [ { "fichier": "beta_80.csv",
                 "rideau": { "motif": "hachures_90", "pas_m": 150.0, "epaisseur": 1.5,
                             "alpha": 0.35, "penetration_sol_m": 30.0, "trace_sol": true },
                 "affichage": { "trajectoire": true, "rideau": true, "trace_sol": true } } ] } ]
```

Les mêmes clés existent dans `rendu.affichage` pour fixer les défauts au
niveau de l'étude. Une clé inconnue est refusée avec la liste des clés
attendues, plutôt qu'ignorée :

```
Configuration : affichage (vol) : clé(s) inconnue(s) trace — attendu : rideau, trace_sol, trajectoire
```

---

## Les vols

### Groupes

Un groupe = une configuration de fusée. Chaque groupe a sa propre portée : la
plus grande distance latérale parmi ses vols marqués `balistique` (tous ses
vols si aucun ne l'est).

```json
"vols": {
  "dossier": "./vols",
  "colonnes": { "temps": 1, "altitude": 2, "distance_laterale": 7 },
  "nb_points": 150,
  "rideau": { "motif": "hachures_90", "pas_m": 100.0, "alpha": 0.35, "trace_sol": true },

  "groupes": [
    {
      "id": "beta", "nom": "Fusée beta", "dossier": "beta",
      "couleur": "#ff3020", "reference": true,
      "rideau": { "motif": "quadrillage_45", "pas_m": 150.0 },
      "liste": [
        { "fichier": "beta_45.csv", "nom": "beta 45 balistique", "balistique": true },
        { "fichier": "beta_80.csv", "nom": "beta 80 nominal", "couleur": "#22d0ff", "rideau": "hachures_90" }
      ]
    },
    {
      "id": "alpha", "nom": "Fusée alpha", "dossier": "alpha",
      "couleur": "#ff9020",
      "rideau": { "motif": "hachures_45" },
      "liste": [
        { "fichier": "alpha_45.csv", "nom": "alpha 45 balistique", "balistique": true },
        { "fichier": "alpha_80.csv", "nom": "alpha 80 nominal", "couleur": "#66ff88",
          "nb_points": 40, "rideau": { "motif": "quadrillage_0", "pas_m": 120.0 } }
      ]
    }
  ]
}
```

Chaque groupe devient un `Folder` KML, donc affichable ou masquable d'un clic
dans Google Earth.

```
  Groupes de vols
  <== beta   Fusée beta    2 vol(s)   portée    2487 m   (beta 45 balistique)
          beta 45 balistique     495 pts   apogée  954 m   portée 2487 m  [balistique]
          beta 80 nominal        641 pts   apogée 2291 m   portée  855 m
      alpha  Fusée alpha   2 vol(s)   portée    1344 m   (alpha 45 balistique)
```

Le groupe dimensionnant est choisi comme les autres références, avec une
nuance au dernier cran : **à défaut de `"reference"`, c'est le groupe de plus
grande portée qui est retenu**, pas le premier. Un domaine de sécurité
couvrant plusieurs fusées doit couvrir la pire.

Sans `groupes`, une `liste` à la racine du bloc produit un groupe unique
implicite — les configurations à une seule fusée restent courtes.

### Héritage en cascade

Les réglages descendent **bloc `vols` → groupe → vol**, et chaque niveau peut
tout redéfinir. Un groupe peut venir d'un autre export que le précédent, un
vol d'un format différent du reste de son groupe.

Sont héritables : `colonnes`, `separateur`, `commentaire`, `nb_points`,
`couleur`, `azimut_deg`, `pas_de_tir`, `balistique`, `rideau`.

Dans l'exemple ci-dessus, `alpha 80 nominal` hérite les colonnes du bloc, la
couleur qu'il redéfinit, le motif qu'il redéfinit, et `nb_points` qu'il
redéfinit aussi.

`rideau` se fusionne clé par clé au lieu d'être remplacé en bloc : un groupe
qui ne change que `pas_m` garde le motif et l'alpha hérités.

### Colonnes déclarées, 1-based

Aucune détection automatique. Les exports changent d'ordre, de langue et
d'entête d'une version à l'autre, et deviner produit des erreurs silencieuses
— une trajectoire tracée avec la vitesse à la place de l'altitude a l'air
plausible.

```json
"colonnes": { "temps": 1, "altitude": 2, "distance_laterale": 7 }
```

`temps` est facultatif (lu, pas encore exploité). Si l'altitude est en
colonne 3 et la distance latérale en colonne 5 :

```json
"colonnes": { "altitude": 3, "distance_laterale": 5 }
```

**Toute ligne commençant par `commentaire` est ignorée**, ce qui couvre
l'entête et les marqueurs d'événements :

```
# Temps (s),Altitude (m),...,Distance latérale (m)
# Event LIFTOFF occurred at t=0.06 seconds
0.06,0.03,1.667,2.744,56.529,79.943,0.043
```

Les EVENT ne sont pas exploités.

### Nombre de points affichés

| `nb_points` | Effet |
|---|---|
| absent ou `null` | **tous** les points du CSV (défaut) |
| `> 0` | répartition régulière sur toute la trajectoire |
| `>= total` | limité au total, **avec un avertissement** |
| `0` ou négatif | vol **non tracé**, avec un avertissement |

Pas de maximum imposé : l'application clippe et le signale.

Premier et dernier point sont toujours conservés, et **l'apogée est forcée**
dans la sélection en remplaçant le point retenu le plus proche. Le nombre
demandé est respecté, et le sommet n'est pas raboté par l'échantillonnage.

### Projection en 3D

Les CSV ne donnent qu'un profil dans le plan de tir (altitude, distance
latérale) — pas de cap. Chaque trajectoire est projetée **en ligne droite le
long d'un azimut**, celui de la vue par défaut, surchargeable par
`azimut_deg`. Le point de départ est le pas de tir du tir, ou celui désigné
par `pas_de_tir` sur le vol ou le groupe.

---

## Le rideau

Le rideau est la surface verticale sous la trajectoire. Comme celle-ci est
projetée en ligne droite, cette surface se décrit entièrement dans le plan
**(distance latérale, altitude)**. Tous les motifs y sont construits en 2D,
puis relevés en 3D.

**Tous les réglages tiennent dans un objet :**

```json
"rideau": {
  "motif": "quadrillage_45",
  "pas_m": 150.0,
  "epaisseur": 1.5,
  "alpha": 0.35,
  "penetration_sol_m": 30.0,
  "trace_sol": true
}
```

Raccourcis acceptés : `"rideau": "hachures_45"` (motif seul), `true`
(= `"plein"`), `false` (= `"aucun"`).

| Motif | Angles | Rendu |
|---|---|---|
| `"aucun"` | — | trajectoire seule (défaut) |
| `"plein"` | — | mur continu, `extrude` KML |
| `"hachures_90"` | 90° | verticales |
| `"hachures_45"` | 45° | diagonales montantes |
| `"hachures_135"` | 135° | diagonales descendantes |
| `"hachures_0"` | 0° | lignes de niveau |
| `"quadrillage_0"` | 0° + 90° | grille droite |
| `"quadrillage_45"` | 45° + 135° | grille diagonale |

`"hachures"` reste un alias de `"hachures_90"`.

**KML n'a pas de motif de remplissage.** `PolyStyle` n'expose que `color`,
`colorMode`, `fill` et `outline` — ni hachure, ni texture. Les motifs sont
donc réellement dessinés : chaque famille est un faisceau de droites
parallèles espacées de `pas_m` **perpendiculairement**, découpé sur la
silhouette du rideau. Tous les segments d'un vol tiennent dans un seul
`MultiGeometry`.

L'angle est **métrique** : à 45°, un mètre horizontal vaut un mètre vertical,
ce qui donne bien 45° à l'écran sans exagération verticale.

Le découpage se fait par échantillonnage puis dichotomie aux bords. La
silhouette d'une trajectoire n'est ni convexe ni monotone ; un clipping
analytique serait fragile pour un gain nul à l'affichage.

### Trois réglages qui comptent

**`pas_m` est mesuré dans le plan (d, h)**, donc en distance latérale pour les
hachures verticales : c'est l'écart que l'œil perçoit sur la carte. Un pas
mesuré sur la longueur 3D parcourue paraît irrégulier — sur un arc balistique
il produit un écart horizontal allant de 16 à 100 m pour un pas nominal de
100 m, soit un rapport de 6.

**`alpha` porte sur le trait en mode motif**, et sur le `PolyStyle` en mode
`"plein"`. Ce sont deux objets KML différents : en mode motif il n'y a aucun
polygone, donc rien pour un `PolyStyle`.

**`penetration_sol_m` enfonce sous le sol les pieds des segments.** Le rideau
est dessiné en altitude absolue depuis `alt_m`, l'altitude déclarée du pas de
tir. Si elle diffère du relief réel, un pied posé pile à h = 0 flotte ou
disparaît selon la pente. 30 m de pénétration masquent le défaut sans fausser
les altitudes utiles — mais renseigner correctement `alt_m` reste préférable.

`trace_sol` est tracée en `clampToGround` : elle suit le relief au lieu de
rester à plat.

---

## Le volume sous plafond

Le plafond est une **contrainte physique**, pas un réglage d'affichage. Il est
porté par l'élément concerné et exprimé en mètres **au-dessus du sol**.

**Sur un groupe de vols** — c'est le plafond réglementaire du vol, vérifié par
l'analyseur :

```json
"vols": { "plafond_m": 3000.0, "groupes": [ ... ] }
```

```
  Plafond de vol : 3000 m au-dessus du sol du pas de tir
    beta 45 balistique       apogée     954 m   marge    +2046 m   [OK ]
    beta 80 nominal          apogée    2291 m   marge     +709 m   [OK ]
```

À 1500 m, `beta 80 nominal` ressort en `[DÉPASSEMENT]` avec −791 m.

**Sur une zone** — hauteur de bâtiment, plafond d'une zone public, volume
d'exclusion. Mêmes réglages de rendu qu'un domaine :

```json
{ "id": "public", "plafond_m": 200.0, "volume": { "style": "plein", "alpha_mur": 0.10 } }
```

### Rendu

```json
"volume": { "style": "aretes", "alpha_mur": 0.15 }
```

| Style | Rendu |
|---|---|
| `"aretes"` | une verticale par sommet, en fil de fer (défaut) |
| `"plein"` | polygone extrudé : murs et couvercle |

`rendu.volume` donne les valeurs par défaut ; chaque zone et chaque domaine
peut les redéfinir. Ce qui s'affiche se règle par `affichage.volume` et
`affichage.plafond`, pas par le bloc `volume`. Le contour du plafond est tracé dans les deux styles —
c'est la ligne que l'on vient lire, et un mur translucide seul ne la donne pas
nettement. `affichage.volume` et `affichage.plafond` les débrayent séparément.

**`"aretes"` est le défaut, et c'est important.** Extruder une empreinte de
plusieurs centaines d'hectares sur 3000 m de haut produit un pavé qui masque
tout ce qu'il est censé contenir, trajectoires comprises. Un fil de fer se lit
sans rien cacher. Passe en `"plein"` quand tu veux la silhouette plutôt que le
contenu, avec `alpha_mur` vers 0,10.

Sans `plafond_m`, l'élément reste plat.

---

## Étape 1 — `analyse_terrain.py`

Cherche tous les secteurs angulaires dans lesquels un tir de portée `R` reste
dans la zone de contrainte, retient le plus large, et recommande sa
bissectrice : l'azimut qui laisse la plus grande marge de part et d'autre.

```
  Portée retenue    : 2487 m   (groupe « Fusée beta », vol « beta 45 balistique »)

  Secteurs admissibles pour R = 2487 m
     192.8° ->  242.4°   ouverture   49.7°   bissectrice  217.6° <== le plus large

  RECOMMANDATION
    Azimut de tir   : 217.6°
    Demi-angle      : ± 24.8°
    Sortie sur axe  : 3756 m  (marge 1269 m)
```

Le rapport donne aussi la distance de sortie tous les 30°, utile pour
comprendre la forme du site (ici 146 m plein est, 3736 m à 210°).

| Option | Effet |
|---|---|
| `--pas-de-tir ID` | force le pas de tir de tous les groupes |
| `--zone ID` | force la zone de contrainte de tous les groupes |
| `--groupe ID` | n'analyser que ce groupe (défaut : tous) |
| `--portee M` | portée imposée, surcharge tout |
| `--plafond M` | plafond de vol, en m au-dessus du sol |
| `--azimut DEG` | évalue un azimut imposé et affiche ses marges. **Répétable** |
| `-o FICHIER` | écrit la vue. Sans cette option, l'analyseur ne fait qu'afficher |
| `--rapport FICHIER` | écrit aussi le rapport en texte |

Pour imposer un azimut de façon durable, mets `azimut_deg` sur le groupe :
c'est une propriété de la configuration, pas d'une invocation.

Ordre de priorité de la portée : `--portee` > `tir.portee_max_m` > groupe.

### Évaluer sans rien générer

```bash
python3 analyse_terrain.py etude.json --azimut 230 --azimut 250
```

```
  AZIMUTS IMPOSÉS
     230.0°  admissible
             marge vers 242.4° : 12.4°
             marge vers 192.8° : 37.2°
             demi-angle symétrique exploitable : ± 12.4°
     250.0°  HORS DOMAINE — sortie à 2328 m pour R = 2487 m
```

---

## Étape 2 — `genere_kml.py`

Ne cherche rien, ne recommande rien. On lui donne un azimut et un demi-angle,
il dessine le secteur découpé par la zone.

**Il ne refuse jamais.** Si le secteur sort de la zone, il le dessine quand
même et le signale — c'est le but : pouvoir visualiser une configuration non
conforme.

```bash
python3 genere_kml.py vue.json --azimut 230
```

```
[!] DÉBORDEMENT : 12.6° du secteur sortent de la zone avant 2487 m
       242.4° ->  254.5°   (12.6°)
```

| Option | Effet |
|---|---|
| `-o FICHIER` | KML de sortie (défaut : nom de la vue avec `.kml`) |
| `--domaine ID` | ouvre une portée pour les surcharges qui suivent. **Répétable** |
| `--azimut DEG` | surcharge l'azimut, sans modifier la vue |
| `--demi-angle DEG` | surcharge le demi-angle |
| `--portee M` | surcharge la portée |
| `--plafond M` | surcharge le plafond |

Les quatre dernières valent pour le domaine ouvert par le `--domaine`
précédent, ou pour tous si aucun n'est ouvert.

Ces options servent aux essais rapides. Pour un réglage à garder, édite
`vue.json`.

### Surcharges ciblées

`--domaine` **ouvre une portée** : les surcharges qui le suivent ne valent que
pour ce domaine. L'option est répétable, et l'ordre de la ligne de commande
est respecté.

```bash
python3 genere_kml.py vue.json --domaine alpha --azimut 300 --domaine beta --azimut 250
```

```
  [beta] Fusée beta     (surchargé : azimut=250)
      Cône : 250.0° ± 24.8°   portée 2487 m
  [alpha] Fusée alpha   (surchargé : azimut=300)
      Cône : 300.0° ± 70.9°   portée 1344 m
```

Une surcharge placée **avant tout `--domaine`** s'applique à **tous** les
domaines. Les deux se combinent :

```bash
python3 genere_kml.py vue.json --plafond 1200 --domaine beta --azimut 200 --portee 1000
```

```
  [beta] Fusée beta     (surchargé : azimut=200, plafond=1200, portee=1000)
  [alpha] Fusée alpha   (surchargé : plafond=1200)
```

Les surcharges appliquées sont rappelées dans le compte rendu : avec une ligne
de commande qui en porte plusieurs, il faut pouvoir vérifier ce qui a
réellement été pris.

> **Pourquoi une action argparse maison.** Une option ordinaire est écrasée par
> sa dernière occurrence : `--domaine alpha --azimut 300 --domaine beta
> --azimut 250` n'aurait retenu que `beta` et `250`, en perdant la surcharge
> sur `alpha` sans rien dire. La séquence est donc enregistrée telle qu'elle
> est tapée.

### Tirer à 230° « parce que pourquoi pas »

```bash
# a) essai jetable, la vue n'est pas touchée
python3 genere_kml.py vue.json --azimut 230 -o essai_230.kml

# b) réglage gardé : éditer "azimut_deg": 230 dans vue.json, puis
python3 genere_kml.py vue.json

# c) depuis l'analyseur, avec le demi-angle conforme calculé pour 230°
# (ou "azimut_deg": 230 sur le groupe dans etude.json, puis relancer l'analyse)
```

En (a) et (b), le demi-angle reste celui de la vue (± 24,8°), calculé pour
217,6°. À 230° il déborde de 12,6°. Le demi-angle réellement exploitable est
± 12,4° — donné par `--azimut 230`. La variante (c) le calcule seule.

---

## Lire le KML dans Google Earth

| Élément | Aspect | Signification |
|---|---|---|
| Zones | polygones remplis clair, volume optionnel | périmètre et zones décoratives |
| Pas de tir | repères | tous, celui analysé est signalé dans sa bulle |
| Domaines | polygones remplis, un `Folder` chacun | secteur **découpé par la zone**, dans la couleur du groupe |
| Limites A / B | pointillés blancs | les deux azimuts extrêmes |
| Contour du secteur | trait orange non rempli | secteur **demandé**, non découpé |
| Cercle de portée | trait bleu | rayon `portee_max_m` |
| Volume | arêtes verticales + contour au plafond | domaine sous `tir.plafond_m` |
| Trajectoires | traits colorés 3D, un `Folder` par groupe | un par vol |

**Le débordement se lit à l'œil** : c'est l'écart entre le contour orange et
le remplissage rouge. S'ils se superposent, la configuration est conforme.

---

## JSON Schema

`etude.schema.json` et `vue.schema.json` sont en draft 2020-12. Ils décrivent
tous les champs, valeurs par défaut et énumérations, et refusent les clés
inconnues — une faute de frappe dans un nom de réglage est signalée au lieu
d'être ignorée.

Ajoute la ligne suivante en tête du fichier pour l'autocomplétion et la
validation à la volée dans VS Code :

```json
"$schema": "./etude.schema.json"
```

En ligne de commande :

```bash
pip install check-jsonschema
check-jsonschema --schemafile etude.schema.json etude.json
```

Les schémas ne sont pas utilisés à l'exécution — les outils restent sans
dépendance et valident eux-mêmes ce dont ils ont besoin, avec des messages
adaptés au contexte.

---

## Messages d'erreur

| Message | Cause | Que faire |
|---|---|---|
| `2 pas de tir déclarés et aucun marqué 'reference'` | référence ambiguë | `--pas-de-tir ID`, ou marquer `"reference": true` |
| `zone 'X' introuvable — disponibles : ...` | identifiant erroné | reprendre un identifiant de la liste |
| `Le pas de tir est hors de la zone de contrainte` | lat/lon inversés, ou mauvaise zone | vérifier l'ordre `[lat, lon]` |
| `Le disque est entièrement contenu dans la zone` | portée trop faible, la zone ne contraint rien | augmenter la portée, ou fixer azimut et demi-angle à la main |
| `Aucun azimut ne permet une portée de N m` | zone trop petite | la config ne tient pas sur ce site |
| `aucune ligne de données exploitable` | mauvais mapping de colonnes | le message donne le nombre de colonnes détectées |
| `rideau : clé(s) inconnue(s)` | faute de frappe dans un réglage | le message liste les clés attendues |

---

## Ce qui n'est pas encore branché

- **Pas de rejeu temporel.** La colonne `temps` est lue mais inutilisée. Un
  `gx:Track` donnerait le curseur temporel de Google Earth.
- **Pas de découpage par phase.** Une trajectoire = un trait d'une couleur.
  Colorier par phase demanderait d'exploiter les marqueurs `# Event`.
- **Un domaine par groupe de vols.** Sans groupe, pas de domaine.
- **Le plafond ne borne pas le domaine.** Il est vérifié et dessiné, mais un
  vol qui le dépasse ne réduit pas le secteur admissible — les deux
  contraintes sont indépendantes.
- **Aucune prise en compte du vent** ni de la dispersion réelle. Le cône
  recommandé est purement géométrique.
