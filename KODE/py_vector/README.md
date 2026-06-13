# Vector Orientation Visualizer 🎯

Lecteur COM port (Windows/Linux) + visualiseur 3D en temps réel pour afficher l'orientation d'un vecteur.

## 📋 Caractéristiques

✅ **Visualisation 3D temps réel** du vecteur dans l'espace  
✅ **Historique** de la trajectoire (trail)  
✅ **Sphère magnitude** montrant l'amplitude du vecteur  
✅ **Axes de référence** (X=rouge, Y=vert, Z=bleu)  
✅ **Statistiques en direct** : magnitude, angles, paquets reçus  
✅ **Formats flexibles** : CSV, JSON, espace-séparé, tab-séparé  
✅ **Multi-plateforme** : Windows, Linux, macOS  

## 🚀 Installation

### 1. Pré-requis

```bash
# Python 3.7+
python --version

# Installer les dépendances
pip install pyserial matplotlib numpy scipy
```

### 2. Télécharger les scripts

```bash
# Les deux fichiers dans le même répertoire
vector_visualizer.py
com_simulator.py  # Pour tester sans matériel
```

## 💻 Utilisation

### Mode 1 : Avec vrai matériel (IMU, BMI088, etc.)

#### 1.1 Identifier le port COM

```bash
# Windows
python vector_visualizer.py --list-ports

# Sortie exemple :
# Available COM ports:
#   • COM3: STM32 Virtual COM Port

python vector_visualizer.py --port COM3 --baudrate 115200
```

#### 1.2 Format de données attendu

Le script accepte plusieurs formats. Envoyez des données depuis votre microcontrôleur :

**CSV (recommandé):**
```
10.5,20.3,-15.2
-5.1,8.4,12.6
```

**Space-separated:**
```
10.5 20.3 -15.2
-5.1 8.4 12.6
```

**Tab-separated:**
```
10.5	20.3	-15.2
-5.1	8.4	12.6
```

**JSON:**
```json
{"x": 10.5, "y": 20.3, "z": -15.2}
{"x": -5.1, "y": 8.4, "z": 12.6}
```

### Mode 2 : Test avec simulateur (sans matériel)

Idéal pour tester sans IMU physique.

#### 2.1 Setup Windows (COM virtuel)

**Installation (une fois):**
1. Télécharger [com0com](https://sourceforge.net/projects/com0com/)
2. Installer et redémarrer
3. Lancer `com0com → Setup`
4. Cliquer "Add Pair"
5. Vous obtiendrez 2 ports, ex: CNCA0 ↔ CNCB0

**Utilisation:**
```bash
# Terminal 1 : Simulator sur CNCB0
python com_simulator.py --port CNCB0 --mode rotating

# Terminal 2 : Visualizer sur CNCA0
python vector_visualizer.py --port CNCA0
```

#### 2.2 Setup Linux (socat)

```bash
# Installer socat
sudo apt install socat

# Créer une paire virtuelle
socat -d -d pty,raw,echo=0 pty,raw,echo=0

# Sortie :
# N set to 4
# /dev/pts/7 <-- /dev/pts/8

# Terminal 1 : Simulator sur /dev/pts/8
python com_simulator.py --port /dev/pts/8 --mode rotating

# Terminal 2 : Visualizer sur /dev/pts/7
python vector_visualizer.py --port /dev/pts/7
```

#### 2.3 Setup macOS

```bash
# Installer socat via Homebrew
brew install socat

# Créer une paire virtuelle
socat -d -d pty,raw,echo=0 pty,raw,echo=0

# Utiliser les ports générés (même que Linux)
```

### Mode 3 : Avec Arduino/STM32 physique

#### 3.1 Code de transmission (Arduino)

```cpp
// Example: Envoyez les valeurs de l'IMU

void setup() {
  Serial.begin(115200);
}

void loop() {
  // Lire IMU (BMI088, MPU6050, etc.)
  float x = read_acc_x();
  float y = read_acc_y();
  float z = read_acc_z();
  
  // Envoyer au format CSV
  Serial.print(x);
  Serial.print(",");
  Serial.print(y);
  Serial.print(",");
  Serial.println(z);
  
  delay(50);  // 20 Hz
}
```

#### 3.2 Code de transmission (STM32 + USB CDC)

```c
// main.c
#include "usbd_cdc_if.h"

void main_loop() {
  float acc_x, acc_y, acc_z;
  BMI088_ReadAcc(&bmi088, &acc_x, &acc_y, &acc_z);
  
  char buf[64];
  snprintf(buf, sizeof(buf), "%.4f,%.4f,%.4f\n", acc_x, acc_y, acc_z);
  CDC_Transmit_FS((uint8_t *)buf, strlen(buf));
  
  osDelay(50);  // 20 Hz
}
```

## 🎮 Contrôles

| Action | Contrôle |
|--------|----------|
| Rotation | Drag souris (clic gauche) |
| Zoom | Scroll souris |
| Reset view | R |
| Fermer | Alt+F4 ou close window |

## 📊 Exemple d'utilisation complète

### Étape 1 : Lancer le simulateur

```bash
$ python com_simulator.py --port COM3 --mode spiral
✓ Connected to COM3 @ 115200 baud
✓ Sending vectors in mode: spiral

Vector data:
------------------------------------------------------------
[00010] X=  50.00 Y=   0.00 Z=  -2.73 Mag=  50.07
[00020] X=  49.24 Y=  10.37 Z=  -5.41 Mag=  50.31
[00030] X=  47.00 Y=  20.01 Z=  -7.98 Mag=  51.00
```

### Étape 2 : Lancer le visualiseur

```bash
$ python vector_visualizer.py --port COM3
==================================================
Vector Orientation Visualizer
==================================================
Port: COM3
Baudrate: 115200

Expected data formats:
  • CSV:    x,y,z
  • Space:  x y z
  • JSON:   {"x": 1.0, "y": 2.0, "z": 3.0}

Controls:
  • Drag mouse to rotate view
  • Scroll to zoom
  • Close window to exit
==================================================

✓ Connected to COM3 @ 115200 baud
```

### Étape 3 : Voir la visualisation

Une fenêtre matplotlib s'ouvre avec :
- **Gauche** : Vue 3D interactive
  - Vecteur en violet
  - Sphère magnitude (surface semi-transparente)
  - Trail cyan (historique)
  - Axes de référence (X rouge, Y vert, Z bleu)

- **Droite** : Panneau statistiques
  - Coordonnées X, Y, Z
  - Magnitude
  - Angles (azimut, élévation)
  - Nombre de paquets reçus
  - Dernière mise à jour

## 🧪 Modes de simulateur

### rotating (défaut)
Vecteur qui tourne dans le plan XY

```bash
python com_simulator.py --port COM3 --mode rotating
```

### spiral
Mouvement en spirale 3D

```bash
python com_simulator.py --port COM3 --mode spiral
```

### pendulum
Oscillation comme un pendule

```bash
python com_simulator.py --port COM3 --mode pendulum
```

### random
Mouvements aléatoires

```bash
python com_simulator.py --port COM3 --mode random
```

## 📈 Exemples de sorties visualisées

### 1. Accéléromètre IMU (rotation statique)
```
Vecteur: [0.2, -0.3, 9.81]
Magnitude: 9.81 m/s²
Azimut: -56.3°
Élévation: 80.1°
```
→ Affiche l'inclinaison et l'orientation du capteur

### 2. Gyroscope (rotation dynamique)
```
Vecteur: [45.2, -23.1, 12.5]
Magnitude: 51.3 °/s
Azimut: -27.2°
Élévation: 14.0°
```
→ Montre la vitesse angulaire autour de chaque axe

### 3. Champ magnétique (boussole)
```
Vecteur: [15.3, 8.2, -5.1]
Magnitude: 18.1 µT
Azimut: 28.4°
Élévation: -16.2°
```
→ Visualise le Nord magnétique

## 🔧 Options avancées

```bash
# Changer la taille de l'historique (par défaut 100)
python vector_visualizer.py --port COM3 --history 200

# Augmenter le baudrate
python vector_visualizer.py --port COM3 --baudrate 230400

# Lister tous les ports disponibles
python vector_visualizer.py --list-ports
```

## 🐛 Dépannage

### Erreur : "Port not found"
```
✗ Serial error: [Errno 2] could not open port COM3: ...
```

**Solutions:**
1. Vérifier le port : `python vector_visualizer.py --list-ports`
2. Utiliser le bon port : `--port COM4`
3. Redémarrer le matériel (reset board)

### Erreur : "Could not parse"
```
⚠ Parse error: Could not parse: invalid_data
```

**Solutions:**
1. Vérifier le format des données envoyées
2. Format supportés : `10,20,30` ou `10 20 30` ou `{"x":10,"y":20,"z":30}`
3. S'assurer que le port COM envoie des données toutes les <1s

### Visualisation lente / freeze

**Solutions:**
1. Réduire l'historique : `--history 50`
2. Réduire la fréquence d'envoi (augmenter `delay`)
3. Fermer d'autres applications

### Port COM non détecté sur Windows

**Solutions:**
1. Installer les drivers USB (ST-Link, CH340, etc.)
2. Vérifier le Device Manager
3. Redémarrer l'IDE Arduino/STM32CubeIDE

## 📚 Format de données détaillé

### CSV (recommandé)
```
x,y,z
1.5,2.3,-0.8
```
**Avantages:** Simple, parsable facilement, compact

### JSON
```json
{"x": 1.5, "y": 2.3, "z": -0.8}
```
**Avantages:** Extensible, peut ajouter des métadonnées

### Space-separated
```
1.5 2.3 -0.8
```
**Avantages:** Lisible, pas de délimiteurs spéciaux

### Tab-separated
```
1.5	2.3	-0.8
```
**Avantages:** Compatible Excel, lisible

## 🎓 Cas d'usage

1. **Débugger une IMU** → Voir l'orientation en temps réel
2. **Tester une fusion de capteurs** → Vérifier la fusion est stable
3. **Valider une calibration** → Voir si le vecteur tourne correctement
4. **Enseigner la 3D** → Visualiser des concepts d'orientation
5. **Démonstration robotique** → Afficher l'inclinaison en live

## 📝 Exemple complet : STM32 + BMI088

```c
// Dans votre projet STM32CubeIDE avec BMI088

#include "usbd_cdc_if.h"
#include "BMI088.h"

void task_telemetry(void *argument) {
    float3_t acc;
    
    while(1) {
        // Lire l'accéléromètre
        BMI088_ReadAcc(&bmi088, &acc);
        
        // Envoyer au format CSV
        char buf[64];
        snprintf(buf, sizeof(buf), "%.3f,%.3f,%.3f\r\n", 
                 acc.x, acc.y, acc.z);
        
        CDC_Transmit_FS((uint8_t *)buf, strlen(buf));
        
        osDelay(50);  // 20 Hz
    }
}
```

Puis côté Python :
```bash
python vector_visualizer.py --port COM3
```

## 📞 Support

Si vous rencontrez des problèmes :

1. Vérifier les dépendances : `pip list | grep -E "pyserial|matplotlib|numpy"`
2. Mettre à jour : `pip install --upgrade pyserial matplotlib numpy`
3. Tester avec le simulateur d'abord : `python com_simulator.py`

## 📄 Licence

Code libre d'utilisation. Merci de mentionner la source si vous l'adaptez ! 

---

**Auteur :** APEX Team  
**Date :** Mars 2026  
**Version :** 1.0
