# Trans3D

Macro VBA pour **MicroStation V8i** : passage du plan 2D coté à la modélisation 3D terrain.

L'utilisateur travaille dans un fichier 3D avec des références 2D (plans topo/VRD cotés). Trans3D lit les altitudes sur le plan 2D et crée les éléments 3D correspondants dans le modèle actif.

## Commandes

| Key-in | Description |
|---|---|
| `vba run [Trans3D]Convertir` | Convertit un élément 2D (ligne, polyligne, arc, courbe, B-spline, cercle) en élément 3D avec altitudes interpolées sur l'abscisse curviligne |
| `vba run [Trans3D]Semer` | Sème des points cotés (cercle repère + texte altitude) le long d'un élément 3D existant, à pas fixe ou N points réguliers |
| `vba run [Trans3D]Points` | Crée un point 3D isolé depuis une position 2D et une altitude cliquée ou saisie |

Les trois commandes partagent un formulaire modeless unique avec les réglages de présentation (texte, cercle repère, format altitude, discrétisation).

## Installation

```
install.cmd
```

Le script copie `Trans3d1.mvba` dans le dossier VBA de MicroStation.

## Architecture

```
Trans3D.bas              Point d'entrée, globals
CSettings / CSymbo*      Paramètres et symbologie
CPointRef                Données géométriques (Double, sans API MST)
CCalcul                  Calcul pur (interpolation curviligne), sans API MST
CGraphique               Création d'éléments et aperçu dynamique (API MST)
CPlacer* / CSelect*      Machines à états (IPrimitiveCommandEvents)
RechercheAltitude.bas    Scan des textes/cellules d'altitude (modèle + réfs)
RechercheElement.bas     Scan des éléments linéaires dans les références
frmTrans3D.frm           Formulaire modeless, contrôles créés au runtime
```

## Prérequis

- MicroStation V8i (SELECTseries)
- VBA 6.x intégré à MicroStation
