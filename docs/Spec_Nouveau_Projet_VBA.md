# Spec de demarrage - Nouvelle fonctionnalite MicroStation V8i VBA

*Document de kickoff a copier dans le depot du nouveau projet. Il capitalise
l'architecture et les pieges rencontres sur InterpolationTopo (2026).
References detaillees : `docs/Guide_Dev_MicroStation_VBA.md` et
`docs/Guide_UserForm_FRM_FRX.md` du depot InterpolationTopo.*

---

## 1. Cahier des charges (valide le 2026-07-18)

### 1.0 Vision

L'utilisateur travaille dans un **fichier MicroStation 3D actif** avec un ou
plusieurs **fichiers 2D en reference** (plans topo/VRD cotes). Trans3D
fournit les outils pour passer du plan 2D a la modelisation 3D :
conversion d'elements 2D en elements 3D, semis d'altitudes le long des
elements 3D, creation de points 3D isoles. La presentation (textes,
cercles) est **uniforme quelle que soit la reference source**, definie au
depart dans le formulaire et modifiable a tout moment.

Trois commandes partagent le meme formulaire modeless, `CSettings`,
`CCalcul` et `CGraphique`. L'architecture doit permettre d'ajouter des
commandes au fur et a mesure (chaines complexes, contenu de cellules,
maillage... envisages plus tard).

### 1.1 Commande Convertir (element 2D -> element 3D)

| Rubrique | Contenu |
|---|---|
| Key-in | `vba run [Trans3D]Convertir` |
| Objectif metier | Creer dans le modele actif 3D l'equivalent 3D d'un element 2D d'une reference, avec altitudes interpolees |
| Types acceptes | Ligne, ligne brisee, arc, courbe, B-spline, cercle, ellipse |
| Acquisition des altitudes | Clic pres d'un texte/cellule/tag d'altitude du plan 2D (lecture auto type `RechercheAltitude`), repli saisie/correction au formulaire |
| Altitudes | Depart + fin + intermediaires optionnelles (0..N, ex. regards). Cercle/ellipse complet : Z constant, une seule altitude |
| Interpolation | Lineaire sur **abscisse curviligne** (longueur developpee, pas la corde) ; par troncons si altitudes intermediaires |
| Element cree | **Type conserve si possible** : ligne -> ligne 3D, ligne brisee -> ligne brisee 3D (Z aux sommets), arc/cercle/ellipse a Z constant -> type natif horizontal. Arc/courbe/B-spline avec pente -> polyligne 3D discretisee au pas parametrable |
| Symbologie | Duplication de l'element source par defaut (niveau cree dans le fichier actif s'il n'existe pas) ; bascule formulaire "attributs actifs" |
| Enchainement | 1) selection element 2D -> 2) clic Z depart -> 3) clic Z fin -> 4) Data = ajouter une altitude intermediaire, Reset = creer l'element -> retour a 1 |
| Comportement Reset | Etape 4 : valider/creer. Etapes 2-3 : remonter d'un cran. Etape 1 : quitter la commande |
| Apercu dynamique | Complet : surlignage de l'element 2D selectionne, apercu de l'element 3D pendant la saisie, pente courante affichee au formulaire |

### 1.2 Commande Semer (altitudes le long d'un element 3D)

| Rubrique | Contenu |
|---|---|
| Key-in | `vba run [Trans3D]Semer` |
| Objectif metier | Materialiser des points cotes le long d'un element lineaire 3D |
| Cible | **Tout element lineaire du modele actif** (issu de Trans3D ou non) ; Z interpole sur la geometrie 3D reelle de l'element |
| Modes de repartition | a) distance fixe (ex. 0,50 m) : point au depart, puis tous les pas, point a la fin meme si dernier intervalle plus court ; b) N points a parts egales (longueur / N) |
| Objets crees par point | **Cercle repere au Z reel** (le point exploitable pour la modelisation) + **texte d'altitude a Z=0.00** (lisible en vue top sans surcharger le modele). Chacun activable au formulaire |
| Apercu dynamique | Apercu des points semes suivant les parametres |

### 1.3 Commande Points (point 3D depuis le plan 2D)

| Rubrique | Contenu |
|---|---|
| Key-in | `vba run [Trans3D]Points` |
| Objectif metier | Creer un point 3D a partir d'une position 2D et d'une altitude du plan |
| Geste | Deux temps : 1) clic/snap = position XY exacte (sommet, intersection...) -> 2) clic sur la cote (texte/cellule/tag) ou saisie du Z |
| Objets crees | Identiques a Semer : cercle repere au Z reel + texte d'altitude a Z=0.00 |
| Enchainement | Repetitif : position -> altitude -> creation -> position... Reset : remonter d'un cran / quitter a la premiere etape |

### 1.4 Reglages communs (formulaire modeless)

| Rubrique | Contenu |
|---|---|
| Presentation textes | Mode **duplication** (reprendre le style du texte source de la reference) ou mode **creation** (police, taille, couleur, masque/background, niveau definis). Applique uniformement a toutes les creations, quel que soit le fichier de reference |
| Format altitude | Decimales parametrables (defaut 2), separateur point ou virgule (defaut point). La lecture des cotes 2D accepte les deux separateurs |
| Cercle repere | Diametre et symbologie parametrables (`CSymboCercle`) |
| Discretisation | Pas de discretisation des courbes pentues (Convertir) |
| Sources d'altitude | Modele actif **et** references attachees (textes, cellules, tags) |
| Regle Z | Le Z des elements 3D et des cercles = altitude calculee ; le Z des textes = 0.00 (voulu). Jamais le Z du point snappe |

---

## 2. Architecture imposee : un metier = une classe

Decomposition validee sur InterpolationTopo. Le module `.bas` ne contient
que le point d'entree, les globals et l'init ; toute la logique vit dans
des classes.

```
MaCommande()                        ' .bas : entree + globals
      |
      +--- CSettings (g_oSettings)  ' agregateur de parametres
      |         +-- CSymboXxx       '   sacs de proprietes (InitDefauts)
      |         +-- CParamXxx
      |
      +--- CPointRef x N            ' donnees geo (Double, pas d'UDT public)
      +--- CCalcul (g_oCalc)        ' calcul pur, zero API MST, testable seul
      +--- CGraphique (g_oMoteur)   ' creation + apercu dynamique (API MST)
      +--- RechercheXxx.bas         ' scan modele + references si besoin
      |
      +--- frmMaCommande            ' modeless, controles crees au runtime
      |
      +--- CommandState
               +-- CSelect1 -> CSelect2 -> CPlacer (machine a etats m_nEtape)
```

| Domaine | Classe(s) | Connait l'API MST ? |
|---|---|---|
| Parametres | `CSettings` + sacs de proprietes | Non |
| Donnees geo | `CPointRef` | Non (composantes en `Double`) |
| Calcul | `CCalcul` | Non |
| Graphique | `CGraphique` | Oui |
| Recherche/selection | module `.bas` + classe resultat | Oui (references) |
| Etats de commande | `CSelectXxx`, `CPlacerXxx` | Oui (`IPrimitiveCommandEvents`) |
| UI | `frmXxx` | Oui (lecture niveaux DGN) |

**Regles qui ont paye :**

- Une classe de placement complexe = **une machine a etats** avec un seul
  membre `m_nEtape` et un `Select Case` dans `Dynamics` / `DataPoint` /
  `Reset`. Documenter les etapes dans l'en-tete de la classe
  (cf. `CPlacerPonctuel.cls` : 6 etapes commentees).
- Les routines graphiques prennent des **parametres generiques**
  (`CPointRef`, `CSettings`), jamais l'etat interne d'une commande.
  C'est ce qui a permis de reutiliser l'indicateur pente + fleche dans
  4 contextes differents sans nouvelle classe.
- Les etats de selection partages entre plusieurs commandes recoivent une
  propriete `Mode` (enum) plutot que d'etre dupliques.

---

## 3. Squelette de fichiers

```
MonProjet/
  src/
    MonProjetV1.bas          ' entree + globals + ShowCommand/ShowPrompt
    CSettings.cls
    CSymbo*.cls              ' un sac de proprietes par famille d'attributs
    CPointRef.cls            ' copier depuis InterpolationTopo
    CCalcul.cls
    CGraphique.cls
    CSelect*.cls / CPlacer*.cls
    frmMonProjet.frm + .frx
  scripts/
    export_frm_via_excel.ps1 ' copier depuis InterpolationTopo
  docs/
  install.cmd / install.ps1  ' copier + adapter depuis InterpolationTopo
  .gitattributes             ' OBLIGATOIRE des le premier commit (cf. 6.1)
  CLAUDE.md                  ' cf. section 8
  MonProjet.mvba             ' binaire regenere depuis MicroStation
```

### Classes reutilisables telles quelles depuis InterpolationTopo

| Fichier | Role | Adaptation |
|---|---|---|
| `CPointRef.cls` | Point topo (X/Y/Z/Altitude/Valide) | Aucune |
| `CSymboTexte.cls`, `CSymboCercle.cls`, `CSymboPente.cls` | Sacs de proprietes symbologie | Renommer/elaguer |
| `CParamPenteDZ.cls` | Couple pente % + DZ activables | Aucune |
| `CInterpolation.cls` | Projection, interpolation, pente, gisement, formatage | Garder les fonctions utiles |
| `CMoteurGraphique.cls` | Creation texte/cercle, apercu, indicateur pente + fleche | Extraire les routines utiles |
| `RechercheAltitude.bas` | Scan texte/cellule/tag, modele + references, coordonnees converties | Adapter les criteres |
| `CAltitudeSelection.cls` | Resultat de recherche (source, origine, reference) | Aucune |

---

## 4. Patterns obligatoires (resume)

Details et code complet dans le Guide_Dev. L'essentiel :

- **Entree `.bas`** : verifier `ActiveDesignFile`, instancier les globals,
  `frmXxx.Initialiser g_oSettings`, `Show vbModeless`, puis
  `CommandState.StartPrimitive New CSelect1`.
- **`CommandState.StartDynamics`** dans le `Start` de chaque etat qui
  dessine — sinon `Dynamics` ne se declenche jamais.
- **Apercu dynamique** : creer les elements avec parent `Nothing` puis
  `.Redraw DrawMode`. Entourer de `On Error Resume Next` quand l'element
  modele peut etre invalide.
- **Formulaire** : 100 % des controles crees dans `ConstruireControles`
  (gardes `m_bConstruit` et `m_bInit`), le `.frx` reste un blob vide.
  Validation continue sur `Change`, reformatage sur `KeyDown` + Entree
  (les evenements extender n'existent pas au runtime).
  `QueryClose` : `Cancel = 1` + `Me.Hide` (jamais decharger un modeless).
- **Jamais d'`InputBox`** : tout parametre passe par le formulaire modeless,
  modifiable en cours de commande.
- **Prompts** : chaque etape affiche quoi faire et ce que font Data/Reset
  (`ShowPrompt "... (Data=placer  Reset=sauter)"`).

---

## 5. Catalogue des pieges rencontres (retour d'experience)

Chaque entree a coute du temps de debug sur InterpolationTopo. A relire
avant de coder.

### 5.1 Encodage — le piege n°1

- L'editeur VBA exige **CRLF + ANSI (Windows-1252), sans BOM**.
- En LF, les `.cls` sont importes comme **modules standard** : `Implements`
  et `New` cassent, la compilation echoue de facon obscure.
- Les outils d'edition modernes (dont le Write tool de Claude) produisent
  du LF/UTF-8 : **normaliser apres chaque modification** (script en 6.2).
- Pas d'accents dans le code ni les commentaires des sources VBA.

### 5.2 Formulaires .frm/.frx

- Un `.frm` ne s'ecrit pas a la main (GUID designer + `OleObjectBlob`).
  Generer le couple via `export_frm_via_excel.ps1`, renommer, corriger
  la ligne `OleObjectBlob`.
- Controles au runtime => le `.frx` ne change **jamais** : si seul le code
  du `.frm` evolue (nouveaux controles dans `ConstruireControles`,
  handlers), inutile de regenerer le couple.
- `.frm` et `.frx` doivent etre dans le meme dossier a l'import.

### 5.3 Unites et elements modeles

- **Ne jamais** creer un texte avec `CreateTextElement1(Nothing, ...)` puis
  appliquer la hauteur lue dans `TextStyle.Height` : avec des references
  aux unites differentes, le texte sort **100 a 1000 fois trop grand**.
  Toujours passer l'element source comme modele :
  `CreateTextElement1(oTexteModele, ...)`, avec repli sur `Nothing`
  seulement en dernier recours.

### 5.4 References attachees

- Convertir systematiquement les coordonnees des elements de reference
  dans le **repere du modele actif** avant tout calcul.
- **ScaleFactor n'est PAS une transformation de coordonnees** (verifie sur
  DGN reel : module Reseaux d'InterpolationTopo, puis premier test Trans3D).
  Les coordonnees scannees d'une reference sont deja compatibles master ;
  le ScaleFactor est l'echelle graphique du contenu. L'appliquer aux
  positions rend la detection impossible (tolerance x500 necessaire) et
  cree les elements a la mauvaise echelle. Conversion correcte :
  translation seule `MasterOrigin - ReferenceOrigin`.
- **Iterer les attachements par index** (`For i = 1 To Attachments.Count`) :
  le `For Each` sur Attachments echoue sans erreur visible selon la
  version V8i (retour Reseaux).
- Gerer les references **2D attachees a un fichier 3D** (Z absent).
- Verifier que la reference est affichee avant de la scanner
  (`AttachmentAffiche`), et entourer l'iteration des attachements de
  `On Error Resume Next` (certains attachements sont invalides).
- Les cellules et tags peuvent venir d'une reference : la source n'est
  alors **pas modifiable** — prevoir un mode degrade (creer cercle + texte
  au lieu de dupliquer la cellule).

### 5.5 Altitude / geometrie

- Le **Z des elements crees = altitude calculee**, jamais le Z du point
  snappe ou projete (piege : le snap renvoie le Z de l'element vise).
- Proteger les calculs de pente contre les **points confondus**
  (distance < epsilon ~1 mm) : sauter l'etape plutot que diviser par zero.
- La pente affichee doit refleter la **realite du segment dessine**
  (difference d'altitude reelle), pas seulement le parametre saisi.

### 5.6 Tags et cellules

- Un tag reassocie a un element supprime devient **orphelin** : cloner
  l'element hote via `BaseElement` puis reassocier le tag.
- Scanner `msdElementTypeCellHeader` **et** `msdElementTypeTag` en plus
  des textes ; le contenu numerique d'une cellule se cherche dans ses
  sous-elements.

### 5.7 API V8i — incompatibilites connues

- `CommandState.StartPrimitive` (pas `StartPrimitiveCommand`, absent sur
  certaines installations).
- Pas de membre `Public` de type UDT (`Point3d`, `Matrix3d`) dans une
  classe : stocker en `Double` + `Sub` de conversion
  (`DefinirPosition` / `CopierPosition`).
- Pas de `Function` retournant un UDT : `Sub` avec parametre `ByRef`.
- Le **Tool Settings natif** est inaccessible en VBA (MDL/C++ seulement) :
  le UserForm modeless est le seul equivalent.

### 5.8 Deploiement — dgnlib et mvba

- Livrer **un seul `.mvba`** regenere depuis MicroStation, copie dans
  `Standards\vba` du workspace par `install.cmd`. Ne jamais toucher au
  `Default.mvba` personnel de l'utilisateur.
- La mise a jour fiable = **rejouer `install.cmd`** ; le reimport manuel
  fichier par fichier provoque des erreurs de compilation si un fichier
  est oublie.
- Une `dgnlib` chargee par MicroStation est **verrouillee** : impossible
  de la modifier a chaud. Prevoir le remplacement fichier ferme, et gerer
  les **conflits de dgnlib multiples** (la meme toolbox presente dans
  plusieurs chemins de workspace).
- Attention aux **copies masquantes** dans `Projects\*\vba` qui prennent
  le pas sur la version de `Standards\vba`.
- Les **touches de fonction** (Utilitaires > Touches de fonction) sont le
  moyen d'acces le plus robuste ; la toolbox dgnlib est un plus fragile.

### 5.9 Formulaire modeless

- `m_bInit` (garde pendant `Initialiser`) et `m_bConstruit` (garde contre
  double construction) sont **obligatoires** — sans eux, les evenements
  `Change` ecrasent les settings pendant le pre-remplissage.
- Erreur 91 apres fermeture par la croix = formulaire decharge puis
  recree vierge : `QueryClose` avec `Cancel = 1` + `Me.Hide`.
- Prevoir une methode publique `ReinitialiserEtat` appelee par les classes
  de commande au Reset.

---

## 6. Encodage et git — a mettre en place au premier commit

### 6.1 .gitattributes

```
*.bas text eol=crlf
*.cls text eol=crlf
*.frm text eol=crlf
*.frx binary
*.ps1 text eol=crlf
*.dgn binary
*.dgnlib binary
*.mvba binary
```

### 6.2 Normalisation apres toute edition externe

```powershell
Get-ChildItem src\* -Include *.bas,*.cls,*.frm | ForEach-Object {
    $t = [IO.File]::ReadAllText($_.FullName)
    $t = $t.Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText($_.FullName, $t, [Text.Encoding]::GetEncoding(1252))
}
```

---

## 7. Workflow de developpement valide

1. Editer les sources texte dans `src/` (jamais le `.mvba` directement).
2. Normaliser l'encodage (6.2) apres chaque modification.
3. Importer/tester dans MicroStation (`Ctrl+M`, verifier que les `.cls`
   arrivent bien en **Modules de classe**), `Debogage > Compiler`.
4. Tester au key-in : `vba run [MonProjetV1]MaCommande`.
5. Regenerer `MonProjet.mvba` depuis MicroStation (enregistrement du
   projet VBA), commit sources + mvba ensemble.

---

## 8. CLAUDE.md du nouveau projet (gabarit)

```markdown
# MonProjet - MicroStation V8i VBA

## Encodage obligatoire des fichiers VBA

Apres chaque creation ou modification de fichier `.cls`, `.bas` ou `.frm`,
normaliser en **CRLF + ANSI (Windows-1252)** avec :

    $t = [IO.File]::ReadAllText($path)
    $t = $t.Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText($path, $t, [Text.Encoding]::GetEncoding(1252))

Raison : le Write tool produit du LF/UTF-8. MicroStation importe les `.cls`
en LF comme des modules standard au lieu de classes, ce qui casse
`Implements` et `New`.

## Formulaire (FRM/FRX)

- Le formulaire est construit entierement au runtime dans
  `ConstruireControles` : le `.frx` est un blob binaire du formulaire vide.
- Si seul le **code** du `.frm` change : le `.frx` reste valide.
- Si les **proprietes du designer** changent (en-tete `Begin...End`) :
  regenerer le couple via `scripts\export_frm_via_excel.ps1`.

## Architecture

- Un metier = une classe (cf. docs/Spec_Nouveau_Projet_VBA.md, section 2).
- Calcul et donnees : zero API MicroStation.
- Placements complexes : machine a etats `m_nEtape` documentee en tete
  de classe.

## Lancement

Key-in MicroStation : `vba run [MonProjetV1]MaCommande`
```

---

## 9. Check-list de demarrage

1. [ ] Remplir le cahier des charges (section 1)
2. [ ] Creer le depot avec `.gitattributes` (6.1) et `CLAUDE.md` (8)
3. [ ] Copier depuis InterpolationTopo : `CPointRef.cls`, les `CSymbo*`
       utiles, `export_frm_via_excel.ps1`, `install.cmd`/`install.ps1`
4. [ ] Ecrire le module d'entree + `CSettings` + sacs de proprietes
5. [ ] Ecrire `CCalcul` (pur) puis `CGraphique`
6. [ ] Ecrire les etats de commande, documenter la machine a etats
7. [ ] Formulaire : code runtime, puis generer le couple `.frm`/`.frx`
8. [ ] Normaliser l'encodage, importer, compiler, tester au key-in
9. [ ] Regenerer le `.mvba`, adapter `install.ps1`, tester `install.cmd`
       sur un poste vierge
