# Trans3D - MicroStation V8i VBA

Projet demarre depuis la spec `docs/Spec_Nouveau_Projet_VBA.md` (a lire en
premier : architecture un-metier-une-classe, catalogue des pieges, workflow).
Le cahier des charges (section 1 de la spec) est a remplir avant de coder.

Les classes de `src/` proviennent du projet InterpolationTopo
(`..\Interpolation`) et sont reutilisables telles quelles ou a elaguer.

## Encodage obligatoire des fichiers VBA

Apres chaque creation ou modification de fichier `.cls`, `.bas` ou `.frm`,
normaliser en **CRLF + ANSI (Windows-1252)** avec :

```powershell
$t = [IO.File]::ReadAllText($path)
$t = $t.Replace("`r`n", "`n").Replace("`n", "`r`n")
[IO.File]::WriteAllText($path, $t, [Text.Encoding]::GetEncoding(1252))
```

Raison : le Write tool produit du LF/UTF-8. MicroStation importe les `.cls`
en LF comme des modules standard au lieu de classes, ce qui casse `Implements`
et `New`.

## Formulaire (FRM/FRX)

- Tout formulaire est construit entierement au runtime dans
  `ConstruireControles` : le `.frx` est un blob binaire du formulaire vide.
- Si seul le **code** du `.frm` change : le `.frx` existant reste valide.
- Si les **proprietes du designer** changent (en-tete `Begin...End`) :
  regenerer le couple via `scripts\export_frm_via_excel.ps1`.

## Lancement

Key-in MicroStation : `vba run [Trans3D]<NomCommande>` (a definir avec le
cahier des charges).
