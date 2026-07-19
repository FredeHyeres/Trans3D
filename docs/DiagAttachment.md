# DiagAttachment - Diagnostic des references

Module utilitaire pour identifier les proprietes disponibles sur les objets
`Attachment` en VBA MicroStation V8i.

## Lancement

```
vba run [Trans3D]DiagReferences
```

## Resultat

- **MsgBox** : affiche le rapport directement
- **Fichier** : `DiagAttachment.txt` cree a cote du DGN actif

## Proprietes testees

| Propriete        | Type    | Utilite                                      |
|------------------|---------|----------------------------------------------|
| `DisplayFlag`    | Boolean | **Flag d'affichage de la reference** (celle utilisee par `AttachmentAffiche`) |
| `IsAttachment`   | Boolean | Toujours True pour un attachement             |
| `DisplayPriority`| Long    | Priorite d'affichage                          |
| `IsActive`       | Boolean | Reference active (chargee)                    |
| `DisplayAsNested`| Boolean | Affichage des sous-references imbriquees      |

## Proprietes V8i inexistantes (err 438)

Les proprietes suivantes, souvent documentees pour des versions plus recentes,
**n'existent pas** sur `Attachment` en V8i SS3 :

- `Display`
- `IsDisplayed`
- `IsDisplayedInView(View)`
- `IsEffectivelyDisplayedInView(View)`
- `IsEffectivelyDisplayed`

## Contexte

Cree lors du fix du bug "references desactivees toujours selectionnables"
(juillet 2026). Le code utilisait `IsDisplayed` qui n'existe pas en V8i ;
l'erreur etait avalee par `On Error Resume Next` et la fonction retournait
toujours `True`. Le fix utilise `DisplayFlag` a la place.
