VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSymboElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSymboElements - Symbologie des elements 3D crees par Convertir (ligne
' brisee, arc, ellipse)
'
' bSymboSource (CSettings) est prioritaire : True = dupliquer la symbologie
' de l'element 2D source. Sinon, deux sous-modes :
'   bActifs = True (defaut) : attributs actifs de MicroStation (comportement
'     historique, CreateXxxElement sans modele les applique deja).
'   bActifs = False : NomNiveau/Couleur/NomStyleLigne/Epaisseur choisis par
'     l'utilisateur dans le formulaire sont appliques a tous les elements
'     crees.
'==============================================================================
Option Explicit

Public bActifs       As Boolean  ' True = attributs actifs (pas de choix explicite)
Public NomNiveau     As String   ' niveau des elements crees (nom ou numero)
Public Couleur       As Long     ' index couleur MicroStation 0-255
Public NomStyleLigne As String   ' nom du style de ligne ("" = actif)
Public Epaisseur     As Long     ' epaisseur de trait 0-31

'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    InitDefauts
End Sub

'------------------------------------------------------------------------------
Sub InitDefauts()
    bActifs = True
    NomNiveau = ""
    Couleur = 0
    NomStyleLigne = ""
    Epaisseur = 0
End Sub
