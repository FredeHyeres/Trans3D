VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSymboTexte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSymboTexte - Attributs de symbologie du texte altitude
'
' Deux modes, communs a toutes les commandes Trans3D (reglage global) :
'   CommeModele = True (defaut, mode duplication) : le texte cree reprend
'     tous les attributs du texte source lu sur la reference (via
'     CreateTextElement1 avec modele) ; NomNiveau/Couleur refletent alors le
'     modele, a titre informatif.
'   CommeModele = False (mode creation) : NomPolice/Hauteur/Largeur/Couleur/
'     NomNiveau/FondActif sont choisis par l'utilisateur dans le formulaire
'     et appliques a tous les textes crees, quelle que soit la reference.
'
' Stocke egalement la convention de formatage numerique (separateur,
' decimales) : defauts "." et 2, ajustes sur le texte source en duplication.
'==============================================================================
Option Explicit

Public CommeModele  As Boolean  ' True = memes attributs que le texte source
Public NomNiveau    As String   ' niveau du texte cree (nom ou numero)
Public Couleur      As Long     ' index couleur MicroStation 0-255
Public NomPolice    As String   ' mode creation : police ("" = police active)
Public Hauteur      As Double   ' mode creation : hauteur du texte (u.m.)
Public Largeur      As Double   ' mode creation : largeur (0 = comme hauteur)
Public FondActif    As Boolean  ' mode creation : fond/masque sous le texte

Public SepDecimal   As String   ' "." ou "," : separateur du texte cree
Public Decimales    As Integer  ' nombre de decimales du texte cree

'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    InitDefauts
End Sub

'------------------------------------------------------------------------------
Sub InitDefauts()
    CommeModele = True
    NomNiveau = ""
    Couleur = 0
    NomPolice = ""
    Hauteur = 0.1
    Largeur = 0#
    FondActif = False
    SepDecimal = "."
    Decimales = 2
End Sub

'------------------------------------------------------------------------------
' Charge NomNiveau et Couleur depuis le TextElement source - uniquement en
' mode duplication, pour ne pas ecraser un choix utilisateur.
' SepDecimal et Decimales sont renseignes par CCalcul lors du parsing.
Sub ChargerDepuisElement(oText As TextElement)
    If Not CommeModele Then Exit Sub
    On Error Resume Next
    NomNiveau = oText.Level.Name
    Couleur = oText.Color
    On Error GoTo 0
End Sub
