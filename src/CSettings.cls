VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSettings - Parametres partages des trois commandes Trans3D
'
' Agregateur unique des parametres, regroupes par domaine :
'   oTexte   : symbologie et format du texte altitude (duplication / creation)
'   oCercle  : symbologie du cercle repere
'   oSemis   : repartition des points semes (Semer)
'
' Regles fixes (cahier des charges, spec section 1) :
'   - le cercle repere est cree au Z reel calcule (c'est le point exploitable
'     pour la modelisation)
'   - le texte altitude est cree a Z = 0.00 (lecture propre en vue top, pas de
'     surcharge du modele 3D)
'   - le Z d'un point snappe n'est jamais utilise comme altitude
'
' Instancie une seule fois dans Trans3D.bas (g_oSettings), passe en parametre
' aux classes qui en ont besoin (CGraphique, etats de commande, formulaire).
'==============================================================================
Option Explicit

Public oTexte           As CSymboTexte      ' symbologie et format du texte
Public oCercle          As CSymboCercle     ' symbologie du cercle repere
Public oSemis           As CParamSemis      ' repartition des points semes

Public bCreerTexte      As Boolean          ' Semer/Points : creer le texte altitude
Public bCreerCercle     As Boolean          ' Semer/Points : creer le cercle repere

Public bSymboSource     As Boolean          ' Convertir : True = dupliquer la
                                            ' symbologie de l'element 2D source,
                                            ' False = attributs actifs
Public dPasDiscretisation As Double         ' Convertir : pas de discretisation
                                            ' des arcs/courbes/B-splines pentus (u.m.)

Public dTolTexte        As Double           ' rayon de recherche des cotes autour du clic (u.m.)

Public oTextModele      As TextElement      ' texte source : modele pour CreateTextElement1
Public oTagModele       As TagElement       ' tag source : modele quand la cote est un tag
Public dModeleHauteur   As Double           ' hauteur du texte source (0 = non defini)

'------------------------------------------------------------------------------
Sub Init()
    Set oTexte = New CSymboTexte
    oTexte.InitDefauts
    Set oCercle = New CSymboCercle
    oCercle.InitDefauts
    Set oSemis = New CParamSemis
    oSemis.InitDefauts

    bCreerTexte = True
    bCreerCercle = True
    bSymboSource = True
    dPasDiscretisation = 0.5
    dTolTexte = 1#

    Set oTextModele = Nothing
    Set oTagModele = Nothing
    dModeleHauteur = 0#
End Sub

'------------------------------------------------------------------------------
' Appele apres la lecture d'une cote texte : memorise le texte-modele et ses
' attributs pour les creations en mode duplication.
' PIEGE 5.3 : la hauteur est lue sur un element cree AVEC modele, jamais sur
' TextStyle.Height du texte de la reference (unites potentiellement
' differentes : texte 100 a 1000 fois trop grand sinon).
Sub ChargerModeleTexte(oText As TextElement, sSep As String, nDec As Integer)
    Set oTextModele = oText
    Set oTagModele = Nothing
    oTexte.ChargerDepuisElement oText
    oTexte.SepDecimal = sSep
    AjusterDecimales nDec

    On Error Resume Next
    Dim oTemp As TextElement
    Dim ptO As Point3d
    Set oTemp = CreateTextElement1(oText, "0", ptO, Matrix3dIdentity())
    If Not oTemp Is Nothing Then
        dModeleHauteur = CDbl(oTemp.TextStyle.Height)
    Else
        Err.Clear
        dModeleHauteur = CDbl(oText.TextStyle.Height)
    End If
    If Err.Number <> 0 Then dModeleHauteur = 0#
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Appele quand la cote lue est un tag : estime la hauteur depuis le Range.
Sub ChargerModeleDepuisTag(oTag As TagElement, sSep As String, nDec As Integer)
    Set oTagModele = oTag
    oTexte.SepDecimal = sSep
    AjusterDecimales nDec
    dModeleHauteur = 0#

    On Error Resume Next
    Dim oRng As Range3d
    oRng = oTag.Range
    Dim ptL As Point3d, ptH As Point3d
    ptL = oRng.Low: ptH = oRng.High
    Dim dHY As Double, dHX As Double
    dHY = Abs(ptH.Y - ptL.Y)
    dHX = Abs(ptH.X - ptL.X)
    If dHY > 0 And dHX > 0 Then
        If dHX < dHY Then dModeleHauteur = dHX Else dModeleHauteur = dHY
    End If
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' En mode duplication, conserve le plus grand nombre de decimales rencontre
' sur les cotes sources. En mode creation, les decimales du formulaire priment.
Sub AjusterDecimales(nDec As Integer)
    If Not oTexte.CommeModele Then Exit Sub
    If nDec > oTexte.Decimales Then oTexte.Decimales = nDec
End Sub

'------------------------------------------------------------------------------
' Vrai si un texte-modele est disponible (une cote texte a ete lue)
Property Get TextModeleDisponible() As Boolean
    TextModeleDisponible = Not (oTextModele Is Nothing)
End Property
