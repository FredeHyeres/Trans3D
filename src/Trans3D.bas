Attribute VB_Name = "Trans3D"
'==============================================================================
' Trans3D - Module principal - MicroStation V8i SS3 (VBA)
'
' Passage de plans 2D (en reference) vers la modelisation 3D dans le fichier
' actif. Trois commandes partagent les globals, le formulaire et les moteurs
' (cahier des charges : docs/Spec_Nouveau_Projet_VBA.md, section 1) :
'
'   Convertir : element 2D de reference -> element 3D (altitudes depart/fin
'               + intermediaires, interpolation sur abscisse curviligne)
'   Semer     : semis de points cotes le long d'un element 3D du modele actif
'               (pas fixe ou parts egales)
'   Points    : point 3D depuis une position 2D snappee + une cote du plan
'
' Regle Z commune : cercle repere au Z reel calcule, texte altitude a Z=0.00.
'
' Ce module ne contient que les globals partages, l'initialisation commune et
' les points d'entree. La recherche des cotes est dans RechercheAltitude.bas ;
' toute la logique vit dans les classes.
'
' LANCEMENT :
'   key-in : vba run [Trans3D]Convertir
'   key-in : vba run [Trans3D]Semer
'   key-in : vba run [Trans3D]Points
'==============================================================================
Option Explicit

' --- Mode de repartition du semis (CParamSemis.Mode) ---
Public Enum eModeSemis
    semisDistanceFixe = 0
    semisPartsEgales = 1
End Enum

' --- Instances partagees entre les classes de commande ---
' Chaque classe recoit ces references mais ne les cree pas.
Public g_oSettings  As CSettings         ' parametres (symbologie, semis, tolerance)
Public g_oCalc      As CCalcul           ' moteur de calcul pur
Public g_oSelectionCourante As CAltitudeSelection ' derniere cote trouvee (RechercheAltitude)

'==============================================================================
' Points d'entree des commandes
'==============================================================================

'------------------------------------------------------------------------------
' Convertir un element 2D de reference en element 3D
Sub Convertir()
    If Not EnvironnementPret("Trans3D - Convertir") Then Exit Sub
    InitialiserContexte

    ' TODO (etape suivante) : AfficherFormulaire frmTrans3D
    ' TODO (etape suivante) : CommandState.StartPrimitive New CPlacerConvertir
    ShowCommand "Trans3D - Convertir"
    ShowPrompt "Machine a etats en cours de developpement"
End Sub

'------------------------------------------------------------------------------
' Semer des points cotes le long d'un element 3D du modele actif
Sub Semer()
    If Not EnvironnementPret("Trans3D - Semer") Then Exit Sub
    InitialiserContexte

    ' TODO (etape suivante) : AfficherFormulaire frmTrans3D
    ' TODO (etape suivante) : CommandState.StartPrimitive New CPlacerSemis
    ShowCommand "Trans3D - Semer"
    ShowPrompt "Machine a etats en cours de developpement"
End Sub

'------------------------------------------------------------------------------
' Creer un point 3D depuis une position 2D snappee et une cote du plan
Sub Points()
    If Not EnvironnementPret("Trans3D - Points") Then Exit Sub
    InitialiserContexte

    ' TODO (etape suivante) : AfficherFormulaire frmTrans3D
    ' TODO (etape suivante) : CommandState.StartPrimitive New CPlacerPoints
    ShowCommand "Trans3D - Points"
    ShowPrompt "Machine a etats en cours de developpement"
End Sub

'==============================================================================
' Initialisation commune aux trois commandes
'==============================================================================

'------------------------------------------------------------------------------
' Verifie qu'un fichier DGN est ouvert et que le modele actif est 3D.
' Trans3D travaille toujours dans un fichier 3D avec les plans 2D en reference.
Public Function EnvironnementPret(sTitre As String) As Boolean
    EnvironnementPret = False

    Dim oDgn As DesignFile
    On Error Resume Next
    Set oDgn = ActiveDesignFile
    On Error GoTo 0
    If oDgn Is Nothing Then
        MsgBox "Ouvrez d'abord un fichier DGN.", vbExclamation, sTitre
        Exit Function
    End If

    Dim bEst3D As Boolean
    bEst3D = False
    On Error Resume Next
    bEst3D = ActiveModelReference.Is3D
    On Error GoTo 0
    If Not bEst3D Then
        MsgBox "Trans3D travaille dans un modele 3D : ouvrez le fichier 3D" & _
               vbCrLf & "de travail (les plans 2D restent en reference).", _
               vbExclamation, sTitre
        Exit Function
    End If

    EnvironnementPret = True
End Function

'------------------------------------------------------------------------------
' Instanciation unique des objets partages par les classes de commande.
Public Sub InitialiserContexte()
    Set g_oSettings = New CSettings
    g_oSettings.Init

    Set g_oCalc = New CCalcul
    Set g_oSelectionCourante = New CAltitudeSelection
End Sub

'------------------------------------------------------------------------------
' Initialise, positionne et affiche le formulaire modeless (Tool Settings).
' oFrm en Object : pret pour frmTrans3D (etape suivante).
Public Sub AfficherFormulaire(oFrm As Object)
    oFrm.Initialiser g_oSettings
    oFrm.StartUpPosition = 0
    oFrm.Left = Application.Width * 0.6
    oFrm.Top = Application.Height * 0.05
    oFrm.Show vbModeless
End Sub
