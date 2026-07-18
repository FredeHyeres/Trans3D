VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlacerPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CPlacerPoints - Machine a etats : point 3D depuis une position 2D + une cote
'
' Etape 1 (POSITION) : snap de la position XY exacte du point
'   Data  = memoriser le XY snappe (sommet, intersection...)
'   Reset = quitter la commande
' Etape 2 (ALTITUDE) : clic pres de la cote d'altitude (texte/cellule/tag)
'   Data  = lire la cote, creer le point cote (cercle au Z reel + texte a
'           Z=0.00), retour etape 1
'   Reset = abandonner cette position, retour etape 1
'
' REGLE : le Z du point snappe n'est JAMAIS utilise - seule la valeur de la
' cote fournit l'altitude (le snap renvoie le Z de l'element vise).
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private m_nEtape    As Integer      ' 1..2
Private m_dPosX     As Double       ' position XY memorisee (repere actif)
Private m_dPosY     As Double

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Trans3D - Points"
    RetourPosition
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Select Case m_nEtape
    Case 1
        ' --- Memoriser la position (le Z snappe est ignore) ---
        m_dPosX = Point.X
        m_dPosY = Point.Y
        ShowPrompt "Position fixee  |  Cliquez la cote d'altitude" & _
                   "  (Reset=autre position)"
        m_nEtape = 2

    Case 2
        ' --- Lire la cote et creer le point ---
        Dim oText As TextElement
        If Not TrouverAltitudeProche(Point, g_oSettings.dTolTexte, oText) Then
            ShowPrompt "Aucune cote a moins de " & g_oSettings.dTolTexte & _
                       " u.m. - recliquez"
            Exit Sub
        End If

        Dim oCote As New CPointRef
        Dim sSep As String, nDec As Integer
        If Not g_oCalc.LireAltitudeTexte(oCote, sSep, nDec) Then
            ShowPrompt "Texte '" & g_oSelectionCourante.ValeurTexte & _
                       "' non numerique - recliquez"
            Exit Sub
        End If

        ' La cote lue sert de modele de style pour le texte cree
        If Not oText Is Nothing Then
            g_oSettings.ChargerModeleTexte oText, sSep, nDec
        ElseIf Not g_oSelectionCourante.oTag Is Nothing Then
            g_oSettings.ChargerModeleDepuisTag g_oSelectionCourante.oTag, _
                                               sSep, nDec
        Else
            g_oSettings.oTexte.SepDecimal = sSep
            g_oSettings.AjusterDecimales nDec
        End If

        g_oMoteur.CreerPointCote m_dPosX, m_dPosY, oCote.Altitude, _
                                 g_oSettings, g_oCalc

        ShowPrompt "Point cree Z=" & _
                   g_oCalc.FormatAltitude(oCote.Altitude, _
                       g_oSettings.oTexte.Decimales, _
                       g_oSettings.oTexte.SepDecimal) & _
                   "  |  Snappez la position suivante  (Reset=quitter)"
        RetourPosition
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    If m_nEtape <> 2 Then Exit Sub

    ' Croix sur la position memorisee (a plat : l'altitude n'est pas connue)
    Dim ptPos As Point3d
    ptPos.X = m_dPosX
    ptPos.Y = m_dPosY
    ptPos.Z = 0#
    g_oMoteur.DessinCroix DrawMode, ptPos, g_oSettings.oCercle.Rayon
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    Select Case m_nEtape
    Case 1
        ' TODO formulaire : frmTrans3D.Hide
        ShowPrompt "Trans3D - Points : termine."
        CommandState.StartDefaultCommand

    Case 2
        RetourPosition
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

'------------------------------------------------------------------------------
Private Sub RetourPosition()
    m_nEtape = 1
    ShowPrompt "Snappez la position du point 3D  (Reset=quitter)"
End Sub
