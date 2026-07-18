VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlacerSemis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CPlacerSemis - Machine a etats : semis de points cotes sur un element 3D
'
' Etape 1 (SELECTION) : clic pres de l'element 3D du modele actif
'   Data  = selectionner l'element lineaire le plus proche, extraire sa
'           geometrie 3D reelle
'   Reset = quitter la commande
' Etape 2 (APERCU/CREATION) : apercu du semis complet
'   Dynamics = apercu des points semes selon les parametres courants du
'              formulaire (pas fixe ou parts egales, modifiables en direct)
'   Data  = creer tous les points cotes (cercle au Z reel + texte a Z=0),
'           retour etape 1
'   Reset = abandonner cet element, retour etape 1
'
' Le Z de chaque point est interpole sur la geometrie 3D de l'element.
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private m_nEtape    As Integer      ' 1..2
Private m_oElem     As Element      ' element 3D selectionne
Private m_dX()      As Double       ' polyligne 3D extraite
Private m_dY()      As Double
Private m_dZ()      As Double
Private m_dS()      As Double       ' abscisses curvilignes 3D
Private m_dLong     As Double       ' longueur developpee 3D
Private m_bGeomOk   As Boolean      ' garde : tableaux alloues

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Trans3D - Semer"
    RetourSelection
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Select Case m_nEtape
    Case 1
        SelectionnerElement Point

    Case 2
        ' --- Creer tous les points cotes du semis ---
        Dim dSemis() As Double, nNb As Long
        nNb = CalculerSemis(dSemis)
        If nNb < 1 Then
            ShowPrompt "Parametres de semis invalides - verifiez le formulaire"
            Exit Sub
        End If

        Dim i As Long
        Dim dPX As Double, dPY As Double, dPZ As Double
        For i = 0 To nNb - 1
            g_oCalc.PointSurPolyligne3D m_dX, m_dY, m_dZ, m_dS, dSemis(i), _
                                        dPX, dPY, dPZ
            g_oMoteur.CreerPointCote dPX, dPY, dPZ, g_oSettings, g_oCalc
        Next

        ShowPrompt nNb & " points crees  |  Cliquez le prochain element" & _
                   "  (Reset=quitter)"
        RetourSelection
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    If m_nEtape <> 2 Or Not m_bGeomOk Then Exit Sub

    Dim dSemis() As Double, nNb As Long
    nNb = CalculerSemis(dSemis)
    If nNb < 1 Then Exit Sub

    Dim i As Long
    Dim dPX As Double, dPY As Double, dPZ As Double
    For i = 0 To nNb - 1
        g_oCalc.PointSurPolyligne3D m_dX, m_dY, m_dZ, m_dS, dSemis(i), _
                                    dPX, dPY, dPZ
        g_oMoteur.DessinPointCoteDyn DrawMode, dPX, dPY, dPZ, _
                                     g_oSettings, g_oCalc
    Next

    ShowPrompt "Semis : " & nNb & " points sur " & _
               Format$(m_dLong, "0.00") & " m  (Data=creer  Reset=abandonner)"
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    Select Case m_nEtape
    Case 1
        ' TODO formulaire : frmTrans3D.Hide
        ShowPrompt "Trans3D - Semer : termine."
        CommandState.StartDefaultCommand

    Case 2
        RetourSelection
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

'==============================================================================
' Etapes et utilitaires prives
'==============================================================================

'------------------------------------------------------------------------------
Private Sub RetourSelection()
    m_nEtape = 1
    m_bGeomOk = False
    Set m_oElem = Nothing
    ShowPrompt "Cliquez pres de l'element 3D a coter  (Reset=quitter)"
End Sub

'------------------------------------------------------------------------------
Private Sub SelectionnerElement(Point As Point3d)
    Dim oAtt As Object
    If Not TrouverElementProche(Point, g_oSettings.dTolTexte, False, _
                                m_oElem, oAtt) Then
        ShowPrompt "Aucun element lineaire du modele actif a moins de " & _
                   g_oSettings.dTolTexte & " u.m. - recliquez"
        Exit Sub
    End If

    If Not g_oMoteur.ExtrairePolyligne3D(m_oElem, _
            g_oSettings.dPasDiscretisation, m_dX, m_dY, m_dZ) Then
        ShowPrompt "Type non supporte - recliquez"
        Set m_oElem = Nothing
        Exit Sub
    End If

    m_dLong = g_oCalc.AbscissesPolyligne3D(m_dX, m_dY, m_dZ, m_dS)
    m_bGeomOk = True

    ShowPrompt "Element L=" & Format$(m_dLong, "0.00") & _
               " m  |  Data=creer le semis  Reset=abandonner"
    m_nEtape = 2
End Sub

'------------------------------------------------------------------------------
' Abscisses du semis selon les parametres courants (relus a chaque appel :
' le formulaire modeless peut les changer pendant l'apercu).
Private Function CalculerSemis(dSemis() As Double) As Long
    If g_oSettings.oSemis.Mode = semisPartsEgales Then
        CalculerSemis = g_oCalc.AbscissesSemisPartsEgales(m_dLong, _
                            g_oSettings.oSemis.Nombre, dSemis)
    Else
        CalculerSemis = g_oCalc.AbscissesSemisPasFixe(m_dLong, _
                            g_oSettings.oSemis.Pas, dSemis)
    End If
End Function
