VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlacerConvertir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CPlacerConvertir - Machine a etats : conversion d'un element 2D en 3D
'
' Etape 1 (SELECTION) : clic pres de l'element 2D a convertir
'   Data  = selectionner l'element lineaire le plus proche (actif + refs)
'   Reset = quitter la commande
' Etape 2 (Z DEPART) : clic pres de la cote d'altitude de depart
'   (element ferme : cote unique, la creation suit immediatement)
'   Data  = lire la cote (texte/cellule/tag)
'   Reset = revenir a la selection (etape 1)
' Etape 3 (Z FIN) : clic pres de la cote d'altitude de fin
'   Data  = lire la cote
'   Reset = revenir au Z de depart (etape 2)
' Etape 4 (INTERMEDIAIRES) : clics sur les cotes intermediaires (0..N)
'   Data  = ajouter une altitude intermediaire (projetee sur l'element)
'   Reset = CREER l'element 3D puis retour etape 1
'
' L'element cree garde son type natif quand le Z est constant (arc, cercle),
' sinon polyligne 3D avec un sommet insere a chaque altitude intermediaire.
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private m_nEtape    As Integer      ' 1..4
Private m_oElem     As Element      ' element 2D selectionne
Private m_oAtt      As Object       ' attachement source (Nothing = modele actif)
Private m_dX()      As Double       ' polyligne extraite (repere modele actif)
Private m_dY()      As Double
Private m_dS()      As Double       ' abscisses curvilignes
Private m_dLong     As Double       ' longueur developpee
Private m_bGeomOk   As Boolean      ' garde : tableaux alloues
Private m_bFerme    As Boolean      ' cercle/ellipse complet : Z constant
Private m_dSRef()   As Double       ' abscisses des altitudes de reference
Private m_dZRef()   As Double       ' altitudes de reference
Private m_nNbRef    As Long

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Trans3D - Convertir"
    RetourSelection
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Select Case m_nEtape
    Case 1
        SelectionnerElement Point

    Case 2
        ' --- Cote de depart (ou cote unique si element ferme) ---
        Dim dAltDep As Double, dPX As Double, dPY As Double
        If Not LireCote(Point, True, dAltDep, dPX, dPY) Then Exit Sub

        AjouterRef 0#, dAltDep
        If m_bFerme Then
            CreerElement3DFinal
        Else
            ShowPrompt "Depart Z=" & FormatZ(dAltDep) & _
                       "  |  Cliquez la cote de FIN  (Reset=autre element)"
            m_nEtape = 3
        End If

    Case 3
        ' --- Cote de fin ---
        Dim dAltFin As Double, dFX As Double, dFY As Double
        If Not LireCote(Point, False, dAltFin, dFX, dFY) Then Exit Sub

        AjouterRef m_dLong, dAltFin
        Dim dPente As Double
        dPente = g_oCalc.PentePctLongueur(m_dZRef(0), dAltFin, m_dLong)
        ShowPrompt "Fin Z=" & FormatZ(dAltFin) & "  Pente=" & _
                   g_oCalc.FormatPente(dPente, g_oSettings.oTexte.SepDecimal) & _
                   "  |  Data=cote intermediaire  Reset=CREER l'element 3D"
        m_nEtape = 4

    Case 4
        ' --- Altitude intermediaire : projetee sur l'element ---
        Dim dAltInt As Double, dIX As Double, dIY As Double
        If Not LireCote(Point, False, dAltInt, dIX, dIY) Then Exit Sub

        Dim dDist As Double, dAbs As Double
        dAbs = g_oCalc.AbscisseProjetee2D(m_dX, m_dY, m_dS, dIX, dIY, dDist)
        AjouterRef dAbs, dAltInt
        ShowPrompt "Intermediaire " & (m_nNbRef - 2) & " : Z=" & _
                   FormatZ(dAltInt) & " a " & Format$(dAbs, "0.00") & _
                   " m  |  Data=autre cote  Reset=CREER l'element 3D"
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    If Not m_bGeomOk Then Exit Sub

    ' Surlignage a plat de l'element selectionne (etapes 2 a 4)
    g_oMoteur.DessinSurlignage2D DrawMode, m_dX, m_dY

    ' Apercu de l'element 3D des que depart + fin sont connus
    If m_nEtape = 4 And m_nNbRef >= 2 Then
        Dim dSRefTri() As Double, dZRefTri() As Double
        dSRefTri = m_dSRef: dZRefTri = m_dZRef
        g_oCalc.TrierRefsParAbscisse dSRefTri, dZRefTri, m_nNbRef

        Dim dZ() As Double
        g_oCalc.CalculerZPolyligne m_dS, dSRefTri, dZRefTri, m_nNbRef, dZ
        g_oMoteur.DessinPolyligneDyn DrawMode, m_dX, m_dY, dZ
    End If
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    Select Case m_nEtape
    Case 1
        ' Quitter la commande
        ' TODO formulaire : frmTrans3D.Hide
        ShowPrompt "Trans3D - Convertir : termine."
        CommandState.StartDefaultCommand

    Case 2
        RetourSelection

    Case 3
        m_nNbRef = 0
        ShowPrompt "Cliquez la cote de DEPART  (Reset=autre element)"
        m_nEtape = 2

    Case 4
        CreerElement3DFinal
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

'==============================================================================
' Etapes
'==============================================================================

'------------------------------------------------------------------------------
Private Sub RetourSelection()
    m_nEtape = 1
    m_bGeomOk = False
    m_bFerme = False
    Set m_oElem = Nothing
    Set m_oAtt = Nothing
    m_nNbRef = 0
    ShowPrompt "Cliquez pres de l'element 2D a convertir  (Reset=quitter)"
End Sub

'------------------------------------------------------------------------------
Private Sub SelectionnerElement(Point As Point3d)
    If Not TrouverElementProche(Point, g_oSettings.dTolTexte, True, _
                                m_oElem, m_oAtt) Then
        ShowPrompt "Aucun element lineaire a moins de " & _
                   g_oSettings.dTolTexte & " u.m. (B-spline ? droppez-la" & _
                   " en ligne brisee) - recliquez"
        Exit Sub
    End If

    If Not g_oMoteur.ExtrairePolyligne2D(m_oElem, m_oAtt, _
            g_oSettings.dPasDiscretisation, m_dX, m_dY) Then
        ShowPrompt "Type non supporte - recliquez"
        Set m_oElem = Nothing
        Set m_oAtt = Nothing
        Exit Sub
    End If

    m_dLong = g_oCalc.AbscissesPolyligne2D(m_dX, m_dY, m_dS)
    m_bGeomOk = True
    m_bFerme = g_oMoteur.EstFerme(m_oElem)
    m_nNbRef = 0
    ReDim m_dSRef(0 To 15)
    ReDim m_dZRef(0 To 15)

    If m_bFerme Then
        ShowPrompt "Element ferme L=" & Format$(m_dLong, "0.00") & _
                   " m  |  Cliquez la cote d'altitude (Z constant)" & _
                   "  (Reset=autre element)"
    Else
        ShowPrompt "Element L=" & Format$(m_dLong, "0.00") & _
                   " m  |  Cliquez la cote de DEPART  (Reset=autre element)"
    End If
    m_nEtape = 2
End Sub

'------------------------------------------------------------------------------
' Creation de l'element 3D d'apres les references saisies, puis retour etape 1.
Private Sub CreerElement3DFinal()
    If m_nNbRef < 1 Then
        RetourSelection
        Exit Sub
    End If

    g_oCalc.TrierRefsParAbscisse m_dSRef, m_dZRef, m_nNbRef

    ' Z constant : toutes les references a la meme altitude
    Dim bZConst As Boolean
    bZConst = True
    Dim i As Long
    For i = 1 To m_nNbRef - 1
        If Abs(m_dZRef(i) - m_dZRef(0)) > 0.0001 Then
            bZConst = False
            Exit For
        End If
    Next

    ' Type natif conserve si possible (arc/cercle a Z constant)
    Dim oNouv As Element
    Set oNouv = Nothing
    If bZConst Then
        Select Case m_oElem.Type
            Case msdElementTypeArc
                Set oNouv = g_oMoteur.CreerArc3DZConstant(m_oElem, _
                                m_dZRef(0), m_oAtt, g_oSettings)
            Case msdElementTypeEllipse
                Set oNouv = g_oMoteur.CreerCercle3DZConstant(m_oElem, _
                                m_dZRef(0), m_oAtt, g_oSettings)
        End Select
    End If

    ' Sinon : polyligne 3D, avec un sommet a chaque altitude intermediaire
    If oNouv Is Nothing Then
        If m_nNbRef > 2 Then
            g_oCalc.InsererAbscisses m_dX, m_dY, m_dS, m_dSRef, m_nNbRef
        End If
        Dim dZ() As Double
        g_oCalc.CalculerZPolyligne m_dS, m_dSRef, m_dZRef, m_nNbRef, dZ
        Set oNouv = g_oMoteur.CreerPolyligne3D(m_dX, m_dY, dZ, m_oElem, _
                                               g_oSettings)
    End If

    Dim sBilan As String
    If bZConst Then
        sBilan = "Element 3D cree a Z=" & FormatZ(m_dZRef(0))
    Else
        Dim dPente As Double
        dPente = g_oCalc.PentePctLongueur(m_dZRef(0), m_dZRef(m_nNbRef - 1), _
                                          m_dLong)
        sBilan = "Element 3D cree (pente moy. " & _
                 g_oCalc.FormatPente(dPente, g_oSettings.oTexte.SepDecimal) & ")"
    End If
    ShowPrompt sBilan & "  |  Cliquez le prochain element  (Reset=quitter)"

    RetourSelection
End Sub

'==============================================================================
' Utilitaires prives
'==============================================================================

'------------------------------------------------------------------------------
' Lit la cote la plus proche du clic. bChargerModele : memorise le texte
' source comme modele de style (fait sur la cote de depart).
' Renvoie l'altitude et la position de la cote (repere du modele actif).
Private Function LireCote(Point As Point3d, ByVal bChargerModele As Boolean, _
                          dAlt As Double, dPosX As Double, _
                          dPosY As Double) As Boolean
    LireCote = False

    Dim oText As TextElement
    If Not TrouverAltitudeProche(Point, g_oSettings.dTolTexte, oText) Then
        ShowPrompt "Aucune cote a moins de " & g_oSettings.dTolTexte & _
                   " u.m. - recliquez"
        Exit Function
    End If

    Dim oCote As New CPointRef
    Dim sSep As String, nDec As Integer
    If Not g_oCalc.LireAltitudeTexte(oCote, sSep, nDec) Then
        ShowPrompt "Texte '" & g_oSelectionCourante.ValeurTexte & _
                   "' non numerique - recliquez"
        Exit Function
    End If

    If bChargerModele Then
        If Not oText Is Nothing Then
            g_oSettings.ChargerModeleTexte oText, sSep, nDec
        ElseIf Not g_oSelectionCourante.oTag Is Nothing Then
            g_oSettings.ChargerModeleDepuisTag g_oSelectionCourante.oTag, _
                                               sSep, nDec
        Else
            g_oSettings.oTexte.SepDecimal = sSep
            g_oSettings.AjusterDecimales nDec
        End If
    Else
        g_oSettings.AjusterDecimales nDec
    End If

    dAlt = oCote.Altitude
    dPosX = oCote.X
    dPosY = oCote.Y
    LireCote = True
End Function

'------------------------------------------------------------------------------
Private Sub AjouterRef(ByVal dAbs As Double, ByVal dZ As Double)
    If m_nNbRef > UBound(m_dSRef) Then
        ReDim Preserve m_dSRef(0 To UBound(m_dSRef) + 8)
        ReDim Preserve m_dZRef(0 To UBound(m_dZRef) + 8)
    End If
    m_dSRef(m_nNbRef) = dAbs
    m_dZRef(m_nNbRef) = dZ
    m_nNbRef = m_nNbRef + 1
End Sub

'------------------------------------------------------------------------------
Private Function FormatZ(ByVal dZ As Double) As String
    FormatZ = g_oCalc.FormatAltitude(dZ, g_oSettings.oTexte.Decimales, _
                                     g_oSettings.oTexte.SepDecimal)
End Function
