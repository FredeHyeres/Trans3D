VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPoints 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmPoints.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmPoints - Formulaire dedie a la commande Points (Tool Settings)
'
' Cadres :
'   Cercle repere  : creer oui/non, diametre, couleur, niveau
'   Texte altitude : creer oui/non, duplication du style de la cote source ou
'                    creation (niveau, style de texte), decimales
'   Recherche      : tolerance de clic (u.m.)
'   Etat           : information de la commande en cours
'
' Tous les controles sont crees au runtime dans ConstruireControles
' (gardes m_bConstruit et m_bInit obligatoires, cf. spec 5.9).
' QueryClose : Cancel = 1 + Hide (jamais decharger un modeless).
'==============================================================================
Option Explicit

Private m_oSettings  As CSettings
Private m_bInit      As Boolean
Private m_bConstruit As Boolean
Private m_bAttenteZ  As Boolean

' --- Cercle repere ---
Private WithEvents chkCreerCercle As MSForms.CheckBox
Attribute chkCreerCercle.VB_VarHelpID = -1
Private WithEvents txtDiametre As MSForms.TextBox
Attribute txtDiametre.VB_VarHelpID = -1
Private WithEvents txtCouleurCercle As MSForms.TextBox
Attribute txtCouleurCercle.VB_VarHelpID = -1
Private WithEvents cmbNiveauCercle As MSForms.ComboBox
Attribute cmbNiveauCercle.VB_VarHelpID = -1
Private WithEvents chkCerclePlein As MSForms.CheckBox
Attribute chkCerclePlein.VB_VarHelpID = -1
' --- Texte altitude ---
Private WithEvents chkCreerTexte As MSForms.CheckBox
Attribute chkCreerTexte.VB_VarHelpID = -1
Private WithEvents chkTexteModele As MSForms.CheckBox
Attribute chkTexteModele.VB_VarHelpID = -1
Private WithEvents cmbNiveauTexte As MSForms.ComboBox
Attribute cmbNiveauTexte.VB_VarHelpID = -1
Private WithEvents cmbStyleTexte As MSForms.ComboBox
Attribute cmbStyleTexte.VB_VarHelpID = -1
Private WithEvents txtDecimales As MSForms.TextBox
Attribute txtDecimales.VB_VarHelpID = -1
' --- Recherche ---
Private WithEvents txtTolerance As MSForms.TextBox
Attribute txtTolerance.VB_VarHelpID = -1
' --- Saisie altitude ---
Private WithEvents txtSaisieZ As MSForms.TextBox
Attribute txtSaisieZ.VB_VarHelpID = -1
Private lblSaisieInfo As MSForms.Label
' --- Etat ---
Private lblEtat1     As MSForms.Label
Private lblEtat2     As MSForms.Label

'==============================================================================
' Construction des controles
'==============================================================================

Private Sub UserForm_Initialize()
    ConstruireControles
End Sub

Private Sub ConstruireControles()
    If m_bConstruit Then Exit Sub
    m_bConstruit = True

    Me.Caption = "Trans3D - Points"
    Me.Width = 212
    Me.Height = 448

    Dim dY As Double
    dY = 6

    ' --- Cadre Cercle repere --------------------------------------------------
    Dim fraCercle As MSForms.Frame
    Set fraCercle = Me.Controls.Add("Forms.Frame.1", "fraCercle")
    fraCercle.Caption = "Cercle repere (au Z reel)"
    fraCercle.Left = 6: fraCercle.Top = dY
    fraCercle.Width = 192: fraCercle.Height = 100

    Set chkCreerCercle = fraCercle.Controls.Add("Forms.CheckBox.1", "chkCreerCercle")
    chkCreerCercle.Caption = "Creer le cercle"
    chkCreerCercle.Left = 6: chkCreerCercle.Top = 10
    chkCreerCercle.Width = 180: chkCreerCercle.Height = 14

    CreerLabel fraCercle, "lblDiam", "Diametre :", 6, 28, 48
    Set txtDiametre = fraCercle.Controls.Add("Forms.TextBox.1", "txtDiametre")
    txtDiametre.Left = 56: txtDiametre.Top = 26
    txtDiametre.Width = 48: txtDiametre.Height = 16

    CreerLabel fraCercle, "lblCoulCer", "Couleur :", 6, 46, 42
    Set txtCouleurCercle = fraCercle.Controls.Add("Forms.TextBox.1", "txtCouleurCercle")
    txtCouleurCercle.Left = 56: txtCouleurCercle.Top = 44
    txtCouleurCercle.Width = 30: txtCouleurCercle.Height = 16

    CreerLabel fraCercle, "lblNivCer", "Niveau :", 6, 64, 42
    Set cmbNiveauCercle = fraCercle.Controls.Add("Forms.ComboBox.1", "cmbNiveauCercle")
    cmbNiveauCercle.Left = 56: cmbNiveauCercle.Top = 62
    cmbNiveauCercle.Width = 130: cmbNiveauCercle.Height = 16

    Set chkCerclePlein = fraCercle.Controls.Add("Forms.CheckBox.1", "chkCerclePlein")
    chkCerclePlein.Caption = "Plein (rempli)"
    chkCerclePlein.Left = 6: chkCerclePlein.Top = 80
    chkCerclePlein.Width = 180: chkCerclePlein.Height = 14

    dY = dY + 106

    ' --- Cadre Texte altitude -------------------------------------------------
    Dim fraTexte As MSForms.Frame
    Set fraTexte = Me.Controls.Add("Forms.Frame.1", "fraTexte")
    fraTexte.Caption = "Texte altitude"
    fraTexte.Left = 6: fraTexte.Top = dY
    fraTexte.Width = 192: fraTexte.Height = 104

    Set chkCreerTexte = fraTexte.Controls.Add("Forms.CheckBox.1", "chkCreerTexte")
    chkCreerTexte.Caption = "Texte a Z=0.00 (sinon Z reel)"
    chkCreerTexte.Left = 6: chkCreerTexte.Top = 10
    chkCreerTexte.Width = 180: chkCreerTexte.Height = 14

    Set chkTexteModele = fraTexte.Controls.Add("Forms.CheckBox.1", "chkTexteModele")
    chkTexteModele.Caption = "Dupliquer le style de la cote"
    chkTexteModele.Left = 6: chkTexteModele.Top = 26
    chkTexteModele.Width = 180: chkTexteModele.Height = 14

    CreerLabel fraTexte, "lblNivTxt", "Niveau :", 6, 46, 40
    Set cmbNiveauTexte = fraTexte.Controls.Add("Forms.ComboBox.1", "cmbNiveauTexte")
    cmbNiveauTexte.Left = 52: cmbNiveauTexte.Top = 44
    cmbNiveauTexte.Width = 134: cmbNiveauTexte.Height = 16

    CreerLabel fraTexte, "lblStyleTxt", "Style :", 6, 66, 40
    Set cmbStyleTexte = fraTexte.Controls.Add("Forms.ComboBox.1", "cmbStyleTexte")
    cmbStyleTexte.Left = 52: cmbStyleTexte.Top = 64
    cmbStyleTexte.Width = 134: cmbStyleTexte.Height = 16

    CreerLabel fraTexte, "lblDec", "Decimales :", 6, 86, 52
    Set txtDecimales = fraTexte.Controls.Add("Forms.TextBox.1", "txtDecimales")
    txtDecimales.Left = 60: txtDecimales.Top = 84
    txtDecimales.Width = 24: txtDecimales.Height = 16

    dY = dY + 110

    ' --- Cadre Recherche ------------------------------------------------------
    Dim fraRech As MSForms.Frame
    Set fraRech = Me.Controls.Add("Forms.Frame.1", "fraRech")
    fraRech.Caption = "Recherche"
    fraRech.Left = 6: fraRech.Top = dY
    fraRech.Width = 192: fraRech.Height = 38

    CreerLabel fraRech, "lblTol", "Tolerance clic (u.m.) :", 6, 14, 100
    Set txtTolerance = fraRech.Controls.Add("Forms.TextBox.1", "txtTolerance")
    txtTolerance.Left = 110: txtTolerance.Top = 12
    txtTolerance.Width = 48: txtTolerance.Height = 16

    dY = dY + 44

    ' --- Cadre Saisie altitude ------------------------------------------------
    Dim fraSaisie As MSForms.Frame
    Set fraSaisie = Me.Controls.Add("Forms.Frame.1", "fraSaisie")
    fraSaisie.Caption = "Saisie altitude manuelle"
    fraSaisie.Left = 6: fraSaisie.Top = dY
    fraSaisie.Width = 192: fraSaisie.Height = 38

    CreerLabel fraSaisie, "lblZ", "Z :", 6, 14, 16
    Set txtSaisieZ = fraSaisie.Controls.Add("Forms.TextBox.1", "txtSaisieZ")
    txtSaisieZ.Left = 24: txtSaisieZ.Top = 12
    txtSaisieZ.Width = 56: txtSaisieZ.Height = 16
    txtSaisieZ.Enabled = False

    Set lblSaisieInfo = CreerLabel(fraSaisie, "lblSaisieInfo", _
        "(Enter = valider)", 86, 14, 100)

    dY = dY + 44

    ' --- Cadre Etat -----------------------------------------------------------
    Dim fraEtat As MSForms.Frame
    Set fraEtat = Me.Controls.Add("Forms.Frame.1", "fraEtat")
    fraEtat.Caption = "Etat"
    fraEtat.Left = 6: fraEtat.Top = dY
    fraEtat.Width = 192: fraEtat.Height = 54

    Set lblEtat1 = CreerLabel(fraEtat, "lblEtat1", "-", 6, 12, 180)
    Set lblEtat2 = CreerLabel(fraEtat, "lblEtat2", "-", 6, 28, 180)
End Sub

'------------------------------------------------------------------------------
Private Function CreerLabel(oParent As MSForms.Frame, sNom As String, _
                            sCaption As String, dLeft As Double, dTop As Double, _
                            dWidth As Double) As MSForms.Label
    Set CreerLabel = oParent.Controls.Add("Forms.Label.1", sNom)
    CreerLabel.Caption = sCaption
    CreerLabel.Left = dLeft: CreerLabel.Top = dTop
    CreerLabel.Width = dWidth: CreerLabel.Height = 12
End Function

'==============================================================================
' Initialisation
'==============================================================================

Sub Initialiser(oSettings As CSettings)
    ConstruireControles
    Set m_oSettings = oSettings
    m_bInit = True

    ' Cercle repere
    chkCreerCercle.Value = m_oSettings.bCreerCercle
    txtDiametre.Text = Format$(m_oSettings.oCercle.Diametre, "0.00")
    txtCouleurCercle.Text = CStr(m_oSettings.oCercle.Couleur)
    RemplirNiveaux cmbNiveauCercle
    PositionnerNiveau cmbNiveauCercle, m_oSettings.oCercle.NomNiveau
    chkCerclePlein.Value = m_oSettings.oCercle.Plein

    ' Texte altitude
    chkCreerTexte.Value = m_oSettings.bTexteAZero
    chkTexteModele.Value = m_oSettings.oTexte.CommeModele
    RemplirNiveaux cmbNiveauTexte
    PositionnerNiveau cmbNiveauTexte, m_oSettings.oTexte.NomNiveau
    RemplirStyles cmbStyleTexte
    PositionnerStyle cmbStyleTexte, m_oSettings.oTexte.NomStyleTexte
    txtDecimales.Text = CStr(m_oSettings.oTexte.Decimales)
    ActiverChampsTexte

    ' Recherche
    txtTolerance.Text = Format$(m_oSettings.dTolTexte, "0.00")

    ReinitialiserEtat
    m_bInit = False
End Sub

'------------------------------------------------------------------------------
Private Sub RemplirNiveaux(cmb As MSForms.ComboBox)
    cmb.Clear
    cmb.AddItem ""
    On Error Resume Next
    Dim oLvl As Level
    For Each oLvl In ActiveDesignFile.Levels
        cmb.AddItem oLvl.Number & " : " & oLvl.Name
    Next
    On Error GoTo 0
    cmb.ListIndex = 0
End Sub

'------------------------------------------------------------------------------
Private Sub RemplirStyles(cmb As MSForms.ComboBox)
    cmb.Clear
    cmb.AddItem ""
    On Error Resume Next
    Dim oStyle As TextStyle
    For Each oStyle In ActiveDesignFile.TextStyles
        cmb.AddItem oStyle.Name
    Next
    On Error GoTo 0
    cmb.ListIndex = 0
End Sub

'------------------------------------------------------------------------------
Private Sub PositionnerStyle(cmb As MSForms.ComboBox, sNom As String)
    If Len(sNom) = 0 Then cmb.ListIndex = 0: Exit Sub
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = sNom Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next
    cmb.ListIndex = 0
End Sub

'------------------------------------------------------------------------------
' Le niveau et le style ne sont saisissables qu'en mode creation.
Private Sub ActiverChampsTexte()
    Dim bLibre As Boolean
    bLibre = Not m_oSettings.oTexte.CommeModele
    cmbNiveauTexte.Enabled = bLibre
    cmbStyleTexte.Enabled = bLibre
End Sub

'==============================================================================
' Mise a jour par la classe de commande
'==============================================================================

Sub AfficherElement(sTexte As String)
    If Not m_bConstruit Then Exit Sub
    lblEtat1.Caption = sTexte
End Sub

Sub AfficherZ(sTexte As String)
    If Not m_bConstruit Then Exit Sub
    lblEtat2.Caption = sTexte
End Sub

Sub ReinitialiserEtat()
    If Not m_bConstruit Then Exit Sub
    lblEtat1.Caption = "-"
    lblEtat2.Caption = "-"
End Sub

'------------------------------------------------------------------------------
' Active la saisie manuelle (appele par CPlacerPoints a l'etape 2).
Sub ActiverSaisieZ()
    If Not m_bConstruit Then Exit Sub
    m_bAttenteZ = True
    txtSaisieZ.Enabled = True
End Sub

'------------------------------------------------------------------------------
' Desactive la saisie manuelle (appele au retour etape 1).
Sub DesactiverSaisieZ()
    If Not m_bConstruit Then Exit Sub
    m_bAttenteZ = False
    txtSaisieZ.Enabled = False
    txtSaisieZ.Text = ""
End Sub

'------------------------------------------------------------------------------
' Retourne l'altitude saisie manuellement (vide = pas de saisie).
Property Get AltitudeManuelle() As String
    If Not m_bConstruit Then AltitudeManuelle = "": Exit Property
    AltitudeManuelle = Trim$(txtSaisieZ.Text)
End Property

'------------------------------------------------------------------------------
' Efface le champ apres utilisation.
Sub EffacerSaisieZ()
    If Not m_bConstruit Then Exit Sub
    txtSaisieZ.Text = ""
End Sub

Sub RafraichirTexte()
    If m_oSettings Is Nothing Then Exit Sub
    m_bInit = True
    cmbNiveauTexte.Text = m_oSettings.oTexte.NomNiveau
    txtDecimales.Text = CStr(m_oSettings.oTexte.Decimales)
    m_bInit = False
End Sub

'==============================================================================
' Evenements Cercle repere
'==============================================================================

Private Sub chkCreerCercle_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.bCreerCercle = (chkCreerCercle.Value = True)
End Sub

Private Sub txtDiametre_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dDiam As Double
    dDiam = Val(Replace(Trim$(txtDiametre.Text), ",", "."))
    If dDiam > 0 Then m_oSettings.oCercle.Diametre = dDiam
End Sub

Private Sub txtDiametre_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtDiametre.Text = Format$(m_oSettings.oCercle.Diametre, "0.00")
End Sub

Private Sub txtCouleurCercle_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String: sVal = Trim$(txtCouleurCercle.Text)
    If sVal = "" Then Exit Sub
    Dim nCoul As Long: nCoul = CLng(Val(sVal))
    If nCoul >= 0 And nCoul <= 255 Then m_oSettings.oCercle.Couleur = nCoul
End Sub

Private Sub cmbNiveauCercle_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oCercle.NomNiveau = ExtraireNiveau(cmbNiveauCercle.Text)
End Sub

Private Sub chkCerclePlein_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oCercle.Plein = (chkCerclePlein.Value = True)
End Sub

'==============================================================================
' Evenements Texte altitude
'==============================================================================

Private Sub chkCreerTexte_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.bTexteAZero = (chkCreerTexte.Value = True)
End Sub

Private Sub chkTexteModele_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.CommeModele = (chkTexteModele.Value = True)
    ActiverChampsTexte
    If m_oSettings.oTexte.CommeModele And m_oSettings.TextModeleDisponible Then
        m_oSettings.oTexte.ChargerDepuisElement m_oSettings.oTextModele
        RafraichirTexte
    End If
End Sub

Private Sub cmbNiveauTexte_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.NomNiveau = ExtraireNiveau(cmbNiveauTexte.Text)
End Sub

Private Sub cmbStyleTexte_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.NomStyleTexte = Trim$(cmbStyleTexte.Text)
End Sub

Private Sub txtDecimales_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim nDec As Integer
    nDec = CInt(Val(Trim$(txtDecimales.Text)))
    If nDec >= 0 And nDec <= 6 Then m_oSettings.oTexte.Decimales = nDec
End Sub

Private Sub txtDecimales_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                 ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtDecimales.Text = CStr(m_oSettings.oTexte.Decimales)
End Sub

'==============================================================================
' Evenements Recherche
'==============================================================================

Private Sub txtTolerance_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dTol As Double
    dTol = Val(Replace(Trim$(txtTolerance.Text), ",", "."))
    If dTol > 0 Then m_oSettings.dTolTexte = dTol
End Sub

Private Sub txtTolerance_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                 ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtTolerance.Text = Format$(m_oSettings.dTolTexte, "0.00")
End Sub

'==============================================================================
' Evenements Saisie altitude
'==============================================================================

Private Sub txtSaisieZ_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                               ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And m_bAttenteZ Then
        If Len(Trim$(txtSaisieZ.Text)) > 0 Then
            Dim pt As Point3d
            CadInputQueue.SendDataPoint pt, 1
        End If
    End If
End Sub

'------------------------------------------------------------------------------
Private Sub PositionnerNiveau(cmb As MSForms.ComboBox, sNom As String)
    If Len(sNom) = 0 Then cmb.ListIndex = 0: Exit Sub
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If ExtraireNiveau(cmb.List(i)) = sNom Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next
    cmb.ListIndex = 0
End Sub

Private Function ExtraireNiveau(ByVal sItem As String) As String
    sItem = Trim$(sItem)
    If InStr(sItem, " : ") > 0 Then
        ExtraireNiveau = Trim$(Mid$(sItem, InStr(sItem, " : ") + 3))
    Else
        ExtraireNiveau = sItem
    End If
End Function

'==============================================================================
' Fermeture
'==============================================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
        CommandState.StartDefaultCommand
    End If
End Sub

