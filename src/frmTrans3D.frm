VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTrans3D 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmTrans3D.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTrans3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmTrans3D - Formulaire commun aux trois commandes Trans3D (Tool Settings)
'
' Cadres :
'   Texte altitude : creer oui/non, duplication du style de la cote source ou
'                    creation (hauteur, couleur, niveau, police, fond),
'                    decimales et separateur
'   Cercle repere  : creer oui/non, diametre, couleur, niveau
'   Semis          : distance fixe (pas) ou parts egales (nombre)
'   Convertir      : pas de discretisation des courbes, symbologie source
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

' --- Texte altitude ---
Private WithEvents chkCreerTexte As MSForms.CheckBox
Attribute chkCreerTexte.VB_VarHelpID = -1
Private WithEvents chkTexteModele As MSForms.CheckBox
Attribute chkTexteModele.VB_VarHelpID = -1
Private WithEvents txtHauteur As MSForms.TextBox
Attribute txtHauteur.VB_VarHelpID = -1
Private WithEvents txtCouleurTexte As MSForms.TextBox
Attribute txtCouleurTexte.VB_VarHelpID = -1
Private WithEvents cmbNiveauTexte As MSForms.ComboBox
Attribute cmbNiveauTexte.VB_VarHelpID = -1
Private WithEvents txtPolice As MSForms.TextBox
Attribute txtPolice.VB_VarHelpID = -1
Private WithEvents chkFond As MSForms.CheckBox
Attribute chkFond.VB_VarHelpID = -1
Private WithEvents txtDecimales As MSForms.TextBox
Attribute txtDecimales.VB_VarHelpID = -1
Private WithEvents chkVirgule As MSForms.CheckBox
Attribute chkVirgule.VB_VarHelpID = -1
' --- Cercle repere ---
Private WithEvents chkCreerCercle As MSForms.CheckBox
Attribute chkCreerCercle.VB_VarHelpID = -1
Private WithEvents txtDiametre As MSForms.TextBox
Attribute txtDiametre.VB_VarHelpID = -1
Private WithEvents txtCouleurCercle As MSForms.TextBox
Attribute txtCouleurCercle.VB_VarHelpID = -1
Private WithEvents cmbNiveauCercle As MSForms.ComboBox
Attribute cmbNiveauCercle.VB_VarHelpID = -1
' --- Decoupage ---
Private WithEvents chkDecoupage As MSForms.CheckBox
Attribute chkDecoupage.VB_VarHelpID = -1
Private WithEvents optPasFixe As MSForms.OptionButton
Attribute optPasFixe.VB_VarHelpID = -1
Private WithEvents optPartsEgales As MSForms.OptionButton
Attribute optPartsEgales.VB_VarHelpID = -1
Private WithEvents txtPas As MSForms.TextBox
Attribute txtPas.VB_VarHelpID = -1
Private WithEvents txtNombre As MSForms.TextBox
Attribute txtNombre.VB_VarHelpID = -1
' --- Convertir ---
Private WithEvents txtDiscret As MSForms.TextBox
Attribute txtDiscret.VB_VarHelpID = -1
Private WithEvents chkSymboSource As MSForms.CheckBox
Attribute chkSymboSource.VB_VarHelpID = -1
' --- Recherche ---
Private WithEvents txtTolerance As MSForms.TextBox
Attribute txtTolerance.VB_VarHelpID = -1
' --- Etat ---
Private lblEtat1     As MSForms.Label   ' element selectionne (type, longueur)
Private lblEtat2     As MSForms.Label   ' altitudes depart / arrivee

'==============================================================================
' Construction des controles
'==============================================================================

Private Sub UserForm_Initialize()
    ConstruireControles
End Sub

Private Sub ConstruireControles()
    If m_bConstruit Then Exit Sub
    m_bConstruit = True

    Me.Caption = "Trans3D"
    Me.Width = 212
    Me.Height = 514

    Dim dY As Double
    dY = 6

    ' --- Cadre Texte altitude -------------------------------------------------
    Dim fraTexte As MSForms.Frame
    Set fraTexte = Me.Controls.Add("Forms.Frame.1", "fraTexte")
    fraTexte.Caption = "Texte altitude"
    fraTexte.Left = 6: fraTexte.Top = dY
    fraTexte.Width = 192: fraTexte.Height = 122

    Set chkCreerTexte = fraTexte.Controls.Add("Forms.CheckBox.1", "chkCreerTexte")
    chkCreerTexte.Caption = "Texte a Z=0.00 (sinon Z reel)"
    chkCreerTexte.Left = 6: chkCreerTexte.Top = 10
    chkCreerTexte.Width = 180: chkCreerTexte.Height = 14

    Set chkTexteModele = fraTexte.Controls.Add("Forms.CheckBox.1", "chkTexteModele")
    chkTexteModele.Caption = "Dupliquer le style de la cote"
    chkTexteModele.Left = 6: chkTexteModele.Top = 26
    chkTexteModele.Width = 180: chkTexteModele.Height = 14

    CreerLabel fraTexte, "lblHaut", "Hauteur :", 6, 46, 44
    Set txtHauteur = fraTexte.Controls.Add("Forms.TextBox.1", "txtHauteur")
    txtHauteur.Left = 52: txtHauteur.Top = 44
    txtHauteur.Width = 36: txtHauteur.Height = 16

    CreerLabel fraTexte, "lblCoulTxt", "Coul. :", 96, 46, 30
    Set txtCouleurTexte = fraTexte.Controls.Add("Forms.TextBox.1", "txtCouleurTexte")
    txtCouleurTexte.Left = 128: txtCouleurTexte.Top = 44
    txtCouleurTexte.Width = 28: txtCouleurTexte.Height = 16

    CreerLabel fraTexte, "lblNivTxt", "Niveau :", 6, 64, 40
    Set cmbNiveauTexte = fraTexte.Controls.Add("Forms.ComboBox.1", "cmbNiveauTexte")
    cmbNiveauTexte.Left = 52: cmbNiveauTexte.Top = 62
    cmbNiveauTexte.Width = 134: cmbNiveauTexte.Height = 16

    CreerLabel fraTexte, "lblPolice", "Police :", 6, 82, 40
    Set txtPolice = fraTexte.Controls.Add("Forms.TextBox.1", "txtPolice")
    txtPolice.Left = 52: txtPolice.Top = 80
    txtPolice.Width = 60: txtPolice.Height = 16

    Set chkFond = fraTexte.Controls.Add("Forms.CheckBox.1", "chkFond")
    chkFond.Caption = "Fond"
    chkFond.Left = 120: chkFond.Top = 81
    chkFond.Width = 60: chkFond.Height = 14

    CreerLabel fraTexte, "lblDec", "Decimales :", 6, 102, 52
    Set txtDecimales = fraTexte.Controls.Add("Forms.TextBox.1", "txtDecimales")
    txtDecimales.Left = 60: txtDecimales.Top = 100
    txtDecimales.Width = 24: txtDecimales.Height = 16

    Set chkVirgule = fraTexte.Controls.Add("Forms.CheckBox.1", "chkVirgule")
    chkVirgule.Caption = "Virgule (12,35)"
    chkVirgule.Left = 96: chkVirgule.Top = 101
    chkVirgule.Width = 90: chkVirgule.Height = 14

    dY = dY + 128

    ' --- Cadre Cercle repere --------------------------------------------------
    Dim fraCercle As MSForms.Frame
    Set fraCercle = Me.Controls.Add("Forms.Frame.1", "fraCercle")
    fraCercle.Caption = "Cercle repere (au Z reel)"
    fraCercle.Left = 6: fraCercle.Top = dY
    fraCercle.Width = 192: fraCercle.Height = 86

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

    dY = dY + 92

    ' --- Cadre Decoupage ------------------------------------------------------
    Dim fraDecoupage As MSForms.Frame
    Set fraDecoupage = Me.Controls.Add("Forms.Frame.1", "fraDecoupage")
    fraDecoupage.Caption = "Decoupage"
    fraDecoupage.Left = 6: fraDecoupage.Top = dY
    fraDecoupage.Width = 192: fraDecoupage.Height = 74

    Set chkDecoupage = fraDecoupage.Controls.Add("Forms.CheckBox.1", "chkDecoupage")
    chkDecoupage.Caption = "Creer des points intermediaires"
    chkDecoupage.Left = 6: chkDecoupage.Top = 10
    chkDecoupage.Width = 180: chkDecoupage.Height = 14

    Set optPasFixe = fraDecoupage.Controls.Add("Forms.OptionButton.1", "optPasFixe")
    optPasFixe.Caption = "Distance fixe (m)"
    optPasFixe.Left = 6: optPasFixe.Top = 28
    optPasFixe.Width = 110: optPasFixe.Height = 14

    Set txtPas = fraDecoupage.Controls.Add("Forms.TextBox.1", "txtPas")
    txtPas.Left = 130: txtPas.Top = 27
    txtPas.Width = 48: txtPas.Height = 16

    Set optPartsEgales = fraDecoupage.Controls.Add("Forms.OptionButton.1", "optPartsEgales")
    optPartsEgales.Caption = "Nombre de points"
    optPartsEgales.Left = 6: optPartsEgales.Top = 48
    optPartsEgales.Width = 110: optPartsEgales.Height = 14

    Set txtNombre = fraDecoupage.Controls.Add("Forms.TextBox.1", "txtNombre")
    txtNombre.Left = 130: txtNombre.Top = 47
    txtNombre.Width = 48: txtNombre.Height = 16

    dY = dY + 80

    ' --- Cadre Convertir ------------------------------------------------------
    Dim fraConv As MSForms.Frame
    Set fraConv = Me.Controls.Add("Forms.Frame.1", "fraConv")
    fraConv.Caption = "Convertir"
    fraConv.Left = 6: fraConv.Top = dY
    fraConv.Width = 192: fraConv.Height = 56

    CreerLabel fraConv, "lblDiscret", "Pas discretisation :", 6, 12, 88
    Set txtDiscret = fraConv.Controls.Add("Forms.TextBox.1", "txtDiscret")
    txtDiscret.Left = 98: txtDiscret.Top = 10
    txtDiscret.Width = 48: txtDiscret.Height = 16

    Set chkSymboSource = fraConv.Controls.Add("Forms.CheckBox.1", "chkSymboSource")
    chkSymboSource.Caption = "Symbologie de l'element source"
    chkSymboSource.Left = 6: chkSymboSource.Top = 32
    chkSymboSource.Width = 180: chkSymboSource.Height = 14

    dY = dY + 62

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

    ' Texte altitude
    chkCreerTexte.Value = m_oSettings.bTexteAZero
    chkTexteModele.Value = m_oSettings.oTexte.CommeModele
    txtHauteur.Text = Format$(m_oSettings.oTexte.Hauteur, "0.00")
    txtCouleurTexte.Text = CStr(m_oSettings.oTexte.Couleur)
    RemplirNiveaux cmbNiveauTexte
    txtPolice.Text = m_oSettings.oTexte.NomPolice
    chkFond.Value = m_oSettings.oTexte.FondActif
    txtDecimales.Text = CStr(m_oSettings.oTexte.Decimales)
    chkVirgule.Value = (m_oSettings.oTexte.SepDecimal = ",")
    ActiverChampsTexte

    ' Cercle repere
    chkCreerCercle.Value = m_oSettings.bCreerCercle
    txtDiametre.Text = Format$(m_oSettings.oCercle.Diametre, "0.00")
    txtCouleurCercle.Text = CStr(m_oSettings.oCercle.Couleur)
    RemplirNiveaux cmbNiveauCercle

    ' Decoupage
    chkDecoupage.Value = m_oSettings.bDecoupage
    optPasFixe.Value = (m_oSettings.oSemis.Mode = semisDistanceFixe)
    optPartsEgales.Value = (m_oSettings.oSemis.Mode = semisPartsEgales)
    txtPas.Text = Format$(m_oSettings.oSemis.Pas, "0.00")
    txtNombre.Text = CStr(m_oSettings.oSemis.Nombre)
    ActiverChampsDecoupage

    ' Convertir
    txtDiscret.Text = Format$(m_oSettings.dPasDiscretisation, "0.00")
    chkSymboSource.Value = m_oSettings.bSymboSource

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
' Les champs de style ne sont saisissables qu'en mode creation.
Private Sub ActiverChampsTexte()
    Dim bLibre As Boolean
    bLibre = Not m_oSettings.oTexte.CommeModele
    txtHauteur.Enabled = bLibre
    txtCouleurTexte.Enabled = bLibre
    cmbNiveauTexte.Enabled = bLibre
    txtPolice.Enabled = bLibre
    chkFond.Enabled = bLibre
End Sub

'------------------------------------------------------------------------------
Private Sub ActiverChampsDecoupage()
    Dim bActif As Boolean
    bActif = m_oSettings.bDecoupage
    optPasFixe.Enabled = bActif
    optPartsEgales.Enabled = bActif
    txtPas.Enabled = bActif And (m_oSettings.oSemis.Mode = semisDistanceFixe)
    txtNombre.Enabled = bActif And (m_oSettings.oSemis.Mode = semisPartsEgales)
End Sub

'==============================================================================
' Mise a jour par les classes de commande
'==============================================================================

'------------------------------------------------------------------------------
' Ligne 1 du cadre Etat : element selectionne (type, longueur).
Sub AfficherElement(sTexte As String)
    If Not m_bConstruit Then Exit Sub
    lblEtat1.Caption = sTexte
End Sub

'------------------------------------------------------------------------------
' Ligne 2 du cadre Etat : altitudes depart / arrivee, pente.
Sub AfficherZ(sTexte As String)
    If Not m_bConstruit Then Exit Sub
    lblEtat2.Caption = sTexte
End Sub

'------------------------------------------------------------------------------
' Appele par les classes de commande au Reset (spec 5.9).
Sub ReinitialiserEtat()
    If Not m_bConstruit Then Exit Sub
    lblEtat1.Caption = "-"
    lblEtat2.Caption = "-"
End Sub

'------------------------------------------------------------------------------
' Recale les champs texte apres lecture d'une cote source (mode duplication).
Sub RafraichirTexte()
    If m_oSettings Is Nothing Then Exit Sub
    m_bInit = True
    txtCouleurTexte.Text = CStr(m_oSettings.oTexte.Couleur)
    cmbNiveauTexte.Text = m_oSettings.oTexte.NomNiveau
    txtDecimales.Text = CStr(m_oSettings.oTexte.Decimales)
    chkVirgule.Value = (m_oSettings.oTexte.SepDecimal = ",")
    m_bInit = False
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

Private Sub txtHauteur_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dH As Double
    dH = Val(Replace(Trim$(txtHauteur.Text), ",", "."))
    If dH > 0 Then m_oSettings.oTexte.Hauteur = dH
End Sub

Private Sub txtHauteur_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                               ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtHauteur.Text = Format$(m_oSettings.oTexte.Hauteur, "0.00")
End Sub

Private Sub txtCouleurTexte_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String: sVal = Trim$(txtCouleurTexte.Text)
    If sVal = "" Then Exit Sub
    Dim nCoul As Long: nCoul = CLng(Val(sVal))
    If nCoul >= 0 And nCoul <= 255 Then m_oSettings.oTexte.Couleur = nCoul
End Sub

Private Sub cmbNiveauTexte_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.NomNiveau = ExtraireNiveau(cmbNiveauTexte.Text)
End Sub

Private Sub txtPolice_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.NomPolice = Trim$(txtPolice.Text)
End Sub

Private Sub chkFond_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.FondActif = (chkFond.Value = True)
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

Private Sub chkVirgule_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    If chkVirgule.Value = True Then
        m_oSettings.oTexte.SepDecimal = ","
    Else
        m_oSettings.oTexte.SepDecimal = "."
    End If
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

'==============================================================================
' Evenements Decoupage
'==============================================================================

Private Sub chkDecoupage_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.bDecoupage = (chkDecoupage.Value = True)
    ActiverChampsDecoupage
End Sub

Private Sub optPasFixe_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    If optPasFixe.Value = True Then m_oSettings.oSemis.Mode = semisDistanceFixe
    ActiverChampsDecoupage
End Sub

Private Sub optPartsEgales_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    If optPartsEgales.Value = True Then m_oSettings.oSemis.Mode = semisPartsEgales
    ActiverChampsDecoupage
End Sub

Private Sub txtPas_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dPas As Double
    dPas = Val(Replace(Trim$(txtPas.Text), ",", "."))
    If dPas > 0 Then m_oSettings.oSemis.Pas = dPas
End Sub

Private Sub txtPas_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                           ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtPas.Text = Format$(m_oSettings.oSemis.Pas, "0.00")
End Sub

Private Sub txtNombre_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim nNb As Long
    nNb = CLng(Val(Trim$(txtNombre.Text)))
    If nNb >= 1 Then m_oSettings.oSemis.Nombre = nNb
End Sub

Private Sub txtNombre_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtNombre.Text = CStr(m_oSettings.oSemis.Nombre)
End Sub

'==============================================================================
' Evenements Convertir
'==============================================================================

Private Sub txtDiscret_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dPas As Double
    dPas = Val(Replace(Trim$(txtDiscret.Text), ",", "."))
    If dPas > 0 Then m_oSettings.dPasDiscretisation = dPas
End Sub

Private Sub txtDiscret_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                               ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtDiscret.Text = Format$(m_oSettings.dPasDiscretisation, "0.00")
End Sub

Private Sub chkSymboSource_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.bSymboSource = (chkSymboSource.Value = True)
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

'------------------------------------------------------------------------------
' Extrait le nom du niveau d'un item "numero : nom" de combo.
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

