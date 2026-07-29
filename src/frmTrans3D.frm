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
'                    creation (niveau, style de texte), decimales
'   Cercle repere  : creer oui/non, diametre, couleur, niveau
'   Semis          : distance fixe (pas) ou parts egales (nombre)
'   Convertir      : pas de discretisation des courbes, arc natif ou polyligne
'   Symbologie elements 3D : dupliquer la source, ou attributs actifs, ou
'                    niveau/couleur/style de ligne/epaisseur explicites
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
Private WithEvents cmbNiveauTexte As MSForms.ComboBox
Attribute cmbNiveauTexte.VB_VarHelpID = -1
Private WithEvents cmbStyleTexte As MSForms.ComboBox
Attribute cmbStyleTexte.VB_VarHelpID = -1
Private WithEvents txtDecimales As MSForms.TextBox
Attribute txtDecimales.VB_VarHelpID = -1
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
Private WithEvents optAuxSommets As MSForms.OptionButton
Attribute optAuxSommets.VB_VarHelpID = -1
Private WithEvents optZInterpole As MSForms.OptionButton
Attribute optZInterpole.VB_VarHelpID = -1
Private WithEvents optZDefini As MSForms.OptionButton
Attribute optZDefini.VB_VarHelpID = -1
' --- Convertir Cercle ---
Private WithEvents optArcNatif As MSForms.OptionButton
Attribute optArcNatif.VB_VarHelpID = -1
Private WithEvents optPolyligne As MSForms.OptionButton
Attribute optPolyligne.VB_VarHelpID = -1
Private WithEvents txtDiscret As MSForms.TextBox
Attribute txtDiscret.VB_VarHelpID = -1
' --- Symbologie elements 3D ---
Private WithEvents chkSymboSource As MSForms.CheckBox
Attribute chkSymboSource.VB_VarHelpID = -1
Private WithEvents chkAttributsActifs As MSForms.CheckBox
Attribute chkAttributsActifs.VB_VarHelpID = -1
Private WithEvents cmbNiveauElements As MSForms.ComboBox
Attribute cmbNiveauElements.VB_VarHelpID = -1
Private WithEvents txtCouleurElements As MSForms.TextBox
Attribute txtCouleurElements.VB_VarHelpID = -1
Private WithEvents txtStyleElements As MSForms.TextBox
Attribute txtStyleElements.VB_VarHelpID = -1
Private WithEvents txtEpaisseurElements As MSForms.TextBox
Attribute txtEpaisseurElements.VB_VarHelpID = -1
' --- Recherche ---
Private WithEvents txtTolerance As MSForms.TextBox
Attribute txtTolerance.VB_VarHelpID = -1
' --- Etat ---
Private lblEtat1     As MSForms.Label   ' element selectionne (type, longueur)
Private lblEtat2     As MSForms.Label   ' altitudes depart / arrivee
Private lblEtat3     As MSForms.Label   ' liste Z sommets (etape 5)

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
    Me.Height = 706

    Dim dY As Double
    dY = 6

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
    fraDecoupage.Width = 192: fraDecoupage.Height = 108

    Set chkDecoupage = fraDecoupage.Controls.Add("Forms.CheckBox.1", "chkDecoupage")
    chkDecoupage.Caption = "Creer des points intermediaires"
    chkDecoupage.Left = 6: chkDecoupage.Top = 10
    chkDecoupage.Width = 180: chkDecoupage.Height = 14

    Set optPasFixe = fraDecoupage.Controls.Add("Forms.OptionButton.1", "optPasFixe")
    optPasFixe.Caption = "Distance fixe (m)"
    optPasFixe.Left = 6: optPasFixe.Top = 28
    optPasFixe.Width = 110: optPasFixe.Height = 14
    optPasFixe.GroupName = "ModeSemis"

    Set txtPas = fraDecoupage.Controls.Add("Forms.TextBox.1", "txtPas")
    txtPas.Left = 130: txtPas.Top = 27
    txtPas.Width = 48: txtPas.Height = 16

    Set optPartsEgales = fraDecoupage.Controls.Add("Forms.OptionButton.1", "optPartsEgales")
    optPartsEgales.Caption = "Nombre de points"
    optPartsEgales.Left = 6: optPartsEgales.Top = 48
    optPartsEgales.Width = 110: optPartsEgales.Height = 14
    optPartsEgales.GroupName = "ModeSemis"

    Set txtNombre = fraDecoupage.Controls.Add("Forms.TextBox.1", "txtNombre")
    txtNombre.Left = 130: txtNombre.Top = 47
    txtNombre.Width = 48: txtNombre.Height = 16

    Set optAuxSommets = fraDecoupage.Controls.Add("Forms.OptionButton.1", "optAuxSommets")
    optAuxSommets.Caption = "Aux sommets"
    optAuxSommets.Left = 6: optAuxSommets.Top = 68
    optAuxSommets.Width = 90: optAuxSommets.Height = 14
    optAuxSommets.GroupName = "ModeSemis"

    Set optZInterpole = fraDecoupage.Controls.Add("Forms.OptionButton.1", "optZInterpole")
    optZInterpole.Caption = "Z interpole"
    optZInterpole.Left = 24: optZInterpole.Top = 86
    optZInterpole.Width = 76: optZInterpole.Height = 14
    optZInterpole.GroupName = "ModeZSommet"

    Set optZDefini = fraDecoupage.Controls.Add("Forms.OptionButton.1", "optZDefini")
    optZDefini.Caption = "Z defini"
    optZDefini.Left = 106: optZDefini.Top = 86
    optZDefini.Width = 76: optZDefini.Height = 14
    optZDefini.GroupName = "ModeZSommet"

    dY = dY + 114

    ' --- Cadre Convertir Cercle ------------------------------------------------
    Dim fraConv As MSForms.Frame
    Set fraConv = Me.Controls.Add("Forms.Frame.1", "fraConv")
    fraConv.Caption = "Convertir Cercle"
    fraConv.Left = 6: fraConv.Top = dY
    fraConv.Width = 192: fraConv.Height = 52

    Set optArcNatif = fraConv.Controls.Add("Forms.OptionButton.1", "optArcNatif")
    optArcNatif.Caption = "Arc 3D"
    optArcNatif.Left = 6: optArcNatif.Top = 12
    optArcNatif.Width = 60: optArcNatif.Height = 14
    optArcNatif.GroupName = "ModeArc"

    Set optPolyligne = fraConv.Controls.Add("Forms.OptionButton.1", "optPolyligne")
    optPolyligne.Caption = "Polyligne"
    optPolyligne.Left = 72: optPolyligne.Top = 12
    optPolyligne.Width = 70: optPolyligne.Height = 14
    optPolyligne.GroupName = "ModeArc"

    CreerLabel fraConv, "lblDiscret", "Pas discretisation :", 6, 30, 88
    Set txtDiscret = fraConv.Controls.Add("Forms.TextBox.1", "txtDiscret")
    txtDiscret.Left = 98: txtDiscret.Top = 28
    txtDiscret.Width = 48: txtDiscret.Height = 16

    dY = dY + 58

    ' --- Cadre Symbologie elements 3D ------------------------------------------
    Dim fraElements As MSForms.Frame
    Set fraElements = Me.Controls.Add("Forms.Frame.1", "fraElements")
    fraElements.Caption = "Symbologie elements 3D"
    fraElements.Left = 6: fraElements.Top = dY
    fraElements.Width = 192: fraElements.Height = 122

    Set chkSymboSource = fraElements.Controls.Add("Forms.CheckBox.1", "chkSymboSource")
    chkSymboSource.Caption = "Dupliquer la symbologie de l'element source"
    chkSymboSource.Left = 6: chkSymboSource.Top = 10
    chkSymboSource.Width = 180: chkSymboSource.Height = 14

    Set chkAttributsActifs = fraElements.Controls.Add("Forms.CheckBox.1", "chkAttributsActifs")
    chkAttributsActifs.Caption = "Attributs actifs (sinon ci-dessous)"
    chkAttributsActifs.Left = 6: chkAttributsActifs.Top = 26
    chkAttributsActifs.Width = 180: chkAttributsActifs.Height = 14

    CreerLabel fraElements, "lblNivElem", "Niveau :", 6, 46, 40
    Set cmbNiveauElements = fraElements.Controls.Add("Forms.ComboBox.1", "cmbNiveauElements")
    cmbNiveauElements.Left = 52: cmbNiveauElements.Top = 44
    cmbNiveauElements.Width = 134: cmbNiveauElements.Height = 16

    CreerLabel fraElements, "lblCoulElem", "Couleur :", 6, 66, 42
    Set txtCouleurElements = fraElements.Controls.Add("Forms.TextBox.1", "txtCouleurElements")
    txtCouleurElements.Left = 56: txtCouleurElements.Top = 64
    txtCouleurElements.Width = 30: txtCouleurElements.Height = 16

    CreerLabel fraElements, "lblStyleElem", "Style ligne :", 6, 84, 56
    Set txtStyleElements = fraElements.Controls.Add("Forms.TextBox.1", "txtStyleElements")
    txtStyleElements.Left = 66: txtStyleElements.Top = 82
    txtStyleElements.Width = 60: txtStyleElements.Height = 16

    CreerLabel fraElements, "lblEpElem", "Epaisseur :", 6, 102, 48
    Set txtEpaisseurElements = fraElements.Controls.Add("Forms.TextBox.1", "txtEpaisseurElements")
    txtEpaisseurElements.Left = 60: txtEpaisseurElements.Top = 100
    txtEpaisseurElements.Width = 24: txtEpaisseurElements.Height = 16

    dY = dY + 128

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
    fraEtat.Width = 192: fraEtat.Height = 70

    Set lblEtat1 = CreerLabel(fraEtat, "lblEtat1", "-", 6, 12, 180)
    Set lblEtat2 = CreerLabel(fraEtat, "lblEtat2", "-", 6, 28, 180)
    Set lblEtat3 = CreerLabel(fraEtat, "lblEtat3", "", 6, 44, 180)
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
    RemplirNiveaux cmbNiveauTexte
    PositionnerNiveau cmbNiveauTexte, m_oSettings.oTexte.NomNiveau
    RemplirStyles cmbStyleTexte
    PositionnerStyle cmbStyleTexte, m_oSettings.oTexte.NomStyleTexte
    txtDecimales.Text = CStr(m_oSettings.oTexte.Decimales)
    ActiverChampsTexte

    ' Cercle repere
    chkCreerCercle.Value = m_oSettings.bCreerCercle
    txtDiametre.Text = Format$(m_oSettings.oCercle.Diametre, "0.00")
    txtCouleurCercle.Text = CStr(m_oSettings.oCercle.Couleur)
    RemplirNiveaux cmbNiveauCercle
    PositionnerNiveau cmbNiveauCercle, m_oSettings.oCercle.NomNiveau

    ' Decoupage
    chkDecoupage.Value = m_oSettings.bDecoupage
    optPasFixe.Value = (m_oSettings.oSemis.Mode = semisDistanceFixe)
    optPartsEgales.Value = (m_oSettings.oSemis.Mode = semisPartsEgales)
    optAuxSommets.Value = (m_oSettings.oSemis.Mode = semisAuxSommets)
    txtPas.Text = Format$(m_oSettings.oSemis.Pas, "0.00")
    txtNombre.Text = CStr(m_oSettings.oSemis.Nombre)
    optZInterpole.Value = Not m_oSettings.oSemis.bZDefini
    optZDefini.Value = m_oSettings.oSemis.bZDefini
    ActiverChampsDecoupage

    ' Convertir Cercle
    optArcNatif.Value = m_oSettings.bArcNatif
    optPolyligne.Value = Not m_oSettings.bArcNatif
    txtDiscret.Text = Format$(m_oSettings.dPasDiscretisation, "0.00")

    ' Symbologie elements 3D
    chkSymboSource.Value = m_oSettings.bSymboSource
    chkAttributsActifs.Value = m_oSettings.oElements.bActifs
    RemplirNiveaux cmbNiveauElements
    PositionnerNiveau cmbNiveauElements, m_oSettings.oElements.NomNiveau
    txtCouleurElements.Text = CStr(m_oSettings.oElements.Couleur)
    txtStyleElements.Text = m_oSettings.oElements.NomStyleLigne
    txtEpaisseurElements.Text = CStr(m_oSettings.oElements.Epaisseur)
    ActiverChampsElements

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

'------------------------------------------------------------------------------
Private Sub ActiverChampsDecoupage()
    Dim bActif As Boolean
    bActif = m_oSettings.bDecoupage
    optPasFixe.Enabled = bActif
    optPartsEgales.Enabled = bActif
    optAuxSommets.Enabled = bActif
    txtPas.Enabled = bActif And (m_oSettings.oSemis.Mode = semisDistanceFixe)
    txtNombre.Enabled = bActif And (m_oSettings.oSemis.Mode = semisPartsEgales)
    Dim bSommets As Boolean
    bSommets = bActif And (m_oSettings.oSemis.Mode = semisAuxSommets)
    optZInterpole.Enabled = bSommets
    optZDefini.Enabled = bSommets
End Sub

'------------------------------------------------------------------------------
' "Attributs actifs" n'a de sens que si on ne duplique pas la source ; les
' champs explicites ne sont saisissables que si "Attributs actifs" est decoche.
Private Sub ActiverChampsElements()
    Dim bSansSource As Boolean
    bSansSource = Not chkSymboSource.Value
    chkAttributsActifs.Enabled = bSansSource

    Dim bLibre As Boolean
    bLibre = bSansSource And Not m_oSettings.oElements.bActifs
    cmbNiveauElements.Enabled = bLibre
    txtCouleurElements.Enabled = bLibre
    txtStyleElements.Enabled = bLibre
    txtEpaisseurElements.Enabled = bLibre
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
' Ligne 3 du cadre Etat : liste des Z sommets (etape 5 - Z defini).
Sub AfficherSommets(sTexte As String)
    If Not m_bConstruit Then Exit Sub
    lblEtat3.Caption = sTexte
End Sub

'------------------------------------------------------------------------------
' Appele par les classes de commande au Reset (spec 5.9).
Sub ReinitialiserEtat()
    If Not m_bConstruit Then Exit Sub
    lblEtat1.Caption = "-"
    lblEtat2.Caption = "-"
    lblEtat3.Caption = ""
End Sub

'------------------------------------------------------------------------------
' Recale les champs texte apres lecture d'une cote source (mode duplication).
Sub RafraichirTexte()
    If m_oSettings Is Nothing Then Exit Sub
    m_bInit = True
    cmbNiveauTexte.Text = m_oSettings.oTexte.NomNiveau
    txtDecimales.Text = CStr(m_oSettings.oTexte.Decimales)
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

Private Sub optAuxSommets_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    If optAuxSommets.Value = True Then m_oSettings.oSemis.Mode = semisAuxSommets
    ActiverChampsDecoupage
End Sub

Private Sub optZInterpole_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    If optZInterpole.Value = True Then m_oSettings.oSemis.bZDefini = False
End Sub

Private Sub optZDefini_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    If optZDefini.Value = True Then m_oSettings.oSemis.bZDefini = True
End Sub

'==============================================================================
' Evenements Convertir Cercle
'==============================================================================

Private Sub optArcNatif_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    If optArcNatif.Value = True Then m_oSettings.bArcNatif = True
End Sub

Private Sub optPolyligne_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    If optPolyligne.Value = True Then m_oSettings.bArcNatif = False
End Sub

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

'==============================================================================
' Evenements Symbologie elements 3D
'==============================================================================

Private Sub chkSymboSource_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.bSymboSource = (chkSymboSource.Value = True)
    ActiverChampsElements
End Sub

Private Sub chkAttributsActifs_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oElements.bActifs = (chkAttributsActifs.Value = True)
    ActiverChampsElements
End Sub

Private Sub cmbNiveauElements_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oElements.NomNiveau = ExtraireNiveau(cmbNiveauElements.Text)
End Sub

Private Sub txtCouleurElements_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String: sVal = Trim$(txtCouleurElements.Text)
    If sVal = "" Then Exit Sub
    Dim nCoul As Long: nCoul = CLng(Val(sVal))
    If nCoul >= 0 And nCoul <= 255 Then m_oSettings.oElements.Couleur = nCoul
End Sub

Private Sub txtStyleElements_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oElements.NomStyleLigne = Trim$(txtStyleElements.Text)
End Sub

Private Sub txtEpaisseurElements_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String: sVal = Trim$(txtEpaisseurElements.Text)
    If sVal = "" Then Exit Sub
    Dim nEp As Long: nEp = CLng(Val(sVal))
    If nEp >= 0 And nEp <= 31 Then m_oSettings.oElements.Epaisseur = nEp
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

