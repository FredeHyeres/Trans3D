VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGraphique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CGraphique - Moteur graphique de Trans3D (seule classe qui cree des elements)
'
' Responsabilites :
'   - extraire la geometrie d'un element lineaire en tableaux de sommets,
'     exprimes dans le repere du modele actif (reference 2D ou modele 3D)
'   - creer l'element 3D de la commande Convertir (type conserve si possible,
'     sinon polyligne discretisee)
'   - creer le cercle repere (au Z reel) et le texte altitude (a Z=0.00)
'   - dessiner les apercus dynamiques (Redraw)
'
' Recoit des donnees deja calculees (CCalcul) et des parametres (CSettings).
' Les routines prennent des parametres generiques, jamais l'etat interne
' d'une commande.
'
' TYPES SUPPORTES en extraction :
'   ligne, ligne brisee, courbe (par ses sommets), arc, ellipse/cercle
'   (discretises au pas dPasDiscretisation), B-spline (tentative GetVertices
'   protegee - A VALIDER dans MicroStation, voir ExtraireBspline)
'==============================================================================
Option Explicit

Private Const PI As Double = 3.14159265358979
Private Const NB_SEG_MIN As Long = 8      ' segments mini pour un arc
Private Const NB_SEG_MAX As Long = 2000   ' garde-fou discretisation

'==============================================================================
' Extraction de geometrie
'==============================================================================

'------------------------------------------------------------------------------
' Extrait la polyligne d'un element 2D de reference, convertie dans le repere
' du modele actif (oAttachment = Nothing si l'element est du modele actif).
' Renvoie False si le type n'est pas supporte.
Function ExtrairePolyligne2D(oElem As Element, oAttachment As Object, _
                             ByVal dPasDiscret As Double, _
                             dX() As Double, dY() As Double) As Boolean
    Dim aPts() As Point3d
    Dim nNb As Long
    ExtrairePolyligne2D = ExtraireSommets(oElem, dPasDiscret, aPts, nNb)
    If Not ExtrairePolyligne2D Then Exit Function

    ReDim dX(0 To nNb - 1)
    ReDim dY(0 To nNb - 1)
    Dim i As Long, ptM As Point3d
    For i = 0 To nNb - 1
        TransformerPointVersMaitre oAttachment, aPts(i), ptM
        dX(i) = ptM.X
        dY(i) = ptM.Y
    Next
End Function

'------------------------------------------------------------------------------
' Extrait la polyligne 3D d'un element du modele actif (commande Semer).
' Renvoie False si le type n'est pas supporte.
Function ExtrairePolyligne3D(oElem As Element, ByVal dPasDiscret As Double, _
                             dX() As Double, dY() As Double, _
                             dZ() As Double) As Boolean
    Dim aPts() As Point3d
    Dim nNb As Long
    ExtrairePolyligne3D = ExtraireSommets(oElem, dPasDiscret, aPts, nNb)
    If Not ExtrairePolyligne3D Then Exit Function

    ReDim dX(0 To nNb - 1)
    ReDim dY(0 To nNb - 1)
    ReDim dZ(0 To nNb - 1)
    Dim i As Long
    For i = 0 To nNb - 1
        dX(i) = aPts(i).X
        dY(i) = aPts(i).Y
        dZ(i) = aPts(i).Z
    Next
End Function

'------------------------------------------------------------------------------
' Vrai si le type de l'element est convertible par Trans3D.
Function TypeSupporte(oElem As Element) As Boolean
    Select Case oElem.Type
        Case msdElementTypeLine, msdElementTypeLineString, _
             msdElementTypeCurve, msdElementTypeArc, _
             msdElementTypeEllipse, msdElementTypeBsplineCurve
            TypeSupporte = True
        Case Else
            TypeSupporte = False
    End Select
End Function

'------------------------------------------------------------------------------
' Vrai si l'element est ferme (cercle/ellipse complet) : une seule altitude
' de reference suffit (Z constant).
Function EstFerme(oElem As Element) As Boolean
    EstFerme = (oElem.Type = msdElementTypeEllipse)
End Function

'------------------------------------------------------------------------------
' Sommets bruts de l'element dans son propre repere.
Private Function ExtraireSommets(oElem As Element, ByVal dPasDiscret As Double, _
                                 aPts() As Point3d, nNb As Long) As Boolean
    ExtraireSommets = False
    nNb = 0

    Select Case oElem.Type
        Case msdElementTypeLine, msdElementTypeLineString
            ' LineElement couvre ligne (2 sommets) et ligne brisee
            Dim oLigne As LineElement
            Set oLigne = oElem
            aPts = oLigne.GetVertices
            nNb = UBound(aPts) - LBound(aPts) + 1
            NormaliserBase aPts, nNb
            ExtraireSommets = (nNb >= 2)

        Case msdElementTypeCurve
            ' Courbe (type 11) : passe par ses sommets, utilises tels quels
            Dim oCourbe As CurveElement
            Set oCourbe = oElem
            aPts = oCourbe.GetVertices
            nNb = UBound(aPts) - LBound(aPts) + 1
            NormaliserBase aPts, nNb
            ExtraireSommets = (nNb >= 2)

        Case msdElementTypeArc
            ExtraireSommets = DiscretiserArc(oElem, dPasDiscret, aPts, nNb)

        Case msdElementTypeEllipse
            ExtraireSommets = DiscretiserEllipse(oElem, dPasDiscret, aPts, nNb)

        Case msdElementTypeBsplineCurve
            ExtraireSommets = ExtraireBspline(oElem, aPts, nNb)
    End Select
End Function

'------------------------------------------------------------------------------
' Rebase un tableau de sommets en 0-based si necessaire.
Private Sub NormaliserBase(aPts() As Point3d, ByVal nNb As Long)
    If LBound(aPts) = 0 Then Exit Sub
    Dim aTmp() As Point3d
    ReDim aTmp(0 To nNb - 1)
    Dim i As Long
    For i = 0 To nNb - 1
        aTmp(i) = aPts(LBound(aPts) + i)
    Next
    aPts = aTmp
End Sub

'------------------------------------------------------------------------------
' Discretise un arc au pas donne (angles en radians, repere de l'arc).
Private Function DiscretiserArc(oElem As Element, ByVal dPasDiscret As Double, _
                                aPts() As Point3d, nNb As Long) As Boolean
    DiscretiserArc = False

    Dim oArc As ArcElement
    Set oArc = oElem

    Dim ptC As Point3d, dR1 As Double, dR2 As Double
    Dim dDebut As Double, dBalayage As Double
    Dim oRot As Matrix3d
    On Error GoTo Echec
    ptC = oArc.CenterPoint
    dR1 = oArc.PrimaryRadius
    dR2 = oArc.SecondaryRadius
    dDebut = oArc.StartAngle
    dBalayage = oArc.SweepAngle
    oRot = oArc.Rotation
    On Error GoTo 0

    GenererPointsElliptiques ptC, dR1, dR2, oRot, dDebut, dBalayage, _
                             dPasDiscret, aPts, nNb
    DiscretiserArc = (nNb >= 2)
    Exit Function
Echec:
    DiscretiserArc = False
End Function

'------------------------------------------------------------------------------
' Discretise une ellipse ou un cercle complet (balayage 2*PI, ferme).
Private Function DiscretiserEllipse(oElem As Element, _
                                    ByVal dPasDiscret As Double, _
                                    aPts() As Point3d, nNb As Long) As Boolean
    DiscretiserEllipse = False

    Dim oEll As EllipseElement
    Set oEll = oElem

    Dim ptC As Point3d, dR1 As Double, dR2 As Double
    Dim oRot As Matrix3d
    On Error GoTo Echec
    ' Le centre n'est pas expose sous le meme nom selon les versions :
    ' tentative CenterPoint puis Origin en liaison tardive.
    Dim oObj As Object
    Set oObj = oEll
    Dim bCentreOk As Boolean
    bCentreOk = False
    On Error Resume Next
    ptC = oObj.CenterPoint
    If Err.Number = 0 Then bCentreOk = True
    If Not bCentreOk Then
        Err.Clear
        ptC = oObj.Origin
        If Err.Number = 0 Then bCentreOk = True
    End If
    On Error GoTo Echec
    If Not bCentreOk Then Exit Function

    dR1 = oEll.PrimaryRadius
    dR2 = oEll.SecondaryRadius
    oRot = oEll.Rotation
    On Error GoTo 0

    GenererPointsElliptiques ptC, dR1, dR2, oRot, 0#, 2# * PI, _
                             dPasDiscret, aPts, nNb
    ' Refermer exactement sur le premier point
    If nNb >= 2 Then aPts(nNb - 1) = aPts(0)
    DiscretiserEllipse = (nNb >= 2)
    Exit Function
Echec:
    DiscretiserEllipse = False
End Function

'------------------------------------------------------------------------------
' Genere les points d'un arc elliptique : local (R1*cos t, R2*sin t, 0),
' monde = centre + Rotation * local (produit fait a la main sur les Row).
Private Sub GenererPointsElliptiques(ptC As Point3d, _
        ByVal dR1 As Double, ByVal dR2 As Double, oRot As Matrix3d, _
        ByVal dDebut As Double, ByVal dBalayage As Double, _
        ByVal dPasDiscret As Double, aPts() As Point3d, nNb As Long)

    ' Nombre de segments d'apres la longueur developpee approchee
    Dim dRMax As Double
    If dR1 > dR2 Then dRMax = dR1 Else dRMax = dR2
    Dim dLongueur As Double
    dLongueur = Abs(dBalayage) * dRMax

    Dim nSeg As Long
    If dPasDiscret > 0.0000001 Then
        nSeg = CLng(dLongueur / dPasDiscret) + 1
    Else
        nSeg = NB_SEG_MIN
    End If
    If nSeg < NB_SEG_MIN Then nSeg = NB_SEG_MIN
    If nSeg > NB_SEG_MAX Then nSeg = NB_SEG_MAX

    nNb = nSeg + 1
    ReDim aPts(0 To nNb - 1)

    Dim i As Long, dT As Double
    Dim vLoc As Point3d, vMonde As Point3d
    For i = 0 To nSeg
        dT = dDebut + dBalayage * CDbl(i) / CDbl(nSeg)
        vLoc.X = dR1 * Cos(dT)
        vLoc.Y = dR2 * Sin(dT)
        vLoc.Z = 0#
        vMonde.X = oRot.RowX.X * vLoc.X + oRot.RowX.Y * vLoc.Y + oRot.RowX.Z * vLoc.Z
        vMonde.Y = oRot.RowY.X * vLoc.X + oRot.RowY.Y * vLoc.Y + oRot.RowY.Z * vLoc.Z
        vMonde.Z = oRot.RowZ.X * vLoc.X + oRot.RowZ.Y * vLoc.Y + oRot.RowZ.Z * vLoc.Z
        aPts(i).X = ptC.X + vMonde.X
        aPts(i).Y = ptC.Y + vMonde.Y
        aPts(i).Z = ptC.Z + vMonde.Z
    Next
End Sub

'------------------------------------------------------------------------------
' B-spline : l'API V8i VBA n'expose pas d'evaluateur documente fiable, et la
' liaison tardive ne sait pas marshaler un tableau de Point3d (UDT).
' Echec propre pour l'instant -> l'etat de commande affichera "type non
' supporte : droppez la B-spline en ligne brisee puis reessayez".
' A EXPLORER sur poste dans MicroStation : BsplineCurve.FromElement puis
' evaluation, ou methodes de BsplineCurveElement visibles dans l'explorateur
' d'objets VBA (F2).
Private Function ExtraireBspline(oElem As Element, _
                                 aPts() As Point3d, nNb As Long) As Boolean
    ExtraireBspline = False
    nNb = 0
End Function

'==============================================================================
' Creation des elements 3D (commande Convertir)
'==============================================================================

'------------------------------------------------------------------------------
' Cree la polyligne 3D (ligne si 2 sommets, ligne brisee sinon) et l'ajoute
' au modele actif. oElemSource peut etre Nothing (symbologie active).
Function CreerPolyligne3D(dX() As Double, dY() As Double, dZ() As Double, _
                          oElemSource As Element, _
                          oSettings As CSettings) As Element
    Dim aPts() As Point3d
    ConstruirePoints dX, dY, dZ, aPts

    Dim oNouv As LineElement
    Set oNouv = CreateLineElement1(Nothing, aPts)
    AppliquerSymbologie oNouv, oElemSource, oSettings
    ActiveModelReference.AddElement oNouv
    Set CreerPolyligne3D = oNouv
End Function

'------------------------------------------------------------------------------
' Arc a Z constant : type natif conserve, pose a l'altitude dZ.
' Renvoie Nothing en cas d'echec (l'appelant discretise alors en polyligne).
Function CreerArc3DZConstant(oArcSource As Element, ByVal dZ As Double, _
                             oAttachment As Object, _
                             oSettings As CSettings) As Element
    Set CreerArc3DZConstant = Nothing

    Dim oArc As ArcElement
    Set oArc = oArcSource

    On Error GoTo Echec
    Dim ptC As Point3d, ptCM As Point3d
    ptC = oArc.CenterPoint
    TransformerPointVersMaitre oAttachment, ptC, ptCM
    ptCM.Z = dZ

    Dim dEchelle As Double
    dEchelle = EchelleAttachment(oAttachment)

    Dim oNouv As ArcElement
    Set oNouv = CreateArcElement2(Nothing, ptCM, _
                    oArc.PrimaryRadius * dEchelle, _
                    oArc.SecondaryRadius * dEchelle, _
                    oArc.Rotation, oArc.StartAngle, oArc.SweepAngle)
    AppliquerSymbologie oNouv, oArcSource, oSettings
    ActiveModelReference.AddElement oNouv
    Set CreerArc3DZConstant = oNouv
    Exit Function
Echec:
    Set CreerArc3DZConstant = Nothing
End Function

'------------------------------------------------------------------------------
' Cercle/ellipse a Z constant : type natif conserve, pose a l'altitude dZ.
' Renvoie Nothing en cas d'echec (l'appelant discretise alors en polyligne).
Function CreerCercle3DZConstant(oEllSource As Element, ByVal dZ As Double, _
                                oAttachment As Object, _
                                oSettings As CSettings) As Element
    Set CreerCercle3DZConstant = Nothing

    Dim oEll As EllipseElement
    Set oEll = oEllSource

    On Error GoTo Echec
    Dim oObj As Object
    Set oObj = oEll
    Dim ptC As Point3d
    Dim bCentreOk As Boolean
    bCentreOk = False
    On Error Resume Next
    ptC = oObj.CenterPoint
    If Err.Number = 0 Then bCentreOk = True
    If Not bCentreOk Then
        Err.Clear
        ptC = oObj.Origin
        If Err.Number = 0 Then bCentreOk = True
    End If
    On Error GoTo Echec
    If Not bCentreOk Then Exit Function

    Dim ptCM As Point3d
    TransformerPointVersMaitre oAttachment, ptC, ptCM
    ptCM.Z = dZ

    Dim dEchelle As Double
    dEchelle = EchelleAttachment(oAttachment)

    Dim oNouv As EllipseElement
    Set oNouv = CreateEllipseElement2(Nothing, ptCM, _
                    oEll.PrimaryRadius * dEchelle, _
                    oEll.SecondaryRadius * dEchelle, oEll.Rotation)
    AppliquerSymbologie oNouv, oEllSource, oSettings
    ActiveModelReference.AddElement oNouv
    Set CreerCercle3DZConstant = oNouv
    Exit Function
Echec:
    Set CreerCercle3DZConstant = Nothing
End Function

'==============================================================================
' Creation du point cote (commandes Semer et Points)
'==============================================================================

'------------------------------------------------------------------------------
' Cree, selon les bascules de CSettings :
'   - le cercle repere centre en (dXpt, dYpt) au Z reel dZpt
'   - le texte altitude a Z = 0.00 (regle fixe du cahier des charges),
'     decale du rayon du cercle pour rester lisible
Sub CreerPointCote(ByVal dXpt As Double, ByVal dYpt As Double, _
                   ByVal dZpt As Double, _
                   oSettings As CSettings, oCalc As CCalcul)

    If oSettings.bCreerCercle Then
        Dim ptCercle As Point3d
        ptCercle.X = dXpt: ptCercle.Y = dYpt: ptCercle.Z = dZpt
        Dim oCercle As EllipseElement
        Set oCercle = CreateEllipseElement2(Nothing, ptCercle, _
                          oSettings.oCercle.Rayon, oSettings.oCercle.Rayon, _
                          Matrix3dIdentity())
        oCercle.Color = oSettings.oCercle.Couleur
        Set oCercle.Level = ResolverNiveau(oSettings.oCercle.NomNiveau)
        ActiveModelReference.AddElement oCercle
    End If

    If oSettings.bCreerTexte Then
        Dim ptTexte As Point3d
        ptTexte.X = dXpt + oSettings.oCercle.Rayon * 1.5
        ptTexte.Y = dYpt
        ptTexte.Z = 0#          ' regle fixe : texte a l'altitude 0.00
        Dim oTxt As TextElement
        Set oTxt = ConstruireTexteAltitude(ptTexte, dZpt, oSettings, oCalc)
        If Not oTxt Is Nothing Then ActiveModelReference.AddElement oTxt
    End If
End Sub

'==============================================================================
' Apercus dynamiques (appeles dans IPrimitiveCommandEvents_Dynamics)
'==============================================================================

'------------------------------------------------------------------------------
' Apercu de la polyligne 3D en cours de construction (Convertir).
Sub DessinPolyligneDyn(ByVal DrawMode As MsdDrawingMode, _
                       dX() As Double, dY() As Double, dZ() As Double)
    Dim aPts() As Point3d
    ConstruirePoints dX, dY, dZ, aPts
    On Error Resume Next
    Dim oLigne As LineElement
    Set oLigne = CreateLineElement1(Nothing, aPts)
    oLigne.Color = 4
    oLigne.Redraw DrawMode
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Surlignage a plat (Z=0) de l'element 2D selectionne.
Sub DessinSurlignage2D(ByVal DrawMode As MsdDrawingMode, _
                       dX() As Double, dY() As Double)
    Dim nLo As Long, nHi As Long, i As Long
    nLo = LBound(dX): nHi = UBound(dX)
    Dim aPts() As Point3d
    ReDim aPts(nLo To nHi)
    For i = nLo To nHi
        aPts(i).X = dX(i)
        aPts(i).Y = dY(i)
        aPts(i).Z = 0#
    Next
    On Error Resume Next
    Dim oLigne As LineElement
    Set oLigne = CreateLineElement1(Nothing, aPts)
    oLigne.Color = 3
    oLigne.Redraw DrawMode
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Apercu d'un point cote : cercle au Z reel + texte a Z=0.
Sub DessinPointCoteDyn(ByVal DrawMode As MsdDrawingMode, _
                       ByVal dXpt As Double, ByVal dYpt As Double, _
                       ByVal dZpt As Double, _
                       oSettings As CSettings, oCalc As CCalcul)
    On Error Resume Next

    If oSettings.bCreerCercle Then
        Dim ptCercle As Point3d
        ptCercle.X = dXpt: ptCercle.Y = dYpt: ptCercle.Z = dZpt
        Dim oCercle As EllipseElement
        Set oCercle = CreateEllipseElement2(Nothing, ptCercle, _
                          oSettings.oCercle.Rayon, oSettings.oCercle.Rayon, _
                          Matrix3dIdentity())
        oCercle.Color = oSettings.oCercle.Couleur
        oCercle.Redraw DrawMode
    End If

    If oSettings.bCreerTexte Then
        Dim ptTexte As Point3d
        ptTexte.X = dXpt + oSettings.oCercle.Rayon * 1.5
        ptTexte.Y = dYpt
        ptTexte.Z = 0#
        Dim oTxt As TextElement
        Set oTxt = ConstruireTexteAltitude(ptTexte, dZpt, oSettings, oCalc)
        If Not oTxt Is Nothing Then oTxt.Redraw DrawMode
    End If

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Croix de reperage (reprise d'InterpolationTopo).
Sub DessinCroix(ByVal DrawMode As MsdDrawingMode, oPt As Point3d, _
                ByVal dTaille As Double)
    Dim oA As Point3d, oB As Point3d, oL As LineElement

    oA.X = oPt.X - dTaille: oA.Y = oPt.Y: oA.Z = oPt.Z
    oB.X = oPt.X + dTaille: oB.Y = oPt.Y: oB.Z = oPt.Z
    Set oL = CreateLineElement2(Nothing, oA, oB)
    oL.Color = 3: oL.Redraw DrawMode

    oA.X = oPt.X: oA.Y = oPt.Y - dTaille
    oB.X = oPt.X: oB.Y = oPt.Y + dTaille
    Set oL = CreateLineElement2(Nothing, oA, oB)
    oL.Color = 3: oL.Redraw DrawMode
End Sub

'==============================================================================
' Utilitaires prives
'==============================================================================

'------------------------------------------------------------------------------
' Construit le texte altitude (sans l'ajouter au modele) : mode duplication
' (texte-modele source) ou mode creation (parametres du formulaire).
Private Function ConstruireTexteAltitude(ptTexte As Point3d, _
        ByVal dZValeur As Double, _
        oSettings As CSettings, oCalc As CCalcul) As TextElement

    Set ConstruireTexteAltitude = Nothing

    Dim sAlt As String
    sAlt = oCalc.FormatAltitude(dZValeur, oSettings.oTexte.Decimales, _
                                oSettings.oTexte.SepDecimal)

    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    If oSettings.oTexte.CommeModele And Not oSettings.oTextModele Is Nothing Then
        oRot = oSettings.oTextModele.Rotation
    End If
    On Error GoTo 0

    Dim oTxt As TextElement
    Dim bSansModele As Boolean
    bSansModele = True

    On Error Resume Next
    If oSettings.oTexte.CommeModele And Not oSettings.oTextModele Is Nothing Then
        ' PIEGE 5.3 : toujours creer AVEC l'element modele, jamais recopier
        ' TextStyle.Height d'un texte de reference (unites differentes).
        Set oTxt = CreateTextElement1(oSettings.oTextModele, sAlt, ptTexte, oRot)
        If Not oTxt Is Nothing Then bSansModele = False
        Err.Clear
    End If
    If oTxt Is Nothing Then
        Set oTxt = CreateTextElement1(Nothing, sAlt, ptTexte, oRot)
    End If
    On Error GoTo 0
    If oTxt Is Nothing Then Exit Function

    If bSansModele Or Not oSettings.oTexte.CommeModele Then
        ' Mode creation : appliquer les parametres du formulaire
        On Error Resume Next
        Dim oTS As TextStyle
        Set oTS = oTxt.TextStyle
        If oSettings.oTexte.Hauteur > 0 Then oTS.Height = oSettings.oTexte.Hauteur
        If oSettings.oTexte.Largeur > 0 Then
            oTS.Width = oSettings.oTexte.Largeur
        ElseIf oSettings.oTexte.Hauteur > 0 Then
            oTS.Width = oSettings.oTexte.Hauteur
        End If
        If Len(oSettings.oTexte.NomPolice) > 0 Then
            Dim oFont As Object
            Set oFont = FindFont(oSettings.oTexte.NomPolice)
            If Not oFont Is Nothing Then Set oTS.Font = oFont
            Err.Clear
        End If
        If oSettings.oTexte.FondActif Then
            ' Fond/masque : nom exact de la propriete a valider sur poste,
            ' liaison tardive pour ne pas bloquer la compilation
            Dim oTSObj As Object
            Set oTSObj = oTS
            oTSObj.Background = True
            Err.Clear
        End If
        oTxt.TextStyle = oTS
        oTxt.Color = oSettings.oTexte.Couleur
        Set oTxt.Level = ResolverNiveau(oSettings.oTexte.NomNiveau)
        Err.Clear
        On Error GoTo 0
    End If

    Set ConstruireTexteAltitude = oTxt
End Function

'------------------------------------------------------------------------------
' Cherche une police par nom. Tout en liaison tardive : l'emplacement exact de
' la collection des polices varie selon les versions (a valider sur poste).
Private Function FindFont(sNom As String) As Object
    Set FindFont = Nothing
    On Error Resume Next
    Dim oApp As Object
    Set oApp = Application
    Dim oFonts As Object
    Set oFonts = oApp.ActiveDesignFile.Fonts
    If oFonts Is Nothing Then
        Err.Clear
        Set oFonts = oApp.Fonts
    End If
    If Not oFonts Is Nothing Then
        Set FindFont = oFonts.Item(sNom)
    End If
    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Applique niveau/couleur/style/epaisseur de l'element source (mode
' bSymboSource), le niveau etant cree dans le fichier actif s'il n'existe
' pas. Si bSymboSource est faux ou la source Nothing : attributs actifs
' (rien a faire, CreateXxx sans modele les applique deja).
Private Sub AppliquerSymbologie(oDst As Element, oSrc As Element, _
                                oSettings As CSettings)
    If oSrc Is Nothing Then Exit Sub
    If Not oSettings.bSymboSource Then Exit Sub

    On Error Resume Next
    oDst.Color = oSrc.Color: Err.Clear
    oDst.LineWeight = oSrc.LineWeight: Err.Clear
    Set oDst.LineStyle = oSrc.LineStyle: Err.Clear

    Dim sNiveau As String
    sNiveau = oSrc.Level.Name
    If Err.Number = 0 And Len(sNiveau) > 0 Then
        Set oDst.Level = NiveauOuCreation(sNiveau)
    End If
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Niveau par nom dans le fichier actif ; cree s'il n'existe pas (la source
' vient d'une reference dont les niveaux ne sont pas forcement presents).
Private Function NiveauOuCreation(sNiveau As String) As Level
    Set NiveauOuCreation = ActiveSettings.Level

    On Error Resume Next
    Dim oLvl As Level
    Set oLvl = ActiveDesignFile.Levels(sNiveau)
    If oLvl Is Nothing Then
        Err.Clear
        Set oLvl = ActiveDesignFile.AddNewLevel(sNiveau)
    End If
    Err.Clear
    On Error GoTo 0

    If Not oLvl Is Nothing Then Set NiveauOuCreation = oLvl
End Function

'------------------------------------------------------------------------------
' Niveau par nom ou numero ; chaine vide ou introuvable = niveau actif.
Private Function ResolverNiveau(sNiveau As String) As Level
    Set ResolverNiveau = ActiveSettings.Level
    If sNiveau = "" Then Exit Function

    Dim oLvl As Level
    On Error Resume Next
    Set oLvl = ActiveDesignFile.Levels(sNiveau)
    On Error GoTo 0
    If oLvl Is Nothing And IsNumeric(sNiveau) Then
        Dim oL As Level
        For Each oL In ActiveDesignFile.Levels
            If oL.Number = CLng(Val(sNiveau)) Then
                Set oLvl = oL: Exit For
            End If
        Next
    End If
    If Not oLvl Is Nothing Then Set ResolverNiveau = oLvl
End Function

'------------------------------------------------------------------------------
' Facteur d'echelle d'un attachement (1.0 si Nothing ou illisible).
Private Function EchelleAttachment(oAttachment As Object) As Double
    EchelleAttachment = 1#
    If oAttachment Is Nothing Then Exit Function
    On Error Resume Next
    Dim dS As Double
    dS = CDbl(oAttachment.ScaleFactor)
    If Err.Number = 0 And Abs(dS) > 0.0000000001 Then EchelleAttachment = dS
    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Assemble trois tableaux de coordonnees en tableau de Point3d.
Private Sub ConstruirePoints(dX() As Double, dY() As Double, dZ() As Double, _
                             aPts() As Point3d)
    Dim nLo As Long, nHi As Long, i As Long
    nLo = LBound(dX): nHi = UBound(dX)
    ReDim aPts(nLo To nHi)
    For i = nLo To nHi
        aPts(i).X = dX(i)
        aPts(i).Y = dY(i)
        aPts(i).Z = dZ(i)
    Next
End Sub
