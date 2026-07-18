Attribute VB_Name = "RechercheElement"
'==============================================================================
' RechercheElement - Recherche de l'element lineaire le plus proche d'un clic
'
' Pendant du module RechercheAltitude pour les elements convertibles :
' scanne le modele actif et, sur demande, les references attachees affichees.
' La distance clic-element est mesuree en plan sur la polyligne extraite par
' CGraphique (donc deja convertie dans le repere du modele actif).
'
' Un prefiltre sur le Range de l'element evite d'extraire la geometrie des
' elements manifestement hors du rayon de recherche (performance sur les
' gros plans).
'
' Fonction publique :
'   TrouverElementProche : renvoie l'element et son attachement (Nothing si
'                          l'element est du modele actif)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Cherche l'element lineaire supporte le plus proche du clic dans le rayon
' donne. bAvecReferences : True = scanner aussi les references (Convertir),
' False = modele actif seul (Semer). Renvoie True si un element est trouve.
Public Function TrouverElementProche(oPtClic As Point3d, ByVal dRayon As Double, _
        ByVal bAvecReferences As Boolean, _
        oElemOut As Element, oAttOut As Object) As Boolean

    Set oElemOut = Nothing
    Set oAttOut = Nothing

    Dim oScan As New ElementScanCriteria
    oScan.ExcludeAllTypes
    oScan.IncludeType msdElementTypeLine
    oScan.IncludeType msdElementTypeLineString
    oScan.IncludeType msdElementTypeCurve
    oScan.IncludeType msdElementTypeArc
    oScan.IncludeType msdElementTypeEllipse
    oScan.IncludeType msdElementTypeBsplineCurve

    Dim dMinDist As Double
    dMinDist = dRayon

    ScannerElementsDansModele ActiveModelReference, Nothing, oScan, oPtClic, _
                              dRayon, dMinDist, oElemOut, oAttOut

    If bAvecReferences Then
        Dim oAtt As Object
        On Error Resume Next
        For Each oAtt In ActiveModelReference.Attachments
            If AttachmentAffiche(oAtt) Then
                ScannerElementsDansModele oAtt, oAtt, oScan, oPtClic, _
                                          dRayon, dMinDist, oElemOut, oAttOut
            End If
        Next
        On Error GoTo 0
    End If

    TrouverElementProche = Not (oElemOut Is Nothing)
End Function

'------------------------------------------------------------------------------
Private Sub ScannerElementsDansModele(oModel As Object, oAttachment As Object, _
        oScan As ElementScanCriteria, oPtClic As Point3d, _
        ByVal dRayon As Double, dMinDist As Double, _
        oElemOut As Element, oAttOut As Object)

    Dim oEnum As ElementEnumerator
    On Error Resume Next
    Set oEnum = oModel.Scan(oScan)
    If Err.Number <> 0 Or oEnum Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Do While oEnum.MoveNext
        Dim oElem As Element
        Set oElem = oEnum.Current

        If Not HorsZone(oElem, oAttachment, oPtClic, dRayon) Then
            Dim dD As Double
            dD = DistanceElement(oElem, oAttachment, oPtClic)
            If dD >= 0# And dD < dMinDist Then
                dMinDist = dD
                Set oElemOut = oElem
                Set oAttOut = oAttachment
            End If
        End If
    Loop
End Sub

'------------------------------------------------------------------------------
' Distance en plan entre le clic et l'element (repere du modele actif).
' Renvoie -1 si la geometrie n'est pas extractible (ex. B-spline).
Private Function DistanceElement(oElem As Element, oAttachment As Object, _
                                 oPtClic As Point3d) As Double
    DistanceElement = -1#

    Dim dX() As Double, dY() As Double, dS() As Double
    If Not g_oMoteur.ExtrairePolyligne2D(oElem, oAttachment, _
            g_oSettings.dPasDiscretisation, dX, dY) Then Exit Function

    g_oCalc.AbscissesPolyligne2D dX, dY, dS

    Dim dDist As Double
    g_oCalc.AbscisseProjetee2D dX, dY, dS, oPtClic.X, oPtClic.Y, dDist
    DistanceElement = dDist
End Function

'------------------------------------------------------------------------------
' Prefiltre : True si le Range de l'element est surement hors du rayon.
' Le clic est ramene dans le repere de la reference pour la comparaison.
' En cas de doute (erreur, echelle illisible) : False, l'element est examine.
Private Function HorsZone(oElem As Element, oAttachment As Object, _
                          oPtClic As Point3d, ByVal dRayon As Double) As Boolean
    HorsZone = False

    On Error GoTo Inconnu
    Dim oRng As Range3d
    oRng = oElem.Range

    Dim ptLocal As Point3d
    TransformerPointVersRef oAttachment, oPtClic, ptLocal

    Dim dRay As Double
    dRay = dRayon / EchelleAtt(oAttachment)

    If ptLocal.X < oRng.Low.X - dRay Then HorsZone = True: Exit Function
    If ptLocal.X > oRng.High.X + dRay Then HorsZone = True: Exit Function
    If ptLocal.Y < oRng.Low.Y - dRay Then HorsZone = True: Exit Function
    If ptLocal.Y > oRng.High.Y + dRay Then HorsZone = True: Exit Function
    Exit Function
Inconnu:
    HorsZone = False
End Function

'------------------------------------------------------------------------------
' Inverse de TransformerPointVersMaitre : ramene un point du repere du modele
' actif dans le repere de la reference (origines + echelle, comme l'aller).
Private Sub TransformerPointVersRef(oAttachment As Object, _
        oPtMaster As Point3d, oPtRef As Point3d)

    oPtRef = oPtMaster
    If oAttachment Is Nothing Then Exit Sub

    Dim ptRefOrigin As Point3d
    Dim ptMasterOrigin As Point3d
    Dim dScale As Double

    dScale = 1#
    ptRefOrigin.X = 0#: ptRefOrigin.Y = 0#: ptRefOrigin.Z = 0#
    ptMasterOrigin.X = 0#: ptMasterOrigin.Y = 0#: ptMasterOrigin.Z = 0#

    On Error Resume Next
    ptRefOrigin = oAttachment.ReferenceOrigin
    If Err.Number <> 0 Then
        Err.Clear
        ptRefOrigin.X = 0#: ptRefOrigin.Y = 0#: ptRefOrigin.Z = 0#
    End If

    ptMasterOrigin = oAttachment.MasterOrigin
    If Err.Number <> 0 Then
        Err.Clear
        ptMasterOrigin.X = 0#: ptMasterOrigin.Y = 0#: ptMasterOrigin.Z = 0#
    End If

    dScale = CDbl(oAttachment.ScaleFactor)
    If Err.Number <> 0 Or Abs(dScale) < 0.0000000001 Then
        Err.Clear
        dScale = 1#
    End If
    On Error GoTo 0

    oPtRef.X = ptRefOrigin.X + (oPtMaster.X - ptMasterOrigin.X) / dScale
    oPtRef.Y = ptRefOrigin.Y + (oPtMaster.Y - ptMasterOrigin.Y) / dScale
    oPtRef.Z = ptRefOrigin.Z + (oPtMaster.Z - ptMasterOrigin.Z) / dScale
End Sub

'------------------------------------------------------------------------------
Private Function EchelleAtt(oAttachment As Object) As Double
    EchelleAtt = 1#
    If oAttachment Is Nothing Then Exit Function
    On Error Resume Next
    Dim dS As Double
    dS = CDbl(oAttachment.ScaleFactor)
    If Err.Number = 0 And Abs(dS) > 0.0000000001 Then EchelleAtt = dS
    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
Private Function AttachmentAffiche(oAttachment As Object) As Boolean
    AttachmentAffiche = True
    On Error Resume Next
    AttachmentAffiche = CBool(oAttachment.IsDisplayed)
    If Err.Number <> 0 Then
        Err.Clear
        AttachmentAffiche = True
    End If
    On Error GoTo 0
End Function
