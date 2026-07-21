VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCalcul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CCalcul - Moteur de calcul pur de Trans3D
'
' Issu de CInterpolation (InterpolationTopo), elague de la projection P1-P2
' et enrichi du calcul curviligne pour la conversion 2D -> 3D.
'
' Responsabilites :
'   - parsing d'un texte altitude (numerique, separateur, decimales)
'   - abscisses curvilignes d'une polyligne (2D et 3D)
'   - point / tangente a une abscisse donnee sur une polyligne
'   - projection d'un point sur une polyligne (abscisse de la cote cliquee)
'   - semis d'abscisses : pas fixe ou parts egales
'   - interpolation d'altitude par troncons (refs depart/fin/intermediaires)
'   - pente, formatage
'
' Conventions :
'   - toutes les polylignes sont des tableaux Double ; les tableaux de sortie
'     sont ReDim en 0-based par cette classe
'   - les abscisses de reference (dSRef/dZRef) sont triees par
'     TrierRefsParAbscisse avant toute interpolation
'
' Cette classe ne connait pas : symbologie, niveaux, elements MST, formulaire.
' Seule exception heritee : LireAltitudeTexte lit g_oSelectionCourante
' (resultat de RechercheAltitude) et le type Point3d (UDT simple).
'==============================================================================
Option Explicit

Private Const PI As Double = 3.14159265358979
Private Const EPSILON As Double = 0.0000000001
Private Const DIST_MIN As Double = 0.001    ' points confondus : ~1 mm

'==============================================================================
' Parsing
'==============================================================================

'------------------------------------------------------------------------------
' Vrai si s est un nombre au format normalise [+-]chiffres[.chiffres]
Function EstNombre(ByVal s As String) As Boolean
    EstNombre = False
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "+" Or Left$(s, 1) = "-" Then s = Mid$(s, 2)
    Dim i As Integer, nDots As Integer, nDigits As Integer, c As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = "." Then
            nDots = nDots + 1
            If nDots > 1 Then Exit Function
        ElseIf c >= "0" And c <= "9" Then
            nDigits = nDigits + 1
        Else
            Exit Function
        End If
    Next
    EstNombre = (nDigits > 0)
End Function

'------------------------------------------------------------------------------
' Lit l'altitude de la selection courante et charge le CPointRef.
' Renseigne aussi sSep ("." ou ",") et nDec (nb de decimales).
' Retourne False si le texte n'est pas numerique.
Function LireAltitudeTexte(oPoint As CPointRef, _
                            sSep As String, nDec As Integer) As Boolean
    LireAltitudeTexte = False
    oPoint.Invalider

    Dim sBrut As String
    sBrut = g_oSelectionCourante.ValeurTexte
    If InStr(sBrut, ",") > 0 Then
        sSep = ","
        sBrut = Replace(sBrut, ",", ".")
    Else
        sSep = "."
    End If

    If Not EstNombre(sBrut) Then Exit Function

    oPoint.Altitude = Val(sBrut)
    Dim ptOrigine As Point3d
    g_oSelectionCourante.CopierOrigine ptOrigine
    oPoint.DefinirPosition ptOrigine
    oPoint.Valide = True

    Dim nDot As Integer
    nDot = InStrRev(sBrut, ".")
    If nDot > 0 Then nDec = Len(sBrut) - nDot Else nDec = 0

    LireAltitudeTexte = True
End Function

'==============================================================================
' Geometrie de base
'==============================================================================

'------------------------------------------------------------------------------
Function Dist2D(oA As Point3d, oB As Point3d) As Double
    Dist2D = Sqr((oB.X - oA.X) ^ 2 + (oB.Y - oA.Y) ^ 2)
End Function

'------------------------------------------------------------------------------
' Angle en radians depuis l'axe X (sens trigonometrique) pour un vecteur.
Function Angle2D(dDX As Double, dDY As Double) As Double
    If Abs(dDX) < EPSILON And Abs(dDY) < EPSILON Then
        Angle2D = 0#
        Exit Function
    End If
    Angle2D = Atn2Priv(dDY, dDX)
End Function

Private Function Atn2Priv(ByVal dY As Double, ByVal dX As Double) As Double
    If dX > 0 Then
        Atn2Priv = Atn(dY / dX)
    ElseIf dX < 0 Then
        Atn2Priv = Atn(dY / dX) + IIf(dY >= 0, PI, -PI)
    Else
        Atn2Priv = IIf(dY >= 0, PI / 2#, -PI / 2#)
    End If
End Function

'==============================================================================
' Polylignes - abscisses curvilignes
'
' Une polyligne est decrite par ses sommets dX()/dY() (et dZ() en 3D).
' dS() recoit l'abscisse curviligne de chaque sommet (0 au premier).
' La longueur totale developpee est renvoyee.
'==============================================================================

'------------------------------------------------------------------------------
' Abscisses en plan (source 2D : la longueur est mesuree en XY).
Function AbscissesPolyligne2D(dX() As Double, dY() As Double, _
                              dS() As Double) As Double
    Dim nLo As Long, nHi As Long, i As Long
    nLo = LBound(dX): nHi = UBound(dX)
    ReDim dS(nLo To nHi)
    dS(nLo) = 0#
    For i = nLo + 1 To nHi
        dS(i) = dS(i - 1) + Sqr((dX(i) - dX(i - 1)) ^ 2 _
                              + (dY(i) - dY(i - 1)) ^ 2)
    Next
    AbscissesPolyligne2D = dS(nHi)
End Function

'------------------------------------------------------------------------------
' Abscisses sur la geometrie 3D reelle (semis sur element 3D).
Function AbscissesPolyligne3D(dX() As Double, dY() As Double, dZ() As Double, _
                              dS() As Double) As Double
    Dim nLo As Long, nHi As Long, i As Long
    nLo = LBound(dX): nHi = UBound(dX)
    ReDim dS(nLo To nHi)
    dS(nLo) = 0#
    For i = nLo + 1 To nHi
        dS(i) = dS(i - 1) + Sqr((dX(i) - dX(i - 1)) ^ 2 _
                              + (dY(i) - dY(i - 1)) ^ 2 _
                              + (dZ(i) - dZ(i - 1)) ^ 2)
    Next
    AbscissesPolyligne3D = dS(nHi)
End Function

'------------------------------------------------------------------------------
' Point a l'abscisse dAbs sur une polyligne 2D (abscisse bornee a [0, L]).
Sub PointSurPolyligne2D(dX() As Double, dY() As Double, dS() As Double, _
                        ByVal dAbs As Double, _
                        dXOut As Double, dYOut As Double)
    Dim i As Long, dT As Double
    i = TrouverSegment(dS, dAbs, dT)
    dXOut = dX(i - 1) + dT * (dX(i) - dX(i - 1))
    dYOut = dY(i - 1) + dT * (dY(i) - dY(i - 1))
End Sub

'------------------------------------------------------------------------------
' Point a l'abscisse dAbs sur une polyligne 3D.
Sub PointSurPolyligne3D(dX() As Double, dY() As Double, dZ() As Double, _
                        dS() As Double, ByVal dAbs As Double, _
                        dXOut As Double, dYOut As Double, dZOut As Double)
    Dim i As Long, dT As Double
    i = TrouverSegment(dS, dAbs, dT)
    dXOut = dX(i - 1) + dT * (dX(i) - dX(i - 1))
    dYOut = dY(i - 1) + dT * (dY(i) - dY(i - 1))
    dZOut = dZ(i - 1) + dT * (dZ(i) - dZ(i - 1))
End Sub

'------------------------------------------------------------------------------
' Direction unitaire (en plan) du segment porteur de l'abscisse dAbs.
' Sert a orienter un texte le long de l'element. Segment degenere : (1, 0).
Sub TangenteSurPolyligne(dX() As Double, dY() As Double, dS() As Double, _
                         ByVal dAbs As Double, _
                         dDirX As Double, dDirY As Double)
    Dim i As Long, dT As Double, dLen As Double
    i = TrouverSegment(dS, dAbs, dT)
    dDirX = dX(i) - dX(i - 1)
    dDirY = dY(i) - dY(i - 1)
    dLen = Sqr(dDirX * dDirX + dDirY * dDirY)
    If dLen < EPSILON Then
        dDirX = 1#: dDirY = 0#
    Else
        dDirX = dDirX / dLen: dDirY = dDirY / dLen
    End If
End Sub

'------------------------------------------------------------------------------
' Abscisse curviligne du point de la polyligne le plus proche de (dPtX, dPtY).
' Utilise pour placer une altitude intermediaire : la cote cliquee est
' projetee sur l'element pour connaitre son abscisse.
' dDistOut recoit la distance en plan entre le point et sa projection.
Function AbscisseProjetee2D(dX() As Double, dY() As Double, dS() As Double, _
                            ByVal dPtX As Double, ByVal dPtY As Double, _
                            Optional dDistOut As Double) As Double
    Dim nLo As Long, nHi As Long, i As Long
    nLo = LBound(dX): nHi = UBound(dX)

    Dim dBestDist As Double, dBestAbs As Double
    dBestDist = -1#

    For i = nLo + 1 To nHi
        Dim dvx As Double, dvy As Double, dL2 As Double, dT As Double
        dvx = dX(i) - dX(i - 1)
        dvy = dY(i) - dY(i - 1)
        dL2 = dvx * dvx + dvy * dvy
        If dL2 < EPSILON Then
            dT = 0#
        Else
            dT = ((dPtX - dX(i - 1)) * dvx + (dPtY - dY(i - 1)) * dvy) / dL2
            If dT < 0# Then dT = 0#
            If dT > 1# Then dT = 1#
        End If
        Dim dPx As Double, dPy As Double, dD As Double
        dPx = dX(i - 1) + dT * dvx
        dPy = dY(i - 1) + dT * dvy
        dD = Sqr((dPtX - dPx) ^ 2 + (dPtY - dPy) ^ 2)
        If dBestDist < 0# Or dD < dBestDist Then
            dBestDist = dD
            dBestAbs = dS(i - 1) + dT * (dS(i) - dS(i - 1))
        End If
    Next

    dDistOut = dBestDist
    AbscisseProjetee2D = dBestAbs
End Function

'------------------------------------------------------------------------------
' Trouve le segment porteur de l'abscisse dAbs : renvoie l'indice i du sommet
' superieur (le segment est [i-1, i]) et dT la position parametrique 0..1.
' L'abscisse est bornee a [0, longueur totale].
Private Function TrouverSegment(dS() As Double, ByVal dAbs As Double, _
                                dT As Double) As Long
    Dim nLo As Long, nHi As Long, i As Long
    nLo = LBound(dS): nHi = UBound(dS)

    If nHi = nLo Then
        ' Polyligne reduite a un point : segment fictif sur ce point
        TrouverSegment = nHi
        dT = 0#
        Exit Function
    End If

    If dAbs <= dS(nLo) Then
        TrouverSegment = nLo + 1
        dT = 0#
        Exit Function
    End If
    If dAbs >= dS(nHi) Then
        TrouverSegment = nHi
        dT = 1#
        Exit Function
    End If

    For i = nLo + 1 To nHi
        If dAbs <= dS(i) Then
            Dim dLen As Double
            dLen = dS(i) - dS(i - 1)
            If dLen < EPSILON Then dT = 0# Else dT = (dAbs - dS(i - 1)) / dLen
            TrouverSegment = i
            Exit Function
        End If
    Next
    TrouverSegment = nHi
    dT = 1#
End Function

'==============================================================================
' Semis d'abscisses
'==============================================================================

'------------------------------------------------------------------------------
' Mode distance fixe : un point au depart (0), puis tous les dPas, et un point
' a la fin (dLongueur) meme si le dernier intervalle est plus court.
' Renvoie le nombre de points ; dSOut est ReDim 0-based.
Function AbscissesSemisPasFixe(ByVal dLongueur As Double, ByVal dPas As Double, _
                               dSOut() As Double) As Long
    If dLongueur < DIST_MIN Or dPas < DIST_MIN Then
        ReDim dSOut(0 To 0)
        dSOut(0) = 0#
        AbscissesSemisPasFixe = 1
        Exit Function
    End If

    Dim nInter As Long
    nInter = Int(dLongueur / dPas)

    ' Si le reliquat final est negligeable, le dernier multiple EST la fin
    Dim bFinSurPas As Boolean
    bFinSurPas = ((dLongueur - nInter * dPas) < DIST_MIN)

    Dim nNb As Long
    If bFinSurPas Then nNb = nInter + 1 Else nNb = nInter + 2
    ReDim dSOut(0 To nNb - 1)

    Dim i As Long
    For i = 0 To nInter
        dSOut(i) = i * dPas
    Next
    dSOut(nNb - 1) = dLongueur

    AbscissesSemisPasFixe = nNb
End Function

'------------------------------------------------------------------------------
' Mode parts egales : nNombre points repartis uniformement, depart et fin
' inclus (minimum 2). Renvoie le nombre de points ; dSOut est ReDim 0-based.
Function AbscissesSemisPartsEgales(ByVal dLongueur As Double, _
                                   ByVal nNombre As Long, _
                                   dSOut() As Double) As Long
    Dim nNb As Long
    nNb = nNombre
    If nNb < 2 Then nNb = 2
    ReDim dSOut(0 To nNb - 1)

    Dim dPas As Double
    dPas = dLongueur / (nNb - 1)
    Dim i As Long
    For i = 0 To nNb - 1
        dSOut(i) = i * dPas
    Next
    dSOut(nNb - 1) = dLongueur

    AbscissesSemisPartsEgales = nNb
End Function

'==============================================================================
' Altimetrie - interpolation par troncons
'
' Les references (depart, fin, intermediaires) sont des couples
' (abscisse dSRef, altitude dZRef) tries par abscisse croissante.
'==============================================================================

'------------------------------------------------------------------------------
Function InterpolerZ(dZ1 As Double, dZ2 As Double, dT As Double) As Double
    InterpolerZ = dZ1 + (dZ2 - dZ1) * dT
End Function

'------------------------------------------------------------------------------
' Altitude a l'abscisse dAbs : lineaire entre les deux references qui
' l'encadrent, bornee aux altitudes extremes en dehors.
Function InterpolerZTroncons(dSRef() As Double, dZRef() As Double, _
                             ByVal nNbRef As Long, _
                             ByVal dAbs As Double) As Double
    If nNbRef <= 0 Then Exit Function
    If nNbRef = 1 Or dAbs <= dSRef(0) Then
        InterpolerZTroncons = dZRef(0)
        Exit Function
    End If
    If dAbs >= dSRef(nNbRef - 1) Then
        InterpolerZTroncons = dZRef(nNbRef - 1)
        Exit Function
    End If

    Dim i As Long
    For i = 1 To nNbRef - 1
        If dAbs <= dSRef(i) Then
            Dim dLen As Double
            dLen = dSRef(i) - dSRef(i - 1)
            If dLen < DIST_MIN Then
                ' Deux references confondues : on saute au lieu de diviser
                InterpolerZTroncons = dZRef(i)
            Else
                InterpolerZTroncons = dZRef(i - 1) _
                    + (dZRef(i) - dZRef(i - 1)) * (dAbs - dSRef(i - 1)) / dLen
            End If
            Exit Function
        End If
    Next
    InterpolerZTroncons = dZRef(nNbRef - 1)
End Function

'------------------------------------------------------------------------------
' Trie les couples (dSRef, dZRef) par abscisse croissante (tri par insertion :
' les intermediaires sont cliquees dans un ordre quelconque).
Sub TrierRefsParAbscisse(dSRef() As Double, dZRef() As Double, _
                         ByVal nNbRef As Long)
    Dim i As Long, j As Long, dS As Double, dZ As Double
    For i = 1 To nNbRef - 1
        dS = dSRef(i): dZ = dZRef(i)
        j = i - 1
        Do While j >= 0
            If dSRef(j) <= dS Then Exit Do
            dSRef(j + 1) = dSRef(j)
            dZRef(j + 1) = dZRef(j)
            j = j - 1
        Loop
        dSRef(j + 1) = dS
        dZRef(j + 1) = dZ
    Next
End Sub

'------------------------------------------------------------------------------
' Calcule l'altitude de chaque sommet d'une polyligne (abscisses dS) a partir
' des references triees. dZOut est ReDim aux memes bornes que dS.
Sub CalculerZPolyligne(dS() As Double, dSRef() As Double, dZRef() As Double, _
                       ByVal nNbRef As Long, dZOut() As Double)
    Dim nLo As Long, nHi As Long, i As Long
    nLo = LBound(dS): nHi = UBound(dS)
    ReDim dZOut(nLo To nHi)
    For i = nLo To nHi
        dZOut(i) = InterpolerZTroncons(dSRef, dZRef, nNbRef, dS(i))
    Next
End Sub

'------------------------------------------------------------------------------
' Insere dans la polyligne un sommet a chaque abscisse de dSIns qui n'y figure
' pas deja (tolerance DIST_MIN, bornes exclues). Les tableaux dX/dY/dS sont
' reconstruits 0-based. Necessaire avant CalculerZPolyligne : sans sommet a
' l'abscisse d'une reference intermediaire, le changement de pente serait
' perdu (ex. ligne droite de 2 sommets avec un regard au milieu).
Sub InsererAbscisses(dX() As Double, dY() As Double, dS() As Double, _
                     dSIns() As Double, ByVal nNbIns As Long)
    Dim dX0() As Double, dY0() As Double, dS0() As Double
    dX0 = dX: dY0 = dY: dS0 = dS

    Dim nLo As Long, nHi As Long, nNb0 As Long
    nLo = LBound(dS0): nHi = UBound(dS0)
    nNb0 = nHi - nLo + 1
    Dim dLong As Double
    dLong = dS0(nHi)

    Dim dTous() As Double
    ReDim dTous(0 To nNb0 + nNbIns - 1)
    Dim n As Long, i As Long, j As Long
    For i = nLo To nHi
        dTous(i - nLo) = dS0(i)
    Next
    n = nNb0

    For i = 0 To nNbIns - 1
        Dim dV As Double
        dV = dSIns(i)
        If dV > DIST_MIN And dV < dLong - DIST_MIN Then
            Dim bExiste As Boolean
            bExiste = False
            For j = 0 To n - 1
                If Abs(dTous(j) - dV) < DIST_MIN Then
                    bExiste = True
                    Exit For
                End If
            Next
            If Not bExiste Then
                dTous(n) = dV
                n = n + 1
            End If
        End If
    Next

    If n = nNb0 Then Exit Sub   ' rien a inserer

    ' Tri par insertion des abscisses fusionnees
    For i = 1 To n - 1
        Dim dV2 As Double
        dV2 = dTous(i)
        j = i - 1
        Do While j >= 0
            If dTous(j) <= dV2 Then Exit Do
            dTous(j + 1) = dTous(j)
            j = j - 1
        Loop
        dTous(j + 1) = dV2
    Next

    ' Reconstruction de la polyligne sur les abscisses fusionnees
    ReDim dX(0 To n - 1)
    ReDim dY(0 To n - 1)
    ReDim dS(0 To n - 1)
    For i = 0 To n - 1
        Dim dPX As Double, dPY As Double
        PointSurPolyligne2D dX0, dY0, dS0, dTous(i), dPX, dPY
        dX(i) = dPX: dY(i) = dPY
        dS(i) = dTous(i)
    Next
End Sub

'------------------------------------------------------------------------------
' Pente en % sur une longueur developpee. Longueur quasi nulle : 0.
Function PentePctLongueur(ByVal dZ1 As Double, ByVal dZ2 As Double, _
                          ByVal dLongueur As Double) As Double
    If dLongueur < DIST_MIN Then
        PentePctLongueur = 0#
    Else
        PentePctLongueur = (dZ2 - dZ1) / dLongueur * 100#
    End If
End Function

'==============================================================================
' Arc 3D par 3 points
'
' Calcule les parametres d'un arc circulaire passant par ptA (debut),
' ptB (milieu) et ptC (fin) dans l'espace 3D. Renvoie False si les
' trois points sont alignes.
' La methode : plan defini par les 3 points, circumcentre dans ce plan,
' rayon, matrice de rotation (RowX/RowY = plan, RowZ = normale),
' angle de depart et balayage.
'==============================================================================

Function ArcDepuis3Points(ptA As Point3d, ptB As Point3d, ptC As Point3d, _
                          ptCentre As Point3d, dRayon As Double, _
                          oRot As Matrix3d, dDebut As Double, _
                          dBalayage As Double) As Boolean
    ArcDepuis3Points = False

    ' Vecteurs AB et AC
    Dim abx As Double, aby As Double, abz As Double
    abx = ptB.X - ptA.X: aby = ptB.Y - ptA.Y: abz = ptB.Z - ptA.Z
    Dim acx As Double, acy As Double, acz As Double
    acx = ptC.X - ptA.X: acy = ptC.Y - ptA.Y: acz = ptC.Z - ptA.Z

    ' Normale au plan (AB x AC)
    Dim nx As Double, ny As Double, nz As Double
    nx = aby * acz - abz * acy
    ny = abz * acx - abx * acz
    nz = abx * acy - aby * acx
    Dim nLen As Double
    nLen = Sqr(nx * nx + ny * ny + nz * nz)
    If nLen < EPSILON Then Exit Function
    nx = nx / nLen: ny = ny / nLen: nz = nz / nLen

    ' Base orthonormee dans le plan : ex = AB/|AB|, ey = n x ex
    Dim dAB As Double
    dAB = Sqr(abx * abx + aby * aby + abz * abz)
    If dAB < EPSILON Then Exit Function
    Dim exx As Double, exy As Double, exz As Double
    exx = abx / dAB: exy = aby / dAB: exz = abz / dAB

    Dim eyx As Double, eyy As Double, eyz As Double
    eyx = ny * exz - nz * exy
    eyy = nz * exx - nx * exz
    eyz = nx * exy - ny * exx

    ' Projection de C dans le repere local (A = origine)
    Dim a As Double, b As Double
    a = acx * exx + acy * exy + acz * exz
    b = acx * eyx + acy * eyy + acz * eyz
    If Abs(b) < EPSILON Then Exit Function

    ' Circumcentre en 2D (A=(0,0), B=(dAB,0), C=(a,b))
    Dim cx As Double, cy As Double
    cx = dAB / 2#
    cy = (a * a + b * b - a * dAB) / (2# * b)

    dRayon = Sqr(cx * cx + cy * cy)

    ' Centre en 3D
    ptCentre.X = ptA.X + cx * exx + cy * eyx
    ptCentre.Y = ptA.Y + cx * exy + cy * eyy
    ptCentre.Z = ptA.Z + cx * exz + cy * eyz

    ' Matrice de rotation : colonnes = axes locaux (convention MicroStation)
    ' Colonne 1 = ex, Colonne 2 = ey, Colonne 3 = n
    oRot.RowX.X = exx: oRot.RowX.Y = eyx: oRot.RowX.Z = nx
    oRot.RowY.X = exy: oRot.RowY.Y = eyy: oRot.RowY.Z = ny
    oRot.RowZ.X = exz: oRot.RowZ.Y = eyz: oRot.RowZ.Z = nz

    ' Angles dans le repere local depuis le centre
    Dim angA As Double, angB As Double, angC As Double
    angA = Atn2Priv(-cy, -cx)
    angB = Atn2Priv(-cy, dAB - cx)
    angC = Atn2Priv(b - cy, a - cx)

    dDebut = angA

    ' Angle relatif de B et C par rapport a A, ramene dans [0, 2*PI)
    Dim dRelB As Double, dRelC As Double
    dRelB = angB - angA
    If dRelB < 0# Then dRelB = dRelB + 2# * PI
    If dRelB >= 2# * PI Then dRelB = dRelB - 2# * PI

    dRelC = angC - angA
    If dRelC < 0# Then dRelC = dRelC + 2# * PI
    If dRelC >= 2# * PI Then dRelC = dRelC - 2# * PI

    ' Si B est entre A et C en sens trigo, balayage = +dRelC
    ' Sinon inverser (sens horaire)
    If dRelB > 0# And dRelB < dRelC Then
        dBalayage = dRelC
    Else
        dBalayage = dRelC - 2# * PI
    End If

    ArcDepuis3Points = True
End Function

'==============================================================================
' Formatage
'==============================================================================

'------------------------------------------------------------------------------
Function FormatAltitude(dZ As Double, nDec As Integer, sSep As String) As String
    Dim sFmt As String
    If nDec <= 0 Then sFmt = "0" Else sFmt = "0." & String$(nDec, "0")
    Dim sRes As String
    sRes = Format$(dZ, sFmt)
    ' Format$ suit la locale Windows : on normalise puis on applique le separateur
    sRes = Replace(sRes, ",", ".")
    If sSep = "," Then sRes = Replace(sRes, ".", ",")
    FormatAltitude = sRes
End Function

'------------------------------------------------------------------------------
' Formate une pente en pourcentage en valeur absolue.
' nDec facultatif : nombre de decimales, defaut 1.
Function FormatPente(dPente As Double, sSep As String, _
                     Optional nDec As Integer = 1) As String
    Dim sFmt As String
    If nDec <= 0 Then sFmt = "0" Else sFmt = "0." & String$(nDec, "0")
    Dim sRes As String
    sRes = Format$(Abs(dPente), sFmt)
    sRes = Replace(sRes, ",", ".")
    If sSep = "," Then sRes = Replace(sRes, ".", ",")
    FormatPente = sRes & "%"
End Function
