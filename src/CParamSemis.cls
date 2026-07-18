VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamSemis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CParamSemis - Parametres de repartition des points semes (commande Semer)
'
' Deux modes (enum eModeSemis dans Trans3D.bas) :
'   semisDistanceFixe : un point au depart, puis tous les Pas, et un point a
'                       la fin meme si le dernier intervalle est plus court
'   semisPartsEgales  : Nombre points repartis uniformement, depart et fin
'                       inclus (minimum 2)
'
' Simple sac de proprietes, sans dependance a l'API MicroStation.
' Le calcul des abscisses est dans CCalcul (AbscissesSemis*).
'==============================================================================
Option Explicit

Public Mode         As eModeSemis    ' repartition : distance fixe / parts egales
Public Pas          As Double        ' distance fixe entre points (u.m.)
Public Nombre       As Long          ' nombre de points en parts egales

'------------------------------------------------------------------------------
Sub InitDefauts()
    Mode = semisDistanceFixe
    Pas = 0.5
    Nombre = 5
End Sub
