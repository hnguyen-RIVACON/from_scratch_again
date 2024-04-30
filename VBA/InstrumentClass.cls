VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstrumentClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public longName As String
Public shortName As String
Public insType As InstrumentType
Public isAsset As Boolean
Public isLiability As Boolean

Public costRatio As Double


Sub init(tp As InstrumentType, sn As String, ln As String, isa As Boolean, isl As Boolean)
    insType = tp
    shortName = sn
    longName = ln
    isAsset = isa
    isLiability = isl
    costRatio = 0#
End Sub

