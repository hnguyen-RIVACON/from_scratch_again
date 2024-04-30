VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public repType As ReportType
Public shortName As String
Public longName As String

Private insTypes As Dictionary

Public Function instrumentTypeEnabled(insType As InstrumentType) As Boolean
    instrumentTypeEnabled = False
    If insTypes.Exists(insType) Then
        instrumentTypeEnabled = insTypes(insType)
    End If
End Function

Public Function getEnabledInstruments() As Dictionary
    Set getEnabledInstruments = insTypes
End Function

Sub init(tp As ReportType, sn As String, ln As String, instrumentClasses As Dictionary)
    repType = tp
    shortName = sn
    longName = ln
    Set insTypes = New Dictionary
    Dim confRange As Range
    Set confRange = ThisWorkbook.Sheets(strConfiguration).Range(strEnabledInstruments)
    Dim repColumn As Integer
    Dim i As Integer
    Dim insType As InstrumentType
    Dim enabled As Boolean
    repColumn = 1
    Do While repColumn < confRange.Columns.count And confRange.Cells(1, repColumn).value <> sn
        repColumn = repColumn + 1
    Loop
    If confRange.Cells(1, repColumn).value <> sn Then
        MsgBox "No Configuration for report " + sn + " present"
    Else
        For i = 2 To confRange.Rows.count
            insType = getTypeFromShortName(instrumentClasses, confRange(i, 1).value)
            enabled = confRange(i, repColumn).value
            insTypes.Add insType, enabled
        Next i
    End If

End Sub

