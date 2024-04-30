Attribute VB_Name = "mdl_Testing"
Option Explicit


Public Function getFwdRateTest(curveName As String, yf1 As Double, yf2 As Double) As Double
Dim curve As New clsRateCurve, index As Integer

    curve.name = curveName
    curve.BasisCurve = Left(curveName, InStrRev(curve.getName, "_") - 1)
    curve.scenName = Right(curveName, InStrRev(curve.getName, "_") - 1)
    curve.BasisCurve = Left(curveName, InStrRev(curve.getName, "_") - 1)
    curve.RefDatum = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    curve.CurveGridData = ThisWorkbook.Sheets(strCurveDataCalibrated).Range(curveName)
    
    getFwdRateTest = curve.calcFwdRate(yf1, yf2)
End Function

Public Function getZeroRateTest(curveName As String, yf1 As Double) As Double
Dim curve As New clsRateCurve, index As Integer

    curve.name = curveName
    curve.BasisCurve = Left(curveName, InStrRev(curve.getName, "_") - 1)
    curve.scenName = Right(curveName, InStrRev(curve.getName, "_") - 1)
    curve.BasisCurve = Left(curveName, InStrRev(curve.getName, "_") - 1)
    curve.RefDatum = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    curve.CurveGridData = ThisWorkbook.Sheets(strCurveDataCalibrated).Range(curveName)
    
    getZeroRateTest = curve.getZR(yf1)
End Function

Public Function getDFTest(curveName As String, dt As Date) As Double
Dim curve As New clsRateCurve, index As Integer

    curve.name = curveName
    curve.BasisCurve = Left(curveName, InStrRev(curve.getName, "_") - 1)
    curve.scenName = Right(curveName, InStrRev(curve.getName, "_") - 1)
    curve.BasisCurve = Left(curveName, InStrRev(curve.getName, "_") - 1)
    curve.RefDatum = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    curve.CurveGridData = ThisWorkbook.Sheets(strCurveDataCalibrated).Range(curveName)
    
    getDFTest = curve.getDF(dt)
End Function
Public Function getSwapRateTest(curveName As String, mat As Double) As Double
Dim curve As New clsRateCurve, index As Integer

    curve.name = curveName
    curve.BasisCurve = Left(curveName, InStrRev(curve.getName, "_") - 1)
    curve.scenName = Right(curveName, InStrRev(curve.getName, "_") - 1)
    curve.BasisCurve = Left(curveName, InStrRev(curve.getName, "_") - 1)
    curve.RefDatum = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    curve.CurveGridData = ThisWorkbook.Sheets(strCurveDataCalibrated).Range(curveName)
    
    getDFTest = curve.calcSwapRate(mat)
End Function

