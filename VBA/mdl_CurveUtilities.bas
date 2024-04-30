Attribute VB_Name = "mdl_CurveUtilities"
Option Explicit

'************************* RateCurve Factory *****************************************
Public Function rateFactory(rg As Range, Optional rgscen As Range, Optional rgten As Range) As clsRateCurve
' reads market data and, as the case may be, scenario data and initializes a rate curve object
'
' Arguments: 
'   rg: type: range - contains market data for a curve, may contain a variable number of curve instruments at different tenors
'   rgscen: type: range - contains scenario data, scenario data are applied to all market curves according to the specifications in rgscen
'   rgten: type: range - contains scenario tenors that are relevant for the scenario shifts in rgscen
'
' output:: clsRateCurve object
' 
    Dim strName As String
    Dim strType As String
    Dim refDate As Date
    Dim strCurrency As String
    Dim strFreq As String
    Dim tenor() As String, quotes() As Double, types() As String
    Dim size As Integer, i As Integer
     
    Dim RateCurve As New clsRateCurve
    
    If Not rgscen Is Nothing Then
        Dim scenTenors() As Double, scenShifts() As Double
        Dim ab As String, floor As String, marketshifts As String

        size = rgscen.Rows.count
        ReDim scenTenors(size - 4)
        ReDim scenShifts(size - 4)
        For i = 4 To rgscen.Rows.count
            scenTenors(i - 4) = rgten.Cells(i)
            scenShifts(i - 4) = rgscen.Cells(i) / 10000
        Next
        RateCurve.ScenarioTenors = scenTenors
        RateCurve.ScenarioShifts = scenShifts
        RateCurve.Market = rgscen.Cells(2)
        RateCurve.floor = rgscen.Cells(3)
        RateCurve.Absolut = rgscen.Cells(1)
        RateCurve.scenName = rgscen.Cells(0)
        
    End If
    
    ' check number of buckets in curve
   
    refDate = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    strCurrency = rg.Cells(1, 2)
    RateCurve.RefDatum = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    RateCurve.Frequency = rg.Cells(2, 2)
    
    ' allow for variable number of buckets
    size = rg.Rows.count - 5
    ReDim tenor(size - 1)
    ReDim quotes(size - 1)
    ReDim types(size - 1)
    
    For i = 0 To size - 1
        tenor(i) = rg.Cells(6 + i, 1)
        quotes(i) = rg.Cells(6 + i, 2) / 100
        types(i) = rg.Cells(6 + i, 3)
    Next
    
    If rgscen Is Nothing Then
        RateCurve.name = rg.Cells(3, 2)
        RateCurve.Kurventyp = rg.Cells(4, 2)
    Else
        RateCurve.name = rg.Cells(3, 2) & "_" & rgscen.Cells(0)
        RateCurve.BasisCurve = rg.Cells(3, 2)
        RateCurve.Kurventyp = rg.Cells(4, 2) & "_" & "scenario"
    End If
    
    RateCurve.CurveData_Tenor = tenor
    RateCurve.CurveData_Quote = quotes
    RateCurve.CurveData_Type = types
    
    ' ToDo: configure multiple curve bootstrapping
    ' here: placeholder for setting a discount curve different from the curve itself, e.g. �Str for 3M EURIBOR
    
    If strType = "Discount Curve" Then
        RateCurve.refcurve = rg.Cells(3, 2)
    Else
        RateCurve.refcurve = rg.Cells(3, 2)
    End If
    
    Set rateFactory = RateCurve

End Function

Function ReadandCalibrateCurve(blnPrint As Boolean, Optional scen As String = "default") As Dictionary
' read curve market data and calibrate curves
' iterate over curve data as given by tables "RateCurves" and "Scenarios" if parameter scen is set to it's default value, otherwise only for the given scenario
' optionally print curves to sheet
'
' Arguments: 
'   blnPrint: boolean, if TRUE curves are printed to sheet "CurveData_Calibrated", if FALSE curves are not printed; scen: string, if default value: all curves with all scenarios, otherwise only for the selected scenario
' 
' Output::
'   Dictionary, each item consists of a clsRateCurve, dictionary keys are the names of the curves (e.g. EUR3M_calibrated)

Dim curveDic As New Scripting.Dictionary, c As Variant, d As Variant, rg As Range, rgscen As Range, rgten As Range, curve As clsRateCurve, nm As String, index As Integer

     If scen <> "default" Then
        'If scen = "" Then
        ' always calibrate base scenario
            For Each c In ThisWorkbook.Worksheets(strConfiguration).Range(strRateCurves).Cells
                If Not c = strRateCurves Then
                    Set rg = ThisWorkbook.Worksheets(strMarketData).Range(c)
                    Set curve = mdl_CurveUtilities.rateFactory(rg)
                    curveDic.Add c.value, curve
                End If
            Next
        If scen <> "" Then                                ' selected scenario only --> the scnenario is applied to all available Curves in Range strRateCurves
            For Each c In ThisWorkbook.Worksheets(strConfiguration).Range(strRateCurves).Cells
                If Not c = strRateCurves Then
                    Set rg = ThisWorkbook.Worksheets(strMarketData).Range(c)
                    Set rgscen = ThisWorkbook.Worksheets(strMarketData).Range(scen)
                    Set rgten = ThisWorkbook.Worksheets(strMarketData).Range(strTenorScenarios)
                    Set curve = mdl_CurveUtilities.rateFactory(rg, rgscen, rgten)
                    nm = c.value & "_" & scen
                    curveDic.Add nm, curve
                End If
            Next
        End If
     Else                                   ' batch
        For Each c In ThisWorkbook.Worksheets(strConfiguration).Range(strRateCurves).Cells
            If Not c = strRateCurves Then
                Set rg = ThisWorkbook.Worksheets(strMarketData).Range(c)
                Set curve = mdl_CurveUtilities.rateFactory(rg)
                curveDic.Add c.value, curve
                For Each d In ThisWorkbook.Worksheets(strConfiguration).Range(strScenarios).Cells
                    If d.value <> "" Then
                        Set rgscen = ThisWorkbook.Worksheets(strMarketData).Range(d)
                        Set rgten = ThisWorkbook.Worksheets(strMarketData).Range(strTenorScenarios)
                        Set curve = mdl_CurveUtilities.rateFactory(rg, rgscen, rgten)
                        nm = c.value & "_" & d.value
                        curveDic.Add nm, curve
                    End If
                Next
            End If
        Next
    End If
    
    ' optionally print curves
    If blnPrint Then
        ' delete existing data first
        ThisWorkbook.Sheets(strCurveDataCalibrated).Cells.ClearContents
        deleteNames ("_calibrated")
        
        With ThisWorkbook.Sheets(strConfiguration).ListObjects(strAvailableCurves)
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete
            End If
        End With
        
        ' print
        index = 0
        For Each c In curveDic.Keys
            curveDic(c).printCurve strCurveDataCalibrated, index
            index = index + 1
        Next
    End If
    Set ReadandCalibrateCurve = curveDic
End Function

Function ReadCurvesFromSheet() As Dictionary
' reads all calibrated curves' grid information from sheet strCurveDataCalibrated
'
' output::
'   dicitionary containing curves, key correspond to curve names

Dim curveDic As New Scripting.Dictionary, c As Variant, rg As Range, curve As clsRateCurve, lo As ListObject, xName As String

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    For Each lo In ThisWorkbook.Sheets(strCurveDataCalibrated).ListObjects
        xName = lo.name
        If InStr(xName, "_calibrated") Then
            Dim index As Integer
            index = InStr(xName, "_calibrated")
            curveDic.Add Left(xName, index - 1), readCurveFromSheet(xName)
        End If
    Next
    
    Set ReadCurvesFromSheet = curveDic
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Function

Function readCurveFromSheet(curveName As String) As clsRateCurve
' Function reads data of a specific rate curve from the standard worksheet
' Arguments: 
'   curveName: name of the curve as available in the IRR Tool, i.e. as named table, must exist
' 
' Returns: 
'   clsRateCurve object
Dim curve As New clsRateCurve, index As Integer

    curve.name = Left(curveName, InStr(curveName, "_calibrated") - 1)
    curve.BasisCurve = Left(curveName, InStrRev(curve.getName, "_") - 1)
    curve.scenName = Right(curveName, InStrRev(curve.getName, "_") - 1)
    curve.BasisCurve = Left(curveName, InStrRev(curve.getName, "_") - 1)
    curve.RefDatum = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    curve.CurveGridData = ThisWorkbook.Sheets(strCurveDataCalibrated).Range(curveName)
    
    Set readCurveFromSheet = curve
End Function

Function getCurveName(curveFreq As Integer, scen As String) As String
' returns curve name based on frequency and scenario name
'
' Arguments: 
'   string: curveFreq as frequency of the intended curve
'   string: scenario name
'
' Returns: 
'   getCurveName: string

Dim rng As Range, name As String
'    If ThisWorkbook.Sheets("Portfolio").Range("d8").value = "None" Then
'        name = ThisWorkbook.Sheets("Portfolio").Range("d7").value
'    Else
'        name = ThisWorkbook.Sheets("Portfolio").Range("d7").value & "_" & ThisWorkbook.Sheets("Portfolio").Range("d8").value
'    End If
    If scen <> "" Then getCurveName = "EURIBOR_" & curveFreq & "M_" & scen Else getCurveName = "EURIBOR_" & curveFreq & "M"
End Function

Function getScenarioCurveName(curveName As String, scen As String) As String
    If scen <> "" Then
        getScenarioCurveName = curveName & "_" & scen
    Else
        getScenarioCurveName = curveName
    End If
End Function


Function getAvgDuration(amortizationScheme As String, startDate As Date, endDate As Date) As Double
    Dim amortization As Variant
    Dim result As Double
    Dim i As Integer
    If LCase(amortizationScheme) = "linear" Or amortizationScheme = "" Then
        getAvgDuration = 0.5 * (endDate - startDate) / 365#
    Else
        amortization = ThisWorkbook.Sheets(strConfiguration).Range(amortizationScheme).value
        For i = 1 To UBound(amortization, 2)
            result = result + i * amortization(1, i)
        Next i
        getAvgDuration = result / 12# ' convert from months to years
    End If
End Function

