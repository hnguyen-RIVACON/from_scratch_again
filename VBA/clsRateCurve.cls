VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRateCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



' Class: clsRateCurve
'
' ****************** Variables *****************************************

Private strName As String
Private strScenName As String
Private strBasisCurve As String

Private intFrequency As Integer
Private refDate As Date
Private strCurveType As String ' Fwd-Curve / Valuation Curve
Private strRefCurve As String ' must be name of other curve if curve type is Fwd-Curve, can be curve itself

Private strAbsolute As String
Private strFloor As String
Private strMarket As String

' Input Data
Private strCurveData_Tenors() As String
Private dblCurveData_Quotes() As Double
Private strCurveData_Types() As String
Private dblScenarioData_Tenors() As Double
Private dblScenarioData_Shifts() As Double

' calculated Data
Private dtCurveDates() As Date
Private lngCurveTenors() As Long
Private dblCurveTenors() As Double
Private dblZeroRates() As Double
Private dblShiftedQuotes() As Double
Private dblApplicableQuotes() As Double

' grid Data
Private strgridTenors() As String
Private dtgridTenors() As Date
Private dblgridTenors() As Double
Private dblGridQuotes() As Double
Private dblgridShifts() As Double
Private dblgridShiftedQuotes() As Double
Private dblgridZeros() As Double
Private dblgridShiftedZeros() As Double

'***********************************************************************
'               Properties
'***********************************************************************
Property Let name(name As String)
    strName = name
End Property

Property Let Frequency(freq As String)
    If freq = "ON" Then
        intFrequency = 0            ' dummy value
    Else
        intFrequency = Left(freq, 1)
    End If
End Property

Property Let RefDatum(datum As Date)
    refDate = datum
End Property

Property Let Kurventyp(typ As String)
    strCurveType = typ
End Property

Property Let refcurve(refcurve As String)
    strRefCurve = refcurve
End Property
Property Let BasisCurve(strBasis As String)
    strBasisCurve = strBasis
End Property

Property Let scenName(strScen As String)
    strScenName = strScen
End Property
Property Let floor(inte As String)
    strFloor = inte
End Property
Property Let Absolut(inte As String)
    strAbsolute = inte
End Property
Property Let Market(inte As String)
    strMarket = inte
End Property
Property Let ScenarioTenors(tenors() As Double)
    Dim size As Integer
    size = UBound(tenors)
    ReDim dblScenarioData_Tenors(size)
    dblScenarioData_Tenors = tenors
End Property
Property Let ScenarioShifts(quotes() As Double)
    Dim size As Integer
    size = UBound(quotes)
    ReDim dblScenarioData_Shifts(size)
    dblScenarioData_Shifts = quotes
End Property

Property Let CurveData_Tenor(tenor() As String)
    Dim size As Integer
    size = UBound(tenor)
    
    ReDim strCurveData_Tenors(size)
    ReDim dtCurveDates(size)
    ReDim dblZeroRates(size)
    ReDim lngCurveTenors(size)
    ReDim dblCurveTenors(size)
    
    strCurveData_Tenors = tenor
    CalcCurveDates_Tenors
    
End Property
Property Let CurveData_Quote(quotes() As Double)
    Dim size As Integer
    size = UBound(quotes)
    ReDim dblCurveData_Quotes(size)
    dblCurveData_Quotes = quotes
End Property

Property Let CurveData_Type(types() As String)
    Dim size As Integer
    size = UBound(types)
    ReDim strCurveData_Types(size)
    strCurveData_Types = types
    
    If Not array_Empty(dblScenarioData_Tenors) And strMarket = "yes" Then CalcShiftedQuotes
    CalcZeroRates
    CalcGrid  ' Routine bestimmt f�r die gew�nschte Granularit�t der Tenoren, den Output-Grid mit Dates, Yearfrac, Tenors, Market quotes, scenario shifts und zero rates (interpol.)
    
End Property
Property Let CurveGridData(grid As Range)
    
ReDim strgridTenors(grid.Rows.count - 1)
ReDim dtgridTenors(grid.Rows.count - 1)
ReDim dblgridTenors(grid.Rows.count - 1)
ReDim dblGridQuotes(grid.Rows.count - 1)
ReDim dblgridShifts(grid.Rows.count - 1)
ReDim dblgridShiftedQuotes(grid.Rows.count - 1)
ReDim dblgridZeros(grid.Rows.count - 1)
ReDim dblgridShiftedZeros(grid.Rows.count - 1)
  Dim i As Integer
  

 For i = 1 To grid.Rows.count
    strgridTenors(i - 1) = grid.Cells(i, 1)
    dtgridTenors(i - 1) = grid.Cells(i, 2)
    dblgridTenors(i - 1) = grid.Cells(i, 3)
    dblGridQuotes(i - 1) = grid.Cells(i, 4)
    dblgridShifts(i - 1) = grid.Cells(i, 5)
    dblgridShiftedQuotes(i - 1) = grid.Cells(i, 6)
    dblgridZeros(i - 1) = grid.Cells(i, 7)
    dblgridShiftedZeros(i - 1) = grid.Cells(i, 8)
 Next
 
End Property

Public Property Get getCurveSize() As Integer
    If Not (Not strCurveData_Tenors) Then
        getCurveSize = UBound(strCurveData_Tenors)
    Else
        MsgBox "The Curve hasn't been initialized yet."
    End If
End Property

Public Property Get getCurveRefDate()
    getCurveRefDate = refDate
End Property
Public Property Get getName()
    getName = strName
End Property


'***********************************************************************
'               Methods
'***********************************************************************
'************************* Public Function Discount Factor **************
Private Function getDFBoot(dtInterpol As Date) As Double
' Function calculates a discount factor based on the original curve grid
' function is used during bootstrapping
'
' Paramters: dtInterpol: date for which discount factor is needed
'
' Uses:
' - members: dblgridTenors, dblgridShiftedZeros, RefDate
' - public method: LIP
Dim rate As Double, date0 As Long, tenor As Double
    Dim df As Double
    tenor = WorksheetFunction.yearfrac(refDate, dtInterpol, 3)
    rate = LIP(dblCurveTenors, dblZeroRates, tenor)
    df = Exp(-rate * tenor)
    getDFBoot = df
End Function

'************************* Zero Rate Calculation *******************************************************************************
Private Function getZRBoot(yearfrac As Double) As Double
' Function calculates a zero rate by linear interpolation from the original curve grid
' function is used during bootstrapping
'
' Paramters: yearfrac
'
' Uses:
' - members: dblgridTenors, dblgridShiftedZeros
' - public method: LIP
    getZRBoot = LIP(dblCurveTenors, dblZeroRates, yearfrac)
End Function

'************************* Swap Rate Calculation *******************************************************************************
Private Function calcSwapRateBoot(mat As Double) As Double
' Function calculates a fair swap rate on the current curve and based on the original grid,
' function is used during bootstrapping
'
' parameters: mat as swap maturity in years
'
' Uses:
' members RefDAte, intFrequency
'
' methods: getDF
Dim i As Integer, df As Double, A As Double
    Dim yf As Double
    A = 0
    For i = 1 To mat / (intFrequency / 12)
        yf = WorksheetFunction.yearfrac(refDate, DateAdd("m", i * intFrequency, refDate), 3)
        df = getDFBoot(DateAdd("m", i * intFrequency, refDate))
        A = A + df
    Next
    calcSwapRateBoot = 12 / intFrequency * (1 - df) / A
    
End Function
'************************* Private Fwd Rate Calculation *******************************************************************************
Private Function calcFwdRateBoot(yearfrac1 As Double, yearfrac2 As Double) As Double
' Function calculates a forward rate based on the oroginal curve grid
' function is needed in the course of bootstrapping
'
' Paramters:
'   yearfrac1
'   yearfrac2
    calcFwdRateBoot = Exp(-getZRBoot(yearfrac1) * yearfrac1) / Exp(-getZRBoot(yearfrac2) * yearfrac2) - 1
End Function


'************************* Public Function Discount Factor **************

Public Function getDF(dtInterpol As Date) As Double
' Function calculates a discount factor based on the final curve grid
'
' Paramters: 
'   dtInterpol: date for which discount factor is needed
'
' Uses:
'   members: dblgridTenors, dblgridShiftedZeros, RefDate
'   public method: LIP

    Dim rate As Double, date0 As Long, tenor As Double

    tenor = WorksheetFunction.yearfrac(refDate, dtInterpol, 3)          ' ToDo: Check for applicable dcc
    rate = LIP(dblgridTenors, dblgridShiftedZeros, tenor)
    getDF = Exp(-rate * tenor)

End Function

'************************* Public Function Discount Factor **************
Public Function getDFByYF(yf As Double) As Double
' Function calculates a discount factor based on the final curve grid as function of year fraction
'
' Paramters: yf: year fraction for which discount factor is needed
'
' Uses::
'   members: dblgridTenors, dblgridShiftedZeros, RefDate
'   public method: LIP

    Dim rate As Double
    rate = LIP(dblgridTenors, dblgridShiftedZeros, yf)
    getDFByYF = Exp(-rate * yf)

End Function


'************************* Zero Rate Calculation *******************************************************************************
Public Function getZR(yearfrac As Double) As Double
' Function calculates a zero rate by linear interpolation from the final curve grid
' Paramters: yearfrac
' Uses:
' - members: dblgridTenors, dblgridShiftedZeros
' - public method: LIP
    getZR = LIP(dblgridTenors, dblgridShiftedZeros, yearfrac)
End Function

'************************* Swap Rate Calculation *******************************************************************************
Public Function calcSwapRate(mat As Double) As Double
' Function calculates a fair swap rate on the current curve and based on the finale grid
' parameters: mat as swap maturity in years
' Uses:
' members RefDAte, intFrequency
' methods: getDF

Dim i As Integer, df As Double, A As Double

    A = 0
    For i = 1 To mat / (intFrequency / 12)
        yf = WorksheetFunction.yearfrac(refDate, DateAdd("m", i * intFrequency, refDate), 3)
        df = getDF(DateAdd("m", i * intFrequency, refDate))
        A = A + df
    Next
    calcSwapRate = 12 / intFrequency * (1 - df) / A
    
End Function
'************************* Fwd Rate Calculation *******************************************************************************
Public Function calcFwdRate(yearfrac1 As Double, yearfrac2 As Double) As Double
' Function calculates a forward rate based on the final curve grid
'
' Paramters:
'   yearfrac1
'   yearfrac2
'
    If yearfrac2 - yearfrac1 = 0# Then
        calcFwdRate = 0#
    Else
        calcFwdRate = (Exp(-getZR(yearfrac1) * yearfrac1) / Exp(-getZR(yearfrac2) * yearfrac2) - 1) / (yearfrac2 - yearfrac1)
    End If
End Function

'************************* Print Data *****************************************************************************************
Public Sub printCurve(sheetName As String, Optional nr As Integer = 0)
' sub prints the final curve grid on a worksheet, formats the printed curve as an excel table and assigns a name: strName + "_calibrated"
' adds the name of the table to the range strAvailableCurves (see globalconst for details)
' paramters
' - sheetName: name of the worksheet where the curve is printed
' - nr: integer that indicates in combination with the hardcoded blocksize, where in sheetName the data is to be printed
' hardcoded: blocksize, ensuring an empty column between two sets of curve data
Dim blockSize As Integer, tbl As ListObject, newRow As ListRow

    Application.ScreenUpdating = False
    
    ThisWorkbook.Sheets(sheetName).Activate
    blockSize = 9
    
    printVecOnSheet strgridTenors, sheetName, strName & "_Tenor", nr * blockSize
    printVecOnSheet dtgridTenors, sheetName, strName & "_Date", nr * blockSize
    printVecOnSheet dblgridTenors, sheetName, strName & "_YearFrac", nr * blockSize
    printVecOnSheet dblGridQuotes, sheetName, strName & "_Market Quotes", nr * blockSize
    printVecOnSheet dblgridShifts, sheetName, strName & "_ScenarioShifts", nr * blockSize
    printVecOnSheet dblgridShiftedQuotes, sheetName, strName & "_Shifted_Quotes", nr * blockSize
    printVecOnSheet dblgridZeros, sheetName, strName & "_ZeroRates", nr * blockSize
    printVecOnSheet dblgridShiftedZeros, sheetName, strName & "_Shifted_ZeroRates", nr * blockSize
    
    Dim src As Range, ws As Worksheet
    ThisWorkbook.Sheets(sheetName).Activate
    Cells(1, 1 + nr * blockSize).Activate
    Set src = Range(ActiveCell, ActiveCell).CurrentRegion
    Set ws = ActiveSheet
    ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=src, _
    xlListObjectHasHeaders:=xlYes, tablestyleName:="TableStyleMedium28").name = strName & "_calibrated"
    
    Set tbl = Range(strAvailableCurves).ListObject
    Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
    newRow.Range = strName & "_calibrated"
    

    ThisWorkbook.Sheets(strConfiguration).Activate
    Cells(1, 1).Activate
    Application.ScreenUpdating = True
End Sub
'************************* Calculate Print Data *****************************************************************************
Private Sub CalcGrid()
' Sub sets up the final grid, i.e. a set of data vectors at a monthly granularity for the first ten years, for maturities beyond 10 years the grid is as granular as the original curve
' members to be set:
' - strgridTenors: Tenors in notation '1M', '2M' etc. for all buckets of the final grid
' - dtgridTenors: dates of the grid
' - dblgridTenors: yearfraction, using act/365 dcc convention, of each bucket as of the reference date
' - dblgridShifts: scenario shifts: if applicable to market quotes dblgridShifts has entry only at tenors of market quotes, otherwise, dblgridShifts holds the interpolated shifts at each tenor of the grid
' - dblgridShiftedQuotes: shifted quotes = market quotes if scenario shifts are applied to zero rates, otherwise shifted quotes = dblShifteQuotes at corresponding buckets
' - dblgridShiftesZeros: zero rates at grid tenors after interpolation and, if applicable, scenario shifts
' - dblgridZeros: zero rates at the original curve tenors

    Dim i As Integer, j As Integer
    Dim nrEntries As Integer
    

' Assume monthly granularity up to 10 years (i.e. 120 entries)
' Tenors, Dates, Yearfrac, Market Quotes, scenario shifts (only for market quotes if applicable), scenario zeros

    ' check how many entries necessary
    For i = 0 To UBound(dblCurveTenors)
        If dblCurveTenors(i) > 10 Then Exit For
    Next
    nrEntries = UBound(dblCurveTenors) - i + intSizeCurveGrid
    
    ReDim strgridTenors(nrEntries)
    ReDim dtgridTenors(nrEntries)
    ReDim dblgridTenors(nrEntries)
    ReDim dblGridQuotes(nrEntries)
    ReDim dblgridShifts(nrEntries)
    ReDim dblgridShiftedQuotes(nrEntries)
    ReDim dblgridShiftedZeros(nrEntries)
    ReDim dblgridZeros(nrEntries)
    Dim dblgridInterpolZeros() As Double
    ReDim dblgridInterpolZeros(nrEntries)
    
    For j = 0 To 120 - 1
        strgridTenors(j) = (j + 1) & "M"
        dtgridTenors(j) = DateAdd("d", -1, DateAdd("m", j + 1, DateAdd("d", 1, refDate)))
        dblgridTenors(j) = WorksheetFunction.yearfrac(refDate, dtgridTenors(j), 3)
    Next
    For j = 120 To nrEntries
        strgridTenors(j) = strCurveData_Tenors(i)
        dtgridTenors(j) = dtCurveDates(i)
        dblgridTenors(j) = dblCurveTenors(i)
        i = i + 1
    Next
    ' get quoted data and zero rates (based on shifted market quotes as the case may be)
    For i = 0 To UBound(dblCurveData_Quotes)
        dblGridQuotes(Application.Match(dblCurveTenors(i), dblgridTenors, False) - 1) = dblCurveData_Quotes(i)
        dblgridZeros(Application.Match(dblCurveTenors(i), dblgridTenors, False) - 1) = dblZeroRates(i)
    Next
    ' get shifts
    If Not array_Empty(dblScenarioData_Shifts) Then
        If strMarket = "yes" Then                   ' shifts only at market quote tenors
            For i = 0 To UBound(dblCurveTenors)
                dblgridShifts(Application.Match(dblCurveTenors(i), dblgridTenors, False) - 1) = LIP(dblScenarioData_Tenors, dblScenarioData_Shifts, dblCurveTenors(i))
                dblgridShiftedQuotes(Application.Match(dblCurveTenors(i), dblgridTenors, False) - 1) = dblShiftedQuotes(i)
            Next
             For i = 0 To UBound(dblgridShifts)
                dblgridInterpolZeros(i) = LIP(dblCurveTenors, dblZeroRates, dblgridTenors(i))
                dblgridShiftedZeros(i) = dblgridInterpolZeros(i)
            Next
        Else                                        ' shifts at all tenors, e.g. at grid level, applicable if zero rates are shifted
            For i = 0 To UBound(dblCurveTenors)
                dblgridShiftedQuotes(Application.Match(dblCurveTenors(i), dblgridTenors, False) - 1) = dblCurveData_Quotes(i) ' quotes are not shifted
            Next
            For i = 0 To UBound(dblgridShifts)
                dblgridShifts(i) = LIP(dblScenarioData_Tenors, dblScenarioData_Shifts, dblgridTenors(i))
                ' Opional: add other interpolation scheme for granular grid
                dblgridInterpolZeros(i) = LIP(dblCurveTenors, dblZeroRates, dblgridTenors(i))
                dblgridShiftedZeros(i) = applyShift(dblgridTenors(i), dblgridInterpolZeros(i), dblgridShifts(i))
            Next
            'Erase dblgridIterpolZeros
        End If
    Else
         For i = 0 To UBound(dblgridShifts)
            dblgridShiftedZeros(i) = LIP(dblCurveTenors, dblZeroRates, dblgridTenors(i))
        Next
    End If

End Sub

Private Sub CalcShiftedQuotes()
' Sub calculates values for the member dblShiftedQuotes() based on members
' dblCurveTenors, dblCurveData_Quotes, and dblScenarioData_Shifts
Dim i As Integer
    
    ReDim dblShiftedQuotes(UBound(dblCurveTenors))
    For i = 0 To UBound(dblCurveTenors)
        dblShiftedQuotes(i) = applyShift(dblCurveTenors(i), dblCurveData_Quotes(i), LIP(dblScenarioData_Tenors, dblScenarioData_Shifts, dblCurveTenors(i)))
    Next
    
End Sub

'**************************** calculate shifted value (regardless of market quote or zero rate) *******
Private Function applyShift(yf As Double, rate As Double, shift As Double) As Double
' Function calculates the scenario rate using the scenario configuration given in the member variables strAbsolute and strFloor
' Parameters:
' - yf: yearfraction of the rate
' - rate: unshifted rate
' - shift: shift size
Dim value As Double, applicableFloor As Double

    If strAbsolute = "absolut" Or strAbsolute = "absolute" Then
        value = rate + shift
    ElseIf strAbsolute <> "absolut" Then
        If shift > 0 Then
            value = rate + Abs(rate * shift)
        Else
            value = rate - Abs(rate * shift)
        End If
    End If
    applicableFloor = ThisWorkbook.Sheets("Configuration").Range("Floor") + ThisWorkbook.Sheets("Configuration").Range("Slope") * yf
    If strFloor = "yes" Then value = WorksheetFunction.Max(applicableFloor, value)
    
    applyShift = value
End Function
Private Sub CalcCurveDates_Tenors()
' Based on user provided tenors of the curve instruments, strCurveData_Tenors(), the sub calculates data for the private members
' - the corresponding bucket dates, dtCurveDates()
' - the corresponing yearfractions (applying a act/365 convention), dblCurveTenors()
' - the number of days between bucket dates and reference date, lngCurveTenors()

Dim i As Integer, nr As Integer
    For i = 0 To getCurveSize()
        nr = Left(strCurveData_Tenors(i), Len(strCurveData_Tenors(i)) - 1)
        Select Case Right(strCurveData_Tenors(i), 1)
            Case "M", "m"
                dtCurveDates(i) = DateAdd("d", -1, DateAdd("m", nr, DateAdd("d", 1, refDate)))
                dblCurveTenors(i) = WorksheetFunction.yearfrac(refDate, dtCurveDates(i), 3)
            Case "Y", "y", "J", "j"
                dtCurveDates(i) = DateAdd("d", -1, DateAdd("yyyy", nr, DateAdd("d", 1, refDate)))
                dblCurveTenors(i) = WorksheetFunction.yearfrac(refDate, dtCurveDates(i), 3)
            Case "W", "w"
                dtCurveDates(i) = DateAdd("d", -1, DateAdd("ww", nr, DateAdd("d", 1, refDate)))
                dblCurveTenors(i) = WorksheetFunction.yearfrac(refDate, dtCurveDates(i), 3)
            Case Else
                dtCurveDates(i) = DateAdd("d", nr, refDate)
                dblCurveTenors(i) = WorksheetFunction.yearfrac(refDate, dtCurveDates(i), 3)                    ' to be checked for leap years
        End Select
        lngCurveTenors(i) = dtCurveDates(i) - refDate
    Next
End Sub

'***********************************************************************************************************************
'                                   Methods: Bootstrapping Functions
'***********************************************************************************************************************
Private Sub CalcZeroRates()
' Sub sets the member dblApplicableQuotes and calculates zero rates, member dblZeroRates,
' Uses:
' - members: dblCurveTenors, dblCurveData_Quotes, i.e. the user provided curve tenors,
' - methods: bootstrap

Dim i As Integer
    
    ' first: determine relevant rates
    If Not array_Empty(dblShiftedQuotes) Then
        dblApplicableQuotes = dblShiftedQuotes
    Else
        dblApplicableQuotes = dblCurveData_Quotes
    End If
    ' second: calculate zero rates
    For i = 0 To getCurveSize()
        Select Case strCurveData_Types(i)
        Case "ZeroRate", "Zero Rate", "zerorate", "zero rate"
            dblZeroRates(i) = dblApplicableQuotes(i)
        Case "Deposit", "deposit"                           ' Deposit are uncollaterlized instruments with a single payment at maturity --> deposit rate equals zero rate (only needs to be converted into continuous rate)
            dblZeroRates(i) = (1 / dblCurveTenors(i)) * Log(1 + dblApplicableQuotes(i) * dblCurveTenors(i))
        Case "Swap Rate", "SwapRate", "Swap", "swap"        ' Swaps are collateralized instruments with intermediate cash flows --> bootstrapping is required
            dblZeroRates(i) = bootstrap(i)
            Debug.Print "line item " & i & ": " & calcSwapRateBoot(dblCurveTenors(i)) & " vs quoted rate: " & dblCurveData_Quotes(i) & " Differenz: " & calcSwapRateBoot(dblCurveTenors(i)) - dblCurveData_Quotes(i)
        Case Else
            MsgBox "Unidentified Curve Type."
        End Select
    Next
End Sub

Private Function bootstrap(item As Integer) As Double
' Function calculates based on already existing zero rates for previous periods the still unknown zero rate till maturity of the swap
' Uses:
' - members dtCurveDates, dblApplicableQuotes, intFrequency, dblZeroRates
' - methods: getZRBoot, swapPaymentDates, Newton
' Output: zero rate for at swap maturity

    Dim i As Integer, index As Integer
    Dim A As Double, yf As Double, start As Double
    Dim t() As Double
    ' required: swap tenor, number of payments
    ' setup SwapDates
    Dim swapPayDates() As Date
    swapPayDates() = swapPaymentDates(item)                  ' assume fixing dates match payment dates
    
    ' calculate known parts
    A = 0
    i = 0
    Do While swapPayDates(i) <= dtCurveDates(item - 1)
        yf = WorksheetFunction.yearfrac(refDate, swapPayDates(i), 3)
        A = A + Exp(-getZRBoot(yf) * yf)
        i = i + 1
    Loop
    'A = A + dblCurveData_Quotes(i) / intFrequency
    ReDim t(UBound(swapPayDates) - i)
    For index = i To UBound(swapPayDates)
        t(index - i) = WorksheetFunction.yearfrac(refDate, swapPayDates(index), 3)
    Next
    start = (dblApplicableQuotes(item) + dblZeroRates(item - 1)) / 2
    bootstrap = Newton(t, A, dblApplicableQuotes(item) / (12 / intFrequency), start)
    

End Function
Private Function swapPaymentDates(item As Integer) As Date()
' Function calclates for a given bucket the payment dates of a swap
' Uses
' - members: intFrequency, strCurveData_Tenors, RefDate
' Output: date vector
Dim i As Integer, nr_payments As Integer, nr As Integer

    nr = Left(strCurveData_Tenors(item), Len(strCurveData_Tenors(item)) - 1)
    Select Case Right(strCurveData_Tenors(item), 1)
        Case "M", "m"
            nr_payments = nr / intFrequency
            ' add check for non matching maturity with frequency
        Case "Y", "y"
            nr_payments = nr * 12 / intFrequency
        Case Else
            MsgBox "Unidentified swap maturity type."
    End Select
    Dim Dates() As Date
    ReDim Dates(nr_payments - 1)
    For i = 0 To nr_payments - 1
        Dates(i) = DateAdd("d", -1, DateAdd("m", intFrequency * (i + 1), DateAdd("d", 1, refDate)))
    Next
    swapPaymentDates = Dates
End Function

'*********************************** Simple Newton ***************************************************************
Private Function Newton(t() As Double, A As Double, coupon As Double, start As Double) As Double
' Function approximates a zero rate by using the Newton method
' parameters:
' - t(): vector of tenors of the underlying swap in yearfraction
' - A: known discounted value of the fixed swap leg
' - Coupon: swap coupon (i.e. swap rate divided by 12 times intefrequency)
' - start: start value for the algorithm: provided average of the last boostrapped zero rate and the applicable quote at the current curve tenor

Dim err As Double, errprev As Double, eps As Double, xopt As Double, xi As Double, change As Integer
    
    eps = 0.0000001
    xopt = start - calcFr(t, A, coupon, start) / calcDerivativeFr(t, coupon, start)
    err = Abs((xopt - start) / xopt) * 100
    'Debug.Print xopt & " | " & start & " | " & err & vbNewLine
    xi = xopt
    errprev = err
    change = 1
    Do While err > eps
        xopt = xi - change * calcFr(t, A, coupon, xi) / calcDerivativeFr(t, coupon, xi)
        err = Abs((xopt - xi) / xopt) * 100
        errprev = err
        If err < errprev Then
            change = -1
        Else
            change = 1
        End If
        'Debug.Print xopt & " | " & xi & " | " & err & vbNewLine
        xi = xopt
    Loop
    
    Newton = xopt

End Function

Private Function calcFr(yearfrac() As Double, A As Double, coupon As Double, x As Double) As Double
' function calculates the value of the swap function at a specified value
' parameters:
' - yearfrac: tenor given as yearfrac
' - A: known discounted value part of the fix leg
' - Coupon: swap coupon (i.e. swap rate divided by 12 times intefrequency)
' - x: current guess of zero rate
Dim i As Integer, value As Double

    value = 0
    For i = 0 To UBound(yearfrac)
        value = value + Exp(-x * yearfrac(i))
    Next
    value = (A + value) * coupon + Exp(-x * yearfrac(i - 1)) - 1
    
    calcFr = value
    'Debug.Print calcFr

End Function

Private Function calcDerivativeFr(yearfrac() As Double, coupon As Double, x As Double) As Double
' function calculates the value of the swap function derivative at a specified value
' parameters:
' - yearfrac: tenor given as yearfrac
' - Coupon: swap coupon (i.e. swap rate divided by 12 times intefrequency)
' - x: current guess of zero rate
Dim i As Integer, value As Double

    value = 0
    For i = 0 To UBound(yearfrac)
        value = value - yearfrac(i) * Exp(-x * yearfrac(i))
    Next
    value = value * coupon - yearfrac(i - 1) * Exp(-x * yearfrac(i - 1))
    
    calcDerivativeFr = value
    'Debug.Print calcDerivativeFr
    
End Function



