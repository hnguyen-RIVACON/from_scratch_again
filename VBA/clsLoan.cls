VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Class: clsLoan
'
'   description here

' ****************** Variables *****************************************
Private strLoanName As String

Private blnIsAsset As Integer
Private strBalanceItem As String

Private dtMat As Date
Private dblStartNom As Double
Private dblOrigNom As Double
Public floatCoupon As Boolean      ' true for floating coupon, false for fixed coupon
Private intPayFreq As Integer
Private intDCC As Integer           ' day count convention
Private strBDC As String            ' business day convention
Private blnAmort As Boolean         ' Amortizing yes / no
Private dblCoupon As Double         ' current coupon if float, coupon if fix
Public prolongationCoupon As Double
Public floatFloor As Double         ' floor for floating coupon
Private varAmortScheme As Variant
Private dtCur As Date
Private intMargin As Integer        ' floating coupon spread in basis points
Private variableMargin() As Double
Private hasVariableMargin As Boolean
Public useAverageNotional As Boolean
Private dtNextPay As Date
Private dtLastDate As Date
Private strAssetType As String
Public forwardCurve As String          ' name of curve used to obtain forward rates (if empty, intPayFreq is used to determine the forward curve)
Private paymentDay As Integer          ' day in month on which payments are made
Public bookValue As Double
Public includeAmortizationCF As Boolean ' whether to include amortization cash flows
Public prolongationDuration As Double ' duration of prolongation modelling
' calculated data
Private dtLoanDates() As Date
Private dblNominals() As Double
Private dblNPV As Double
Private dblACFs() As Double
Private dblIRCFs() As Double
Private dblCFs() As Double
Private dblNCFs() As Double
Private dblDFs() As Double ' discount factors
Private expiringNotional() As Double


Private Sub Class_Initialize()
    includeAmortizationCF = True
    hasVariableMargin = False
    useAverageNotional = True
    floatFloor = -1E+99
End Sub

'***********************************************************************
'               Properties
'***********************************************************************
Property Let name(name As String)
    strLoanName = name
End Property
Property Let BalanceItem(name As String)
    strBalanceItem = name
End Property
Property Let isAsset(bln As Boolean)
    If bln Then blnIsAsset = 1 Else blnIsAsset = -1
End Property
Property Let maturity(mat As Date)
    dtMat = mat
End Property
Property Let LastDate(start As Date)
    dtLastDate = start
End Property
Property Let Nominal(n As Double)
    dblStartNom = n
End Property
Property Let OrigNominal(n As Double)
    dblOrigNom = n
End Property
Property Let assetType(t As String)
    strAssetType = t
End Property

Property Let varMargin(m() As Double)
    variableMargin = m
    hasVariableMargin = True
End Property

Property Let typ(typ As String)
    Dim lType As String
    lType = LCase(typ)
    If lType = "" Then
        floatCoupon = False
        MsgBox "The coupon type ('fix' vs 'float') hasn't been provided and will be set to 'fix'."
    ElseIf lType = "fix" Or lType = "fixed" Or lType = "fixed" Then
        floatCoupon = False
    ElseIf lType = "float" Then
        floatCoupon = True
    Else
        MsgBox "Unknown coupon type: " & typ & ", assuming fixed coupon."
        floatCoupon = False
    End If
    
End Property
Property Let PayFreq(pf As Integer)
    intPayFreq = pf
End Property
Property Let payDay(d As Integer)
    paymentDay = d
End Property
Property Let dcc(dcc As String)
    If dcc = "" Then
        intDCC = getDCC(ThisWorkbook.Sheets(strConfiguration).Range(strDefaultDCC))
    ElseIf IsNumeric(dcc) Then
        intDCC = dcc
    Else
        intDCC = getDCC(dcc)
    End If
End Property
Property Let bdc(bdc As String)
    If bdc = "" Then strBDC = ThisWorkbook.Sheets(strConfiguration).Range(strDefaultBDC) Else strBDC = bdc   'modified following
End Property
Property Let isAmort(am As Boolean)
    blnAmort = am
End Property
Property Let coupon(c As Double)
    dblCoupon = c
    prolongationCoupon = c
End Property
Property Let AmortScheme(ref As String)
    Select Case ref
    Case "linear", "Linear"
        varAmortScheme = "linear"
        blnAmort = True
    Case ""
        varAmortScheme = ""
        blnAmort = False
    Case Else
        varAmortScheme = ThisWorkbook.Sheets(strConfiguration).Range(ref).value
        blnAmort = True
    End Select
End Property
Property Let SetAmortScheme(amo As Variant)
    varAmortScheme = amo
    If Not array_Empty(amo) Then
        blnAmort = True
    ElseIf amo <> "" Then
        blnAmort = True
    Else
        blnAmort = False
    End If
End Property
Property Let refDate(dt As Date)
    dtCur = dt
End Property
Property Let Margin(ma As Integer)
    intMargin = ma
End Property
Property Let NextPayDt(NP As Date)
    dtNextPay = NP
End Property
Property Get getName()
    getName = strLoanName
End Property
Property Get getAssetType()
    getAssetType = strAssetType
End Property
Property Get getPayFreq()
    getPayFreq = intPayFreq
End Property
Property Get getNominal()
    getNominal = dblStartNom
End Property
Property Get getNPV()
    getNPV = dblNPV
End Property
Property Get getDates()
    getDates = dtLoanDates
End Property
Property Get getNCFs()
    getNCFs = dblNCFs
End Property
Property Get getCFs()
    getCFs = dblCFs
End Property
Property Get getIRCFs()
    getIRCFs = dblIRCFs
End Property
Property Get getACFs()
    getACFs = dblACFs
End Property
Property Get getDFs()
    getDFs = dblDFs
End Property
Property Get getNominals()
    getNominals = dblNominals
End Property


'***********************************************************************
'               Methods
'***********************************************************************
Public Sub calcCF(curveDic As Dictionary, discCurveName As String, scenName As String)
' routines sets and calculates members dblNominals, dblIRCFs, dblACRs, dblCFs, dblNCFs
' 
' Arguments:
'   curveDic: dictionary with available curves used to determin discounting and forward curve
'   discCurveName: name of discount curve
'   scenName: name of the applicable rate curve scenario
'
' output / transformed variables::
'   dtLoanDates(): load dates
'   dblNominals(): Outstanding notionals a end of period, i.e. after amortization, is no amortization:
'   dblCoupons(): interest coupon rates: fix rate if instrument typ fix; else: last fixing a index 0 and forward rates (plus margin) for index i to n-1
'   dblIRCFs(): interest rate payments for period from i-1 to i, calculated as average nominal in period - calculated from dblNominals() - multiplied by cuupon rate (dblCoupons()) at beginning of period
'   dblACFs(): amorization payments in period from i-1 to i, last entry contains payment of remaining notional
'   dblCFs(): sum of dblACFs() and dblIRCFs()
'   dblNCFs(): discounted values of dblCFs()
'
' See Also:
'   ['mdl_CurveUtilities.getCurveName'][/VBA/mdl_CurveUtilities.getCurveName], 
'   [mdl_Utilities.dtGenerator][/VBA/mdl_Utilities.dtGenerator], 
'   ['clsRateCurve.getDF'][/VBA/clsRateCurve.getDF], 
'   [clsRateCurve.calcFwdRate][/VBA/clsRateCurve.calcFwdRate]
'

    Dim yf1 As Double, yf2 As Double, i As Integer, j As Integer
    Dim coupon As Double
    Dim discCurve As clsRateCurve, fwdCurve As clsRateCurve, dtHypDate As Date
    

    ' determine discount and forward curve
    
    If curveDic.Exists(getScenarioCurveName(discCurveName, scenName)) Then
        Set discCurve = curveDic(getScenarioCurveName(discCurveName, scenName))
    Else
        MsgBox "The chosen discount curve " & discCurveName & " does not exist for scenario " & scenName & ". Please double check"
        Exit Sub
    End If
    If forwardCurve <> "" Then ' forward curve given -> use it
        If Not curveDic.Exists(getScenarioCurveName(forwardCurve, scenName)) Then
            MsgBox "Curve " & forwardCurve & " does not exist for scenario " & scenName
            Exit Sub
        Else
            Set fwdCurve = curveDic(getScenarioCurveName(forwardCurve, scenName))
        End If
    Else ' forward curve not given -> use payment frequency to determine curve
        If curveDic.Exists(getCurveName(intPayFreq, scenName)) Then
            Set fwdCurve = curveDic(getCurveName(intPayFreq, scenName))
        Else
            If intPayFreq > 0 Then MsgBox "The required forward curve does not exist for instrument '" & strLoanName & "'. The chosen discount curve will be used instead."
            Set fwdCurve = discCurve
        End If
    End If
    
    If dtLastDate = "00.00.00" Then dtHypDate = dtCur Else dtHypDate = dtLastDate
    dtLoanDates = dtGenerator(dtHypDate, dtMat, intPayFreq, intDCC, strBDC, 0, paymentDay)
    dblNPV = 0#
    ReDim dblNominals(UBound(dtLoanDates, 1))
    ReDim dblCoupons(UBound(dtLoanDates, 1))
    ReDim dblIRCFs(UBound(dtLoanDates, 1))
    ReDim dblACFs(UBound(dtLoanDates, 1))
    ReDim expiringNotional(UBound(dtLoanDates, 1))
    
    ReDim dblCFs(UBound(dtLoanDates, 1))
    ReDim dblNCFs(UBound(dtLoanDates, 1))
    ReDim dblDFs(UBound(dtLoanDates, 1))
    
    'determine nominals
    If blnAmort Then
        dblNominals(0) = dblStartNom
        If array_Empty(varAmortScheme) Then
            If varAmortScheme = "linear" Then
                For i = 1 To UBound(dblNominals)
                    dblNominals(i) = dblNominals(i - 1) - dblOrigNom / UBound(dtLoanDates, 1)
                Next
            End If
        Else
            ' assume there is a given liquidation curve serving as amortizatios scheme; amortization amounts are date dependent
            ' assume annual amortization rate is given, assume monthly amortization rates are linear fractions of annual rates
            Dim dblAmort() As Double
            ReDim dblAmort(UBound(dblNominals))
            For i = 1 To (UBound(dblNominals))
                ' get rate from schedule
                    For j = 1 To UBound(varAmortScheme, 2)
                        If WorksheetFunction.yearfrac(dtCur, dtLoanDates(i, 0), intDCC) * 12 / intPayFreq <= j Then         ' ToDo: check if number of month is adequately calculated
                            dblAmort(i) = varAmortScheme(1, j) * dblNominals(0)
                            Exit For
                        End If
                        dblAmort(i) = varAmortScheme(1, UBound(varAmortScheme, 2)) * dblNominals(0)                         ' flat extrapolation
                    Next
                ' calc amort amount
                'dblAmort(i) = dblStartNom
                dblNominals(i) = Application.WorksheetFunction.Max(dblNominals(i - 1) - dblAmort(i), 0)
                
            Next
        End If
    Else
        For i = 0 To UBound(dblNominals)
            dblNominals(i) = dblStartNom
        Next
    End If
    Dim nFixedFloatCoupons As Integer ' number of float coupons that are already fixed
    If forwardCurve = "" Then
        nFixedFloatCoupons = 1 ' if forw
    Else
        nFixedFloatCoupons = (CInt(Left(Right(forwardCurve, 2), 1)) + intPayFreq - 1) \ intPayFreq
    End If
    coupon = dblCoupon ' For fixed coupon, this value remains constant. For floating coupon, value is updated for each period.
    Dim effectiveNotional As Double
    For i = 1 To UBound(dblCoupons)
        ' this is the period from (i-1) to i
        ' i.e. float coupons are fixed at i-1 (for period [i-1, i]), all payments are made at i
        ' determine float coupons
        If floatCoupon Then
            If i > nFixedFloatCoupons Then ' for i=1 use fixing
                yf1 = WorksheetFunction.yearfrac(dtCur, dtLoanDates(i - 1, 1), intDCC)
                yf2 = WorksheetFunction.yearfrac(dtCur, dtLoanDates(i, 1), intDCC)
                coupon = fwdCurve.calcFwdRate(yf1, yf2)  ' forward rate as of i-1 from t(i-1) to t(i) based on fixing dates
            End If
            coupon = WorksheetFunction.Max(coupon, floatFloor)
            coupon = coupon + intMargin / 10000
            If hasVariableMargin Then
                If UBound(variableMargin) >= i - 1 Then
                    coupon = coupon + variableMargin(i - 1)
                End If
            End If
        End If
        If useAverageNotional Then
            effectiveNotional = (dblNominals(i - 1) + dblNominals(i)) / 2#
        Else
            effectiveNotional = dblNominals(i - 1)
        End If
        dblIRCFs(i) = blnIsAsset * coupon * WorksheetFunction.yearfrac(dtLoanDates(i - 1, 1), dtLoanDates(i, 1), intDCC) * effectiveNotional
        expiringNotional(i) = blnIsAsset * (dblNominals(i - 1) - dblNominals(i))
        If includeAmortizationCF Then
            dblACFs(i) = expiringNotional(i)
        Else
            dblACFs(i) = 0#
        End If
        dblCFs(i) = dblIRCFs(i) + dblACFs(i)
        dblNCFs(i) = dblCFs(i) * discCurve.getDF(dtLoanDates(i, 2))
        dblDFs(i) = discCurve.getDF(dtLoanDates(i, 2))
        dblNPV = dblNPV + dblNCFs(i)
    Next
    If dblNominals(i - 1) > 0 Then
     ' treat payment of nominal at maturity as amortization
        expiringNotional(i - 1) = expiringNotional(i - 1) + blnIsAsset * dblNominals(i - 1)
        If includeAmortizationCF Then
            dblCFs(i - 1) = dblCFs(i - 1) + blnIsAsset * dblNominals(i - 1)
            dblNCFs(i - 1) = dblNCFs(i - 1) + blnIsAsset * dblNominals(i - 1) * discCurve.getDF(dtLoanDates(i - 1, 2))
            dblACFs(i - 1) = dblACFs(i - 1) + blnIsAsset * dblNominals(i - 1)
            dblNPV = dblNPV + dblNominals(i - 1) * discCurve.getDF(dtLoanDates(i - 1, 2)) * blnIsAsset
        End If
    End If
    
End Sub


Public Function printCF() As Variant()
' Function returns matrix containing vectors dtLoanDates(*,2), dbIRCFs(), dblNCFs(), dblCFs(), dblACFs(), dblNominals(),
' Matrix is used for printing the IR run of balance
Dim dblData() As Variant, i As Integer

    ReDim dblData(UBound(dblCFs), 5)
        dblData(0, 0) = dtCur
        dblData(0, 1) = 0
        dblData(0, 2) = dblNPV
        dblData(0, 3) = 0
        dblData(0, 4) = 0
        dblData(0, 5) = dblNominals(0)
    For i = 1 To UBound(dblCFs) ' start at i = 1 which is the first payment date beyond the portfolio date (at i=0 there is only the fixing of the applicable rate and the calculation of the applicable notional)
        dblData(i, 0) = dtLoanDates(i, 2)
        dblData(i, 3) = dblCFs(i)
        dblData(i, 2) = dblNCFs(i)
        dblData(i, 1) = dblIRCFs(i)
        dblData(i, 4) = dblACFs(i)
        dblData(i, 5) = dblNominals(i)
    Next
    dblData(UBound(dblData), 5) = 0 ' set remaining notional to 0 at end
    printCF = dblData
End Function


Public Function getCF() As clsCashFlow()
' Function: getCF
'
    Dim cf() As clsCashFlow, i As Integer

    ReDim cf(UBound(dblCFs) - 1)
    For i = 1 To UBound(dblCFs) ' start at i = 1 which is the first payment date beyond the portfolio date (at i=0 there is only the fixing of the applicable rate and the calculation of the applicable notional)
        Set cf(i - 1) = CreateCashFlow(dtLoanDates(i, 2), dblIRCFs(i), dblACFs(i), dblDFs(i))
    Next
    getCF = cf
End Function


Public Sub addNii(ByRef nii() As Double, buckets() As Date, simulationHorizon As Date, curveDic As Dictionary, discCurveName As String, scenName As String)
' subroutine for NII
'
' Arugments:
'   nii: double
'   buckets: date
'   simulationHorizon: date
'   curveDic: Dictionary
'   discCurveName: String
'   scenName: String
'
' Returns:
'   Nothing

    Dim result As New Dictionary
    Dim i As Integer
    ' Start with contractual cash flows
    For i = 1 To UBound(dblCFs)
        result.Add dtLoanDates(i, 2), dblIRCFs(i)
    Next i
    
    ' Now add prolongation cash flows for each amortization
    Dim amNotional As Double
    Dim d As Date, d2 As Date, d3 As Date, d4 As Date, d5 As Date
    Dim newCoupon As Double
    
    Dim baseFwdCurve As clsRateCurve
    Dim scenFwdCurve As clsRateCurve
        
    Dim baseFwdCurveName As String
    Dim scenFwdCurveName As String
    If forwardCurve <> "" Then ' forward curve given -> use it
        baseFwdCurveName = getScenarioCurveName(forwardCurve, "")
        scenFwdCurveName = getScenarioCurveName(forwardCurve, scenName)
    ElseIf intPayFreq <> 0 Then
        baseFwdCurveName = getCurveName(intPayFreq, "")
        scenFwdCurveName = getCurveName(intPayFreq, scenName)
    Else
        baseFwdCurveName = getScenarioCurveName(discCurveName, "")
        scenFwdCurveName = getScenarioCurveName(discCurveName, scenName)
    End If
    If Not curveDic.Exists(baseFwdCurveName) Then
        MsgBox "Curve " & baseFwdCurveName & " does not exist."
        Exit Sub
    End If
    If Not curveDic.Exists(scenFwdCurveName) Then
        MsgBox "Curve " & scenFwdCurveName & " does not exist."
        Exit Sub
    End If
    Set baseFwdCurve = curveDic(baseFwdCurveName)
    Set scenFwdCurve = curveDic(scenFwdCurveName)
    Dim yf As Double
    Dim interest As Double
    Dim prolongationPayFreq As Integer
    If intPayFreq > 0 Then
        prolongationPayFreq = intPayFreq
    Else
        prolongationPayFreq = 1
    End If
    Dim yf1 As Double
    Dim yf2 As Double
    For i = 1 To UBound(expiringNotional)
        amNotional = expiringNotional(i)
        d = dtLoanDates(i, 2)
        If amNotional <> 0 Then
            Do While d < simulationHorizon
                yf = (d - dtCur) / 365#
                If Not floatCoupon Then
                    newCoupon = prolongationCoupon + scenFwdCurve.calcFwdRate(yf, yf + prolongationDuration) - baseFwdCurve.calcFwdRate(0#, prolongationDuration)
                End If
                d2 = d + Round(prolongationDuration * 365# + 0.5, 0) ' round up to next integer date
                d3 = d

                d4 = DateAdd("m", prolongationPayFreq, d3)
                Do
                    d5 = DateAdd("m", prolongationPayFreq, d3)
                    d4 = WorksheetFunction.Min(d5, d2)
                    If floatCoupon Then
                        yf1 = WorksheetFunction.yearfrac(dtCur, d3, intDCC)
                        yf2 = WorksheetFunction.yearfrac(dtCur, d4, intDCC)
                        newCoupon = WorksheetFunction.Max(scenFwdCurve.calcFwdRate(yf1, yf2), floatFloor) + intMargin / 10000#
                    End If
                    interest = amNotional * WorksheetFunction.yearfrac(d3, d4, intDCC) * newCoupon
                    result(d4) = result(d4) + interest
                    d3 = d4
                Loop While d5 < d2
                
                d = d2
                
'                    If i > 1 Then ' for i=1 use fixing
'                        yf1 = WorksheetFunction.yearfrac(dtCur, dtLoanDates(i - 1, 1), intDCC)
'                        yf2 = WorksheetFunction.yearfrac(dtCur, dtLoanDates(i, 1), intDCC)
'                        coupon = fwdCurve.calcFwdRate(yf1, yf2)  ' forward rate as of i-1 from t(i-1) to t(i) based on fixing dates
'                    End If
'                    coupon = WorksheetFunction.Max(coupon, floatFloor)
'                    coupon = coupon + intMargin / 10000
'                    If hasVariableMargin Then
'                        If UBound(variableMargin) >= i - 1 Then
'                            coupon = coupon + variableMargin(i - 1)
'                        End If
'                    End If
                

        
            Loop
        End If
    Next i
    
    Set result = funcSortKeysAsc(result)
    Dim bucketIndex As Integer
    bucketIndex = 0
    
    
    Dim niiDate As Variant

    For Each niiDate In result.Keys
        Do While bucketIndex < UBound(buckets) - 1 And buckets(bucketIndex) < niiDate
            bucketIndex = bucketIndex + 1
        Loop
        nii(bucketIndex) = nii(bucketIndex) + result(niiDate)
    Next niiDate
End Sub


