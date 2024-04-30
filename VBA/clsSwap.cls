VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ****************** Variables *****************************************
Private strSwapName As String
Private strBalanceItem As String

Private dtMat As Date
Private dblStartNom As Double
Private dblOrigNom As Double
Private intFixPayFreq As Integer
Private intFloatPayFreq As Integer
Private intFixDCC As Integer               ' day count convention
Private intFloatDCC As Integer               ' day count convention
Private strBDC As String            ' business day convention
Private blnAmort As Boolean         ' Amortizing yes / no
Private dblFixCoupon As Double         ' current coupon if float, coupon if fix
Private dblFloatCoupon As Double
Private varAmortScheme As Variant
Private dtCur As Date
Private intMargin As Integer
Private dtNextPay As Date
Private dblLastFixing As Double

Private dtLastFloatDate As Date
Private dtLastFixDate As Date
Private strAssetType As String
Private isPayerSwap As Boolean
Public floatFloor As Double         ' floor for floating coupon

Private fixLeg As clsLoan
Private floatLeg As clsLoan

' calculated data
Private dtSwapDates() As Date
Private dblNominals() As Long
Private dblCoupons() As Double
Private dblNPV As Double
Private dblCFs() As Double
Private dblNCFs() As Double

'***********************************************************************
'               Properties
'***********************************************************************
Property Let name(name As String)
    strSwapName = name
End Property

Property Let BalanceItem(name As String)
    strBalanceItem = name
End Property
Property Let maturity(mat As Date)
    dtMat = mat
End Property
Property Let Nominal(n As Double)
    dblStartNom = n
End Property
Property Let OrigNominal(n As Double)
    dblOrigNom = n
End Property

Property Let FixPayFreq(pf As Integer)
    intFixPayFreq = pf
End Property
Property Let FloatPayFreq(pf As Integer)
    intFloatPayFreq = pf
End Property
Property Let Fixdcc(dcc As String)
    intFixDCC = getDCC(dcc)
End Property
Property Let Floatdcc(dcc As String)
    intFloatDCC = getDCC(dcc)
End Property
Property Let bdc(bdc As String)
    strBDC = bdc
End Property
Property Let isAmort(am As Boolean)
    blnAmort = am
End Property
Property Let FixCoupon(c As Double)
    dblFixCoupon = c
End Property
Property Let floatCoupon(c As Double)
    dblFloatCoupon = c
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
        varAmortScheme = ThisWorkbook.Sheets("Configuration").Range(ref).value
        blnAmort = True
    End Select
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
Property Let payReceive(payReceive As String)
    isPayerSwap = LCase(payReceive) = "pay"
End Property
Property Let lastFixing(lfix As Double)
    dblLastFixing = lfix
End Property
Property Let LastFloatDate(start As Date)
    dtLastFloatDate = start
End Property
Property Let LastFixDate(start As Date)
    dtLastFixDate = start
End Property
Property Let assetType(t As String)
    strAssetType = t
End Property
Property Get getAssetType()
    getAssetType = strAssetType
End Property

'***********************************************************************
'               Methods
'***********************************************************************

Private Sub Class_Initialize()
    floatFloor = -1E+99
End Sub


Public Sub InitLegs()
Dim misData As String, txt As String

    misData = checkData()
    If misData <> "" Then
        txt = "Please check the available data. The following Information is missing: " & vbCrLf & misData
        MsgBox txt
    Else
        Set floatLeg = New clsLoan
                floatLeg.name = strSwapName & "_float"
                floatLeg.Nominal = dblStartNom
                floatLeg.OrigNominal = dblOrigNom
                floatLeg.maturity = dtMat
                floatLeg.PayFreq = intFloatPayFreq
                floatLeg.SetAmortScheme = varAmortScheme
                floatLeg.dcc = intFloatDCC
                floatLeg.bdc = strBDC
                floatLeg.typ = "float"
                floatLeg.coupon = dblLastFixing
                floatLeg.Margin = intMargin
                floatLeg.NextPayDt = dtNextPay
                floatLeg.refDate = dtCur
                floatLeg.isAsset = True                       ' swap is regarded as asset, net payments indicate if swap is a liability or asset
                floatLeg.floatFloor = floatFloor
        Set fixLeg = New clsLoan
                fixLeg.refDate = dtCur
                fixLeg.name = strSwapName & "_fix"
                fixLeg.Nominal = dblStartNom
                fixLeg.OrigNominal = dblOrigNom
                fixLeg.maturity = dtMat
                fixLeg.PayFreq = intFixPayFreq
                fixLeg.SetAmortScheme = varAmortScheme
                fixLeg.dcc = intFixDCC
                fixLeg.bdc = strBDC
                fixLeg.typ = "fix"
                fixLeg.coupon = dblFixCoupon
                fixLeg.Margin = intMargin
                fixLeg.NextPayDt = dtNextPay
                fixLeg.refDate = dtCur
                fixLeg.isAsset = True                       ' swap is regarded as asset, net payments indicate if swap is a liability or asset
    End If

End Sub
Public Function printCF() As Variant()
    Dim dblData() As Variant, i As Integer, sign As Double
    If isPayerSwap Then
        sign = 1#
    Else
        sign = -1#
    End If
    
    ReDim dblData(UBound(fixLeg.getCFs), 5)
        dblData(0, 0) = dtCur
        dblData(0, 1) = 0
        dblData(0, 2) = sign * (-fixLeg.getNPV() + floatLeg.getNPV())
        dblData(0, 3) = 0
        dblData(0, 4) = 0
        dblData(0, 5) = fixLeg.getNominals(0)
    For i = 1 To UBound(fixLeg.getCFs) ' start at i = 1 which is the first payment date beyond the portfolio date (at i=0 there is only the fixing of the applicable rate and the calculation of the applicable notional)
        dblData(i, 0) = fixLeg.getDates(i, 2)
        dblData(i, 3) = sign * (-fixLeg.getCFs(i) + floatLeg.getCFs(i))
        dblData(i, 2) = sign * (-fixLeg.getNCFs(i) + floatLeg.getNCFs(i))
        dblData(i, 1) = sign * (-fixLeg.getIRCFs(i) + floatLeg.getIRCFs(i))
        dblData(i, 4) = sign * (-fixLeg.getACFs(i) + floatLeg.getACFs(i))
        dblData(i, 5) = fixLeg.getNominals(i)
    Next
    printCF = dblData
End Function

Public Function getCF() As clsCashFlow()
    Dim cf() As clsCashFlow, i As Integer, sign As Double

    ReDim cf(UBound(fixLeg.getCFs()) - 1)
    
    If isPayerSwap Then
        sign = 1#
    Else
        sign = -1#
    End If
    
    Dim d As Variant
    Dim df As Double
    For i = 1 To UBound(fixLeg.getCFs())
        d = fixLeg.getDates(i, 2)
        df = fixLeg.getDFs(i)
        Set cf(i - 1) = CreateCashFlow((d), sign * (-fixLeg.getIRCFs(i) + floatLeg.getIRCFs(i)), sign * (-fixLeg.getACFs(i) + floatLeg.getACFs(i)), df)
    Next
            

    getCF = cf
End Function

Public Function getPayLeg() As clsLoan
    Dim result As clsLoan
    If isPayerSwap Then
        Set result = fixLeg
    Else
        Set result = floatLeg
    End If
    Set getPayLeg = result
End Function

Public Function getReceiveLeg() As clsLoan
    Dim result As clsLoan
    If isPayerSwap Then
        Set result = floatLeg
    Else
        Set result = fixLeg
    End If
    Set getReceiveLeg = result
End Function

Public Sub calcCF(curveDic As Dictionary, disCurveName As String, scenName As String)

    fixLeg.calcCF curveDic, disCurveName, scenName
    floatLeg.calcCF curveDic, disCurveName, scenName

End Sub

Private Function checkData() As String
Dim misData As String, i As Integer
    i = 1
    If dtMat = "00.00.00" Then
        misData = i & ". Maturity Date " & vbCrLf
        i = i + 1
    End If
    If dblStartNom = 0# Then
        misData = i & ". Current Nominal " & vbCrLf
        i = i + 1
    End If
    If dblOrigNom = 0# Then
        misData = i & ". Original Nominal " & vbCrLf
        i = i + 1
    End If
    If intFixPayFreq = 0 Then
        misData = i & ". Pay Frequency " & vbCrLf
        i = i + 1
    End If
    If intFloatPayFreq = 0 Then
        misData = i & ". Pay Frequency " & vbCrLf
        i = i + 1
    End If
    If dblFixCoupon = 0# Then
        misData = i & ". Fix Coupon Rate " & vbCrLf
        i = i + 1
    End If
    If dblFloatCoupon = 0# Then
        misData = i & ". Float Coupon Rate " & vbCrLf
        i = i + 1
    End If
    If dtCur = "00.00.00" Then                        ' To Do klären!
        misData = i & ". Current Date " & vbCrLf
        i = i + 1
    End If
    
    checkData = misData

End Function

