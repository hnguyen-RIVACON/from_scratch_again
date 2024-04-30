Attribute VB_Name = "mdl_AssetUtilities"
Option Explicit

'************************* Instrument Factory *****************************************
Public Function InstrumentFactory(curves As Dictionary, Optional scen As String = "default") As Scripting.Dictionary

    Dim assets As New Scripting.Dictionary, assetsind As New Scripting.Dictionary, assetType As New Scripting.Dictionary
    Dim rng As Range, rngAsset As Range, i As Integer, index As Integer
    Dim anzAssets As Integer
    Dim loan As clsLoan
    Dim Swap As clsSwap
    
    If scen = "default" Then
        Set rng = ThisWorkbook.Sheets(strConfiguration).Range(strCurrentPosition)
    Else
        Set rng = ThisWorkbook.Sheets(strConfiguration).Range(strPositions)
    End If
    index = 1
    For i = 1 To rng.Rows.count
        Select Case rng.Cells(i, 1).value
        Case "Retail", "Wholesale", "RetailCommitment", "Leasing"
            RetailWholesaleFactory assets, assetsind, assetType, index, rng.Cells(i, 1).value, curves
        Case "ABSRetainedNotes", "ABSSynthLiabilities"
            ABSFactory assets, assetsind, assetType, index, rng.Cells(i, 1).value
        Case "IntercompanyLoans"
            IntercompanyLoansFactory assets, assetsind, assetType, index
        Case "Deposit"
            DepositFactory assets, assetsind, assetType, index
        Case "Cash", "ECBCash", "ECBTender"
            CashFactory assets, assetsind, assetType, index, rng.Cells(i, 1).value
        Case "Swap", "swap"
            SwapFactory assets, assetsind, assetType, index
        Case Else
            MsgBox "The instrument " & rng.Cells(i, 2) & " has not been setup yet."
        End Select
    Next
    
    Set InstrumentFactory = assets
    Set rng = Nothing
    Set assets = Nothing
End Function

'***************************** Position specific factories ******************************
Sub RetailWholesaleFactory(assets As Dictionary, assetsind As Dictionary, assetType As Dictionary, i As Integer, strType As String, curves As Dictionary)
' content:
' Sub creates for each instrument of the position type an element of the loan class.
' When reading the input data, the specific setup of the position type within the data sheet is considered
' Parameters: assets: Dictionary: dictionary to which each instrument will be added, i.e. the parameter is manipulated/extended by this sub

    Dim rngAsset As Range, loan As clsLoan, rfDate As Date, row As Range
    Dim fraction As Double

    Set rngAsset = ThisWorkbook.Sheets(strWholesaleData).Range(strType & "_DATA")
    rfDate = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    Dim subsidized As Boolean
    Dim maturity As Date
    For Each row In rngAsset.Rows
            Set loan = New clsLoan
            loan.refDate = rfDate
            loan.name = row.Cells(1, 1)
            loan.isAsset = True
            fraction = 1#
            If Not IsEmpty(row.Cells(1, 11)) Then
                fraction = row.Cells(1, 11).value
            End If
            If Not IsNumeric(row.Cells(1, 2)) Then
                loan.Nominal = fraction * Abs(getRetailNominal(row.Cells(1, 1)))
            Else
                loan.Nominal = fraction * Abs(row.Cells(1, 2))
            End If
            loan.OrigNominal = loan.getNominal
            loan.bookValue = loan.getNominal
            maturity = DateAdd("M", 180, rfDate)
            loan.maturity = maturity
            loan.coupon = row.Cells(1, 3)
            loan.typ = row.Cells(1, 4)
            loan.AmortScheme = row.Cells(1, 5)
            loan.PayFreq = 1
            loan.dcc = row.Cells(1, 7)
            loan.forwardCurve = row.Cells(1, 8).value
            loan.bdc = ""
            loan.Margin = row.Cells(1, 9)
            loan.prolongationDuration = getAvgDuration(row.Cells(1, 5), rfDate, maturity)
            If InStr(row.Cells(1, 1).value, "Wholesale") Then
                subsidized = row.Cells(1, 10).value
                If subsidized Then
                    Dim avgRate As Double
                    avgRate = row.Cells(1, 3).value
                    Dim spreads() As Double
                    ReDim spreads(0 To 2)
                    
                    Dim fwdCurve As clsRateCurve
                    Set fwdCurve = curves(getScenarioCurveName(row.Cells(1, 8).value, "")) ' fwd curve in base scenario used to calculate spread
                    spreads(0) = 0#
                    spreads(1) = avgRate - fwdCurve.calcFwdRate(1 / 12, 2 / 12)
                    spreads(2) = avgRate - fwdCurve.calcFwdRate(2 / 12, 3 / 12)
                    loan.varMargin = spreads
                End If
            End If
            loan.NextPayDt = DateAdd("M", loan.getPayFreq, rfDate)
            loan.assetType = row.Cells(1, 1).value
            assets.Add row.Cells(1, 1).value, loan
            assetsind.Add i, row.Cells(1, 1).value
            i = i + 1
    Next
    assetType.Add strType, i - 1

End Sub
 
Sub IntercompanyLoansFactory(assets As Dictionary, assetsind As Dictionary, assetType As Dictionary, i As Integer)
' content:
' Sub creates for each instrument of the position type an element of the loan class.
' When reading the input data, the specific setup of the position type within the data sheet is considered
' Parameters: assets: Dictionary: dictionary to which each instrument will be added, i.e. the parameter is manipulated/extended by this sub

Dim rngAsset As Range, loan As clsLoan, rfDate As Date, row As Range

    Set rngAsset = ThisWorkbook.Sheets(strIntercompanyLoanData).Range(strIntercompanyLoansData)
    rfDate = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    
    For Each row In rngAsset.Rows
        If row.Cells(1, 7) > rfDate Then
            Set loan = New clsLoan
            loan.refDate = rfDate
            loan.name = row.Cells(1, 1) & "_" & row.Cells(1, 2)
            loan.isAsset = False
            loan.Nominal = -row.Cells(1, 9)
            loan.bookValue = -loan.getNominal
            loan.OrigNominal = loan.getNominal
            loan.maturity = row.Cells(1, 7)
            loan.coupon = row.Cells(1, 16)
            loan.typ = row.Cells(1, 11)
            loan.AmortScheme = ""
            loan.PayFreq = Left(row.Cells(1, 14), 1)
            loan.prolongationDuration = Left(row.Cells(1, 14), 1) / 12#
            loan.dcc = row.Cells(1, 12)
            loan.bdc = ""
            loan.Margin = row.Cells(1, 15)
            loan.assetType = "IntercompanyLoan"
            assets.Add loan.getName, loan
            assetsind.Add i, row.Cells(1, 1)
            i = i + 1
        End If
    Next
    assetType.Add "IntercompanyLoans", i - 1

End Sub
Sub ABSFactory(assets As Dictionary, assetsind As Dictionary, assetType As Dictionary, i As Integer, strType As String)
' content:
' Sub creates for each instrument of the position type an element of the loan class.
' When reading the input data, the specific setup of the position type within the data sheet is considered
' Parameters: assets: Dictionary: dictionary to which each instrument will be added, i.e. the parameter is manipulated/extended by this sub

Dim rngAsset As Range, loan As clsLoan, rfDate As Date, row As Range

    rfDate = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    Set rngAsset = ThisWorkbook.Sheets(strABSData).Range(strType & "_DATA")
    Dim maturity As Date
    For Each row In rngAsset.Rows
        If row.Cells(1, 4) <> "" And Abs(row.Cells(1, 3)) > 0 And row.Cells(1, 8) > rfDate Then
            Set loan = New clsLoan
            loan.refDate = rfDate
            loan.Nominal = Abs(row.Cells(1, 3))
            loan.bookValue = loan.getNominal()
            If Right(strType, 5) <> "Notes" Then
                loan.isAsset = False
                loan.name = row.Cells(1, 1) & "_" & row.Cells(1, 2) & "_Liabs"
                loan.bookValue = -loan.getNominal()
            Else
                loan.isAsset = True
                loan.name = row.Cells(1, 1) & "_" & row.Cells(1, 2) & "_Asset"
                loan.bookValue = loan.getNominal()
            End If
            loan.OrigNominal = loan.getNominal
            If row.Cells(1, 9) = "" Then
                maturity = row.Cells(1, 8)
            Else
                maturity = row.Cells(1, 9)
            End If
            loan.maturity = maturity
            loan.coupon = row.Cells(1, 5)
            loan.typ = row.Cells(1, 4)
            loan.AmortScheme = row.Cells(1, 14)
            If row.Cells(1, 4) = "Float" Then
                loan.PayFreq = Left(row.Cells(1, 12), 1)
            Else
                loan.PayFreq = 1
            End If
            If row.Cells(1, 6) <> "" Then
                loan.payDay = row.Cells(1, 6).value
            End If
            loan.dcc = row.Cells(1, 7)
            loan.bdc = ""
            loan.Margin = row.Cells(1, 13)
            loan.assetType = "ABS"
            loan.prolongationDuration = getAvgDuration(row.Cells(1, 14), rfDate, maturity)
            assets.Add loan.getName, loan
            assetsind.Add i, row.Cells(1, 1)
            i = i + 1
        End If
    Next
    assetType.Add strType, i - 1
    
    'If strRange = strABSAssetsData Then ABSFactory assets, assetsind, assettype, i, strABSLiabilitiesData

End Sub
Sub DepositFactory(assets As Dictionary, assetsind As Dictionary, assetType As Dictionary, i As Integer, tp As InstrumentType, typeName As String)
' content:
' Sub creates for each instrument of the position type an element of the loan class.
' When reading the input data, the specific setup of the position type within the data sheet is considered
' Parameters: assets: Dictionary: dictionary to which each instrument will be added, i.e. the parameter is manipulated/extended by this sub

Dim rngAsset As Range, loan As clsLoan, rfDate As Date, row As Range, j As Integer

    Set rngAsset = ThisWorkbook.Sheets(strDepositData).Range(strDepositsData)
    rfDate = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    j = 1
    Dim duration As Double
    Dim flex As Boolean
    flex = (tp = DepositFlex)
    Dim flexRow As Boolean
    For Each row In rngAsset.Rows
        flexRow = LCase(Left(row.Cells(1, 1), 9)) = "tagesgeld"
        If (flex = flexRow) And row.Cells(1, 3) = "Y" And ((row.Cells(1, 9) <> "-" And row.Cells(1, 9) > rfDate) Or flexRow) Then
            Set loan = New clsLoan
            loan.refDate = rfDate
            loan.name = row.Cells(1, 1) & "_" & j
            loan.isAsset = False
            loan.Nominal = row.Cells(1, 10)
            loan.OrigNominal = loan.getNominal
            loan.bookValue = -loan.getNominal
            loan.coupon = row.Cells(1, 4)
            If flex Then
                loan.AmortScheme = "Tagesgeld_LC"
                loan.prolongationDuration = getAvgDuration("Tagesgeld_LC", rfDate, rfDate)
                loan.maturity = DateAdd("m", 180, rfDate)
                loan.typ = "fix"
            Else
                loan.AmortScheme = ""
                loan.prolongationDuration = (CDate(row.Cells(1, 9)) - rfDate) / 365#
                loan.maturity = row.Cells(1, 9)
            loan.typ = row.Cells(1, 5)
            End If
            loan.PayFreq = 1
            loan.dcc = ""
            loan.bdc = ""
            loan.Margin = 0
            loan.assetType = "Deposit"
            assets.Add loan.getName, loan
            j = j + 1
            assetsind.Add i, row.Cells(1, 1)
            i = i + 1
        End If
    Next
    assetType.Add typeName, i - 1

End Sub

Sub CashFactory(assets As Dictionary, assetsind As Dictionary, assetType As Dictionary, Optional i As Integer, Optional strType As String)
' content:
' Sub creates for each instrument of the position type an element of the loan class.
' When reading the input data, the specific setup of the position type within the data sheet is considered
' Parameters: assets: Dictionary: dictionary to which each instrument will be added, i.e. the parameter is manipulated/extended by this sub

Dim rngAsset As Range, loan As clsLoan, rfDate As Date, row As Range

    Set rngAsset = ThisWorkbook.Sheets(strDepositData).Range(strType & "_DATA")
    rfDate = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)

    For Each row In rngAsset.Rows
        Set loan = New clsLoan
        loan.refDate = rfDate
        loan.name = row.Cells(1, 1)
        loan.Nominal = row.Cells(1, 2)
        loan.maturity = row.Cells(1, 3)
        loan.LastDate = row.Cells(1, 4)
        loan.OrigNominal = loan.getNominal
        loan.bookValue = loan.getNominal
        If Right(row.Cells(1, 1), 6) = "Tender" Or Right(row.Cells(1, 1), 6) = "tender" Then
            loan.coupon = row.Cells(1, 5)
            loan.isAsset = False
            loan.assetType = "Tender"
            loan.PayFreq = 0
            loan.bookValue = -loan.getNominal()
            loan.prolongationDuration = (row.Cells(1, 3).value - row.Cells(1, 4).value) / 365#
        Else
            loan.coupon = 0
            loan.prolongationCoupon = row.Cells(1, 5).value
            loan.prolongationDuration = 1# / 365#
            loan.isAsset = True
            loan.assetType = "Cash"
            loan.PayFreq = 1
            loan.bookValue = loan.getNominal()
        End If
        loan.typ = "fix"
        loan.AmortScheme = ""
        
        loan.dcc = row.Cells(1, 6)
        loan.bdc = ""
        loan.Margin = 0
        assets.Add loan.getName, loan
        assetsind.Add i, row.Cells(1, 1)
        i = i + 1
    Next
    assetType.Add strType, i - 1

End Sub
Sub SwapFactory(assets As Dictionary, assetsind As Dictionary, assetType As Dictionary, i As Integer, Optional split As Boolean = False, Optional pay As Boolean = True, Optional alm As Boolean = True)
' content:
' Sub creates for each instrument of the position type an element of the loan class.
' When reading the input data, the specific setup of the position type within the data sheet is considered
' Parameters: assets: Dictionary: dictionary to which each instrument will be added, i.e. the parameter is manipulated/extended by this sub
' if splitSwaps=true, the two legs are created as separate instruments (loans)

Dim rngAsset As Range, Swap As clsSwap, rfDate As Date, row As Range, j As Integer

    Set rngAsset = ThisWorkbook.Sheets(strSwapData).Range(strSwapsData)
    rfDate = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    Dim leg As clsLoan
    Dim swapIsAlm As Boolean

    For j = 1 To rngAsset.Rows.count
        If rngAsset.Cells(j, 4) > rfDate Then
            Set Swap = New clsSwap
            If rngAsset.Cells(j, 9) <> "" Then
                Swap.assetType = "ABS Swap"
                swapIsAlm = False
            Else
                Swap.assetType = "ALM Swap"
                swapIsAlm = True
            End If
            If (Not split Or (swapIsAlm = alm)) Then
                Swap.name = rngAsset.Cells(j, 1)
                Swap.Nominal = Abs(rngAsset.Cells(j, 2))
                Swap.OrigNominal = rngAsset.Cells(j, 3)
                Swap.maturity = rngAsset.Cells(j, 4)
                Swap.floatCoupon = rngAsset.Cells(j + 1, 8) / 100
                Swap.FixCoupon = rngAsset.Cells(j, 5) / 100
                Swap.Floatdcc = rngAsset.Cells(j + 1, 6)
                Swap.Fixdcc = rngAsset.Cells(j, 6)
                Swap.FixPayFreq = Left(rngAsset.Cells(j + 1, 7), 1)
                Swap.FloatPayFreq = Left(rngAsset.Cells(j + 1, 7), 1)
                If Not IsEmpty(rngAsset.Cells(j + 1, 11)) Then
                    Swap.floatFloor = rngAsset.Cells(j + 1, 11).value
                End If
                Swap.AmortScheme = rngAsset.Cells(j, 9)
                Swap.bdc = ""
                Swap.Margin = rngAsset.Cells(j + 1, 10)
                Swap.refDate = rfDate
                If rngAsset.Cells(j, 3) < 0 Then Swap.payReceive = "pay" Else Swap.payReceive = "receive"
                Swap.lastFixing = rngAsset.Cells(j + 1, 8) / 100
                Swap.InitLegs
                If Not split Then
                    assets.Add rngAsset.Cells(j, 1).value, Swap
                ElseIf pay Then
                    Set leg = Swap.getPayLeg()
                    leg.isAsset = False
                    leg.includeAmortizationCF = False
                    If leg.floatCoupon Then
                        leg.prolongationDuration = Left(rngAsset.Cells(j + 1, 7), 1) / 12#
                    Else
                        leg.prolongationDuration = (rngAsset.Cells(j, 4) - rfDate) / 365#
                    End If
                    assets.Add rngAsset.Cells(j, 1).value & "_Pay", leg
                Else
                    Set leg = Swap.getReceiveLeg()
                    leg.isAsset = True
                    leg.includeAmortizationCF = False
                    If leg.floatCoupon Then
                        leg.prolongationDuration = Left(rngAsset.Cells(j + 1, 7), 1) / 12#
                    Else
                        leg.prolongationDuration = (rngAsset.Cells(j, 4) - rfDate) / 365#
                    End If
                    assets.Add rngAsset.Cells(j, 1).value & "_Receive", leg
                End If
                assetsind.Add i, rngAsset.Cells(j, 1).value
                j = j + 1
                i = i + 1
            End If
        End If
    Next
    If split Then
        If pay And alm Then
            assetType.Add "ALMSwapPayLeg", i - 1
        ElseIf pay And Not alm Then
            assetType.Add "ABSSwapPayLeg", i - 1
        ElseIf Not pay And alm Then
            assetType.Add "ALMSwapReceiveLeg", i - 1
        ElseIf Not pay And Not alm Then
            assetType.Add "ABSSwapReceiveLeg", i - 1
        End If
    Else
        assetType.Add "Swap", i - 1
    End If

End Sub

'***************************** calculate IR cash flows for all instruments ******************************
Sub calcIRCF(assets As Dictionary, curveDic As Dictionary, disCurveName As String, scen As String)
' Content:
' CashFlows for all assets are calculated. CFs include ir as well as amortization cashflows
' NPV and discounted CFs are calculated as well
' calculated data is stored within the rate curve object
' Parameter: Assets - dictionary with all assets for which ir cf are calculated, key name corresponds each asset's name
' Parameter: Curve - clsRateCurve for discouting and forward calculation

Dim c As Variant
    For Each c In assets.Keys
        If scen <> "default" Then assets(c).calcCF curveDic, disCurveName, scen
    Next
End Sub

'***************************** print cash flows for all instruments ******************************
Sub printIRCF(assets As Dictionary)
Dim c As Variant, intAnz As Integer, intCFdt As Integer, dblNumbers() As Double, data As Variant, dataExist As Variant
Dim p As New Scripting.Dictionary, strRef As String, rng As Range

    Sheets(strIRBalance).Activate
    Cells.Clear
    Cells(2, 1) = "Date"
    Cells(2, 1).AddComment ("All cashflows occuring after the reference date are included")
    strRef = "A3"
    Set rng = ThisWorkbook.Sheets(strIRBalance).Range(strRef)
    ' setUp output data structure (Matrix: Column 1: dates, column 2: CF, column 3: discounted CF, row 1: current date and NPV (in 3rd column)
    ReDim dblNumber(assets.count * 5 - 1)
    intAnz = 0
    For Each c In assets.Keys
        Application.ScreenUpdating = False
            ThisWorkbook.Sheets(strIRBalance).Cells(rng.row - 2, rng.column + 1 + intAnz * 5).Activate
            ActiveCell = c  '.value
            Range(ActiveCell, Cells(ActiveCell.row, ActiveCell.column + 4)).Merge
            Cells(ActiveCell.row + 1, ActiveCell.column) = "IR CF"
            Cells(ActiveCell.row + 1, ActiveCell.column + 1) = "Disc. CF"
            Cells(ActiveCell.row + 1, ActiveCell.column + 2) = "CF"
            Cells(ActiveCell.row + 1, ActiveCell.column + 3) = "Amortization"
            Cells(ActiveCell.row + 1, ActiveCell.column + 4) = "RemainingNotional"
            Range(ActiveCell, Cells(ActiveCell.row + 1, ActiveCell.column + 4)).Font.Bold = True
        Application.ScreenUpdating = False
        data = assets(c).printCF
        For intCFdt = 0 To UBound(data)
            If p.Exists(data(intCFdt, 0)) Then
                dataExist = p(data(intCFdt, 0))
                dataExist(intAnz * 5) = data(intCFdt, 1)
                dataExist(intAnz * 5 + 1) = data(intCFdt, 2)
                dataExist(intAnz * 5 + 2) = data(intCFdt, 3)
                dataExist(intAnz * 5 + 3) = data(intCFdt, 4)
                dataExist(intAnz * 5 + 4) = data(intCFdt, 5)
                p(data(intCFdt, 0)) = dataExist
            Else
                dblNumber(intAnz * 5) = data(intCFdt, 1)
                dblNumber(intAnz * 5 + 1) = data(intCFdt, 2)
                dblNumber(intAnz * 5 + 2) = data(intCFdt, 3)
                dblNumber(intAnz * 5 + 3) = data(intCFdt, 4)
                dblNumber(intAnz * 5 + 4) = data(intCFdt, 5)
                p.Add data(intCFdt, 0), dblNumber
                Erase dblNumber
                ReDim dblNumber(assets.count * 5 - 1)
            End If
        Next
        intAnz = intAnz + 1
    Next
    Set p = funcSortKeysAsc(p)
    printIRMatrixOnSheet p, strIRBalance, strRef
End Sub

Public Sub printIRMatrixOnSheet(p As Dictionary, strSheet As String, strCell As String)
' Routine prints the cash flow matrix order by dates in a worksheet
' parameters:
' - p: cash flow matrix as dictionary ordered by dates, key: cash flow date, item: cash flows, remaining notionals etc at the cash flow date
' - strSheet: name of target worksheet
' - strCell: Cell at which the printing starts
Dim i As Integer, j As Integer, data As Variant, c As Variant, rng As Range

    Application.ScreenUpdating = False
    
    ThisWorkbook.Sheets(strSheet).Activate
    ActiveSheet.Range(strCell).Activate
    i = 0
    For Each c In p.Keys
        j = 0
        ActiveSheet.Cells(ActiveCell.row + i, ActiveCell.column + j) = c
        data = p(c)
        j = UBound(data)
        Set rng = ActiveSheet.Range(Cells(ActiveCell.row + i, ActiveCell.column + 1), Cells(ActiveCell.row + i, ActiveCell.column + 1))
        rng.Resize(1, UBound(data) + 1) = data
        i = i + 1
    Next
    ActiveSheet.Range("b3").Activate
    ActiveCell.Resize(p.count, UBound(data) + 1).NumberFormat = "0,000.00"
    
    Application.ScreenUpdating = True
End Sub

Function getRetailNominal(xName As String) As Double
' Function calculated the current nominal for retail or wholesale receivables based on a configured list of balance sheet account numbers
' implementation considers 'Retail' oder other, if other the Wholesale accounts are used for the aggregation
' Parameter: Name of the balance sheet position for which the notional has to be calculated
' Output: current notional of supplied balance sheet position

Dim nom As Double, rng As Range, data As Range, strAccName As String, row As Range
    Set data = ThisWorkbook.Sheets(strWholesaleData).Range(strBSData)
    If InStr(xName, "Retail") Then
        Set rng = ThisWorkbook.Sheets(strConfiguration).Range(strRetailAccounts)
    ElseIf InStr(xName, "Wholesale") Then
        Set rng = ThisWorkbook.Sheets(strConfiguration).Range(strWholesaleAccounts)
    ElseIf InStr(xName, "Leasing") Then
        Set rng = ThisWorkbook.Sheets(strConfiguration).Range(strLeasingAccounts)
    End If
    For Each row In rng.Rows
        nom = nom + Application.WorksheetFunction.SumIf(data.Columns(9), row.value, data.Columns(7))
    Next
    getRetailNominal = nom
End Function

