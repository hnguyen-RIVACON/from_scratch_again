Attribute VB_Name = "mdl_Printing"
Option Explicit


Sub printDefaultBucketedCashFlowMatrix(scen As String, indexStart As Integer, cashFlows() As Double, assetTypes As Dictionary, buckets() As String)

    ThisWorkbook.Sheets(strIRGAP).Activate
    Dim nrBuckets As Integer
    Dim nrInstrumentTypes As Integer
    Dim i As Integer
    Dim j As Integer
    Dim rng As Range
    
    
    nrBuckets = UBound(cashFlows, 1) + 1
    nrInstrumentTypes = UBound(cashFlows, 2) + 1
    
    ' 1. print buckets
    If indexStart = 0 Then
        Cells.Clear
        Set rng = ActiveSheet.Range(Cells(2, 2), Cells(2, UBound(buckets) + 2))
        rng = buckets
        rng.Font.Bold = True
    End If
        
    Dim resultGrid() As Double
    ReDim resultGrid(nrInstrumentTypes, nrBuckets - 1)
    For j = 0 To nrBuckets - 1
        resultGrid(nrInstrumentTypes, j) = 0
    Next j
        
        
    ' 2. print IR data by instrument typ
    Cells(4 + indexStart * (nrInstrumentTypes + 3), 1).Activate
        
    For i = 0 To nrInstrumentTypes - 1
        Cells(ActiveCell.row + i, ActiveCell.column) = assetTypes.Keys(i)
    Next i
    
    For i = 0 To nrInstrumentTypes - 1
        For j = 0 To nrBuckets - 1
            resultGrid(i, j) = cashFlows(j, i, 0)
            resultGrid(nrInstrumentTypes, j) = resultGrid(nrInstrumentTypes, j) + cashFlows(j, i, 0)
        Next
    Next
    Cells(ActiveCell.row + i, ActiveCell.column) = "IR GAP"
    Set rng = ActiveSheet.Range(Cells(ActiveCell.row, ActiveCell.column + 1), Cells(ActiveCell.row + nrInstrumentTypes, ActiveCell.column + nrBuckets))
    rng = resultGrid
    rng.NumberFormat = "0,000.00"
        
    ' 3. print scenario
    Cells(ActiveCell.row - 1, ActiveCell.column) = scen
    Cells(ActiveCell.row - 1, ActiveCell.column).Font.Bold = True
    Set rng = ActiveCell.Resize(nrInstrumentTypes + 1, 1)
    rng.Font.Bold = True


End Sub



Sub printBucketedCashFlowMatrix(scen As String, sheet As String, indexStart As Integer, cashFlows() As Double, assetTypes As Dictionary, buckets() As String)
' The routine calculates for each period of the IR GAP Analysis the IR GAP and prints aggregated interest cash flows in a worksheet
' parameters:
' - scen: name of the rate curve scenario that has been applied to the rate curves
' - indexStart: needed to print GAP results blockwise in the worksheet without overlap, paraemters gives the number of the scenario that is printed
    Dim rng As Range, i As Integer, j As Integer, k As Integer, nrBuckets As Integer, nrInstrumentTypes As Integer, nrCFTypes As Integer

    ThisWorkbook.Sheets(sheet).Activate
    
    nrBuckets = UBound(cashFlows, 1) + 1
    nrInstrumentTypes = UBound(cashFlows, 2) + 1
    nrCFTypes = UBound(cashFlows, 3) + 1
    
    ' 1. print buckets
    If indexStart = 0 Then
        Cells.Clear
    End If
    
    Set rng = ActiveSheet.Range(Cells(3 + indexStart * (nrBuckets + 4), 2), Cells(-2 + (indexStart + 1) * (nrBuckets + 4), 2))
    For j = 0 To nrBuckets - 1
        Cells(3 + indexStart * (nrBuckets + 4) + j, 2) = buckets(j)
    Next j
    rng.Font.Bold = True
        
    ' 2. print IR data by instrument typ
    Cells(1 + indexStart * (nrBuckets + 4), 4).Activate
        
    For i = 0 To nrInstrumentTypes
        For k = 0 To nrCFTypes - 1
            Cells(ActiveCell.row + 1, ActiveCell.column + i * nrCFTypes + k) = getCfNames(k)
        Next k
    Next i
    Dim resultGrid() As Double
    ReDim resultGrid(nrBuckets - 1, (nrInstrumentTypes + 1) * nrCFTypes - 1)
    For i = 0 To nrBuckets - 1
        For j = 0 To nrCFTypes - 1
            resultGrid(i, nrInstrumentTypes * nrCFTypes + j) = 0
        Next j
    Next i
    
    For i = 0 To nrInstrumentTypes - 1
        Cells(ActiveCell.row, ActiveCell.column + i * nrCFTypes) = assetTypes.Keys(i)
        
        For j = 0 To nrBuckets - 1
            For k = 0 To nrCFTypes - 1
                resultGrid(j, i * nrCFTypes + k) = cashFlows(j, i, k)
                resultGrid(j, nrInstrumentTypes * nrCFTypes + k) = resultGrid(j, nrInstrumentTypes * nrCFTypes + k) + cashFlows(j, i, k)
            Next k
        Next j
    Next i
    
    For i = 0 To nrInstrumentTypes - 1
        Cells(ActiveCell.row, ActiveCell.column + i * nrCFTypes) = assetTypes.Keys(i)
    Next i
    Cells(ActiveCell.row, ActiveCell.column + i * nrCFTypes) = "IR Gap"
    Set rng = ActiveSheet.Range(Cells(ActiveCell.row + 2, ActiveCell.column), Cells(ActiveCell.row + nrBuckets + 1, ActiveCell.column + (nrInstrumentTypes + 1) * nrCFTypes - 1))
    rng = resultGrid
    rng.NumberFormat = "0,000.00"
        
    ' 3. print scenario
    Cells(ActiveCell.row, ActiveCell.column - 1) = scen
    Cells(ActiveCell.row, ActiveCell.column - 1).Font.Bold = True
    Set rng = ActiveCell.Resize(1, (nrInstrumentTypes + 1) * nrCFTypes)
    rng.Font.Bold = True
End Sub

Sub printBFA3Report(sheet As String, insClasses As Dictionary, _
                    assetCashFlows() As Double, assetTypes As Dictionary, assetNotionals() As Double, assetBookValues() As Double, _
                    liabilityCashFlows() As Double, liabilityTypes As Dictionary, liabilityNotionals() As Double, liabilityBookValues() As Double, _
                    buckets() As String, discountCurve As clsRateCurve)

    Dim rng As Range, i As Integer, j As Integer
    Dim nrBuckets As Integer
    nrBuckets = UBound(buckets) + 1
    Dim nrAssetTypes As Integer
    nrAssetTypes = assetTypes.count
    Dim nrLiabilityTypes As Integer
    nrLiabilityTypes = liabilityTypes.count
    
    ThisWorkbook.Sheets(sheet).Activate
    
    Dim multiplier As Double
    multiplier = 0.000001 ' show all values in millions
    Cells.Clear
    Dim bucketLengths() As Double
    bucketLengths = getBucketLengths(buckets)
    
    ' print buckets header
    Set rng = ActiveSheet.Range(Cells(2, 4), Cells(2, 5 + nrBuckets))
    rng.Merge
    rng = "Cash Flows (Interest and Amortization)"
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter
    
    Set rng = ActiveSheet.Range(Cells(2, 5 + nrBuckets), Cells(2, 6 + nrBuckets))
    rng.Merge
    rng = "Balance Sheet & PV"
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter
    
    Set rng = ActiveSheet.Range(Cells(2, 8 + nrBuckets), Cells(2, 9 + nrBuckets))
    rng.Merge
    rng = "Result"
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter
    
    Set rng = ActiveSheet.Range(Cells(3, 4), Cells(3, 3 + nrBuckets))
    rng = buckets
    ActiveSheet.Cells(3, 5 + nrBuckets) = "Total"
    ActiveSheet.Cells(3, 7 + nrBuckets) = "Book Value"
    ActiveSheet.Cells(3, 8 + nrBuckets) = "Present Value"
    ActiveSheet.Cells(3, 10 + nrBuckets) = "H/(L) Book Value"
    ActiveSheet.Cells(3, 11 + nrBuckets) = "Comment"
    Set rng = ActiveSheet.Range(Cells(3, 4), Cells(3, 11 + nrBuckets))
    rng.Font.Bold = True
    
    Dim synthAssetCf() As Double
    ReDim synthAssetCf(nrBuckets - 1)
    Dim synthAssetNPV As Double
    Dim synthLiabilityCf() As Double
    ReDim synthLiabilityCf(nrBuckets - 1)
    Dim synthLiabilityNPV As Double
    Dim totalAssetCf() As Double
    ReDim totalAssetCf(nrBuckets - 1)
    Dim totalLiabilityCf() As Double
    ReDim totalLiabilityCf(nrBuckets - 1)
    Dim totalAssetBookValue As Double
    Dim totalLiabilityBookValue As Double
    Dim totalAssetNPV As Double
    Dim totalLiabilityNPV As Double
    
    Dim totalCf As Double
    Dim npv As Double
    Dim insType As InstrumentType
    Dim insClass As InstrumentClass
        
    Cells(4, 2) = "Interest Sensitive Assets"
    Cells(4, 2).Font.Bold = True
    
    For i = 0 To nrAssetTypes - 1
        insType = assetTypes.Keys(i)
        Set insClass = insClasses(insType)
        Cells(5 + i, 2) = insClass.longName
        
        totalCf = 0
        npv = 0
        For j = 0 To nrBuckets - 1
            Cells(5 + i, 4 + j) = assetCashFlows(j, i, 4) * multiplier
            totalAssetCf(j) = totalAssetCf(j) + assetCashFlows(j, i, 4)
            totalCf = totalCf + assetCashFlows(j, i, 4)
            npv = npv + assetCashFlows(j, i, 5)
        Next j
        Cells(5 + i, 5 + nrBuckets) = totalCf * multiplier
        Cells(5 + i, 7 + nrBuckets) = assetBookValues(i) * multiplier
        Cells(5 + i, 8 + nrBuckets) = npv * multiplier
        Cells(5 + i, 10 + nrBuckets) = (npv - assetBookValues(i)) * multiplier
        totalAssetBookValue = totalAssetBookValue + assetBookValues(i)
        totalAssetNPV = totalAssetNPV + npv
    Next i
    Cells(5 + nrAssetTypes, 2) = "Synthetic Assets"
    Cells(6 + nrAssetTypes, 2) = "Total Assets"
    Cells(6 + nrAssetTypes, 2).Font.Bold = True
    Cells(6 + nrAssetTypes, 11 + nrBuckets) = "PV Assets minus Book Value Assets"
    Set rng = ActiveSheet.Range(Cells(6 + nrAssetTypes, 2), Cells(6 + nrAssetTypes, 11 + nrBuckets))
    rng.Font.Bold = True
        
    Cells(9 + nrAssetTypes, 2) = "Interest Sensitive Liabilities"
    Cells(9 + nrAssetTypes, 2).Font.Bold = True
    For i = 0 To nrLiabilityTypes - 1
        insType = liabilityTypes.Keys(i)
        Set insClass = insClasses(insType)
        Cells(10 + nrAssetTypes + i, 2) = insClass.longName
        totalCf = 0
        npv = 0
        For j = 0 To nrBuckets - 1
            Cells(10 + nrAssetTypes + i, 4 + j) = liabilityCashFlows(j, i, 4) * multiplier
            totalLiabilityCf(j) = totalLiabilityCf(j) + liabilityCashFlows(j, i, 4)
            totalCf = totalCf + liabilityCashFlows(j, i, 4)
            npv = npv + liabilityCashFlows(j, i, 5)
        Next j
        Cells(10 + nrAssetTypes + i, 5 + nrBuckets) = totalCf * multiplier
        Cells(10 + nrAssetTypes + i, 7 + nrBuckets) = liabilityBookValues(i) * multiplier
        Cells(10 + nrAssetTypes + i, 8 + nrBuckets) = npv * multiplier
        Cells(10 + nrAssetTypes + i, 10 + nrBuckets) = (npv - liabilityBookValues(i)) * multiplier
        totalLiabilityBookValue = totalLiabilityBookValue + liabilityBookValues(i)
        totalLiabilityNPV = totalLiabilityNPV + npv
    Next i
    Cells(10 + nrAssetTypes + nrLiabilityTypes, 2) = "Synthetic Debt"
    Cells(11 + nrAssetTypes + nrLiabilityTypes, 2) = "Total Liabilities"
    Cells(11 + nrAssetTypes + nrLiabilityTypes, 2).Font.Bold = True
    Cells(11 + nrAssetTypes + nrLiabilityTypes, 11 + nrBuckets) = "PV Liabilities minus Book Value Liabilities"
    Set rng = ActiveSheet.Range(Cells(11 + nrAssetTypes + nrLiabilityTypes, 2), Cells(11 + nrAssetTypes + nrLiabilityTypes, 11 + nrBuckets))
    rng.Font.Bold = True
    Dim newTotal As Double
    For j = 0 To nrBuckets - 1
        newTotal = WorksheetFunction.Max(totalAssetCf(j), -totalLiabilityCf(j))
        synthAssetCf(j) = newTotal - totalAssetCf(j)
        synthAssetNPV = synthAssetNPV + discountCurve.getDFByYF((bucketLengths(j))) * synthAssetCf(j)
        synthLiabilityCf(j) = -newTotal - totalLiabilityCf(j)
        synthLiabilityNPV = synthLiabilityNPV + discountCurve.getDFByYF((bucketLengths(j))) * synthLiabilityCf(j)
        totalAssetCf(j) = newTotal
        totalLiabilityCf(j) = -newTotal
    Next j
    totalAssetNPV = totalAssetNPV + synthAssetNPV
    totalLiabilityNPV = totalLiabilityNPV + synthLiabilityNPV
     
    Dim totalSyntheticAssetCf As Double
    Dim totalSyntheticLiabilityCf As Double
    Dim totalTotalAssetCf As Double
    Dim totalTotalLiabilityCf As Double
    For j = 0 To nrBuckets - 1
        Cells(5 + nrAssetTypes, 4 + j) = synthAssetCf(j) * multiplier
        Cells(6 + nrAssetTypes, 4 + j) = totalAssetCf(j) * multiplier
        totalSyntheticAssetCf = totalSyntheticAssetCf + synthAssetCf(j)
        totalTotalAssetCf = totalTotalAssetCf + totalAssetCf(j)
        Cells(10 + nrAssetTypes + nrLiabilityTypes, 4 + j) = synthLiabilityCf(j) * multiplier
        Cells(11 + nrAssetTypes + nrLiabilityTypes, 4 + j) = totalLiabilityCf(j) * multiplier
        totalSyntheticLiabilityCf = totalSyntheticLiabilityCf + synthLiabilityCf(j)
        totalTotalLiabilityCf = totalTotalLiabilityCf + totalLiabilityCf(j)
    Next j
    Cells(5 + nrAssetTypes, 5 + nrBuckets) = totalSyntheticAssetCf * multiplier
    Cells(5 + nrAssetTypes, 8 + nrBuckets) = synthAssetNPV * multiplier
    Cells(6 + nrAssetTypes, 5 + nrBuckets) = totalTotalAssetCf * multiplier
    Cells(6 + nrAssetTypes, 7 + nrBuckets) = totalAssetBookValue * multiplier
    Cells(6 + nrAssetTypes, 8 + nrBuckets) = totalAssetNPV * multiplier
    Cells(6 + nrAssetTypes, 10 + nrBuckets) = (totalAssetNPV - totalAssetBookValue) * multiplier
    Cells(10 + nrAssetTypes + nrLiabilityTypes, 5 + nrBuckets) = totalSyntheticLiabilityCf * multiplier
    Cells(10 + nrAssetTypes + nrLiabilityTypes, 8 + nrBuckets) = synthLiabilityNPV * multiplier
    Cells(11 + nrAssetTypes + nrLiabilityTypes, 5 + nrBuckets) = totalTotalLiabilityCf * multiplier
    Cells(11 + nrAssetTypes + nrLiabilityTypes, 7 + nrBuckets) = totalLiabilityBookValue * multiplier
    Cells(11 + nrAssetTypes + nrLiabilityTypes, 8 + nrBuckets) = totalLiabilityNPV * multiplier
    Cells(11 + nrAssetTypes + nrLiabilityTypes, 10 + nrBuckets) = (totalLiabilityNPV - totalLiabilityBookValue) * multiplier
    
    Dim row As Integer
    row = nrAssetTypes + nrLiabilityTypes + 14
    Cells(row, 2) = "Total"
    Cells(row, 11 + nrBuckets) = "PV minus Book Value"
    For j = 0 To nrBuckets - 1
        Cells(row, 4 + j) = totalAssetCf(j) + totalLiabilityCf(j)
    Next j
    Cells(row, 5 + nrBuckets) = (totalTotalAssetCf + totalTotalLiabilityCf) * multiplier
    Cells(row, 7 + nrBuckets) = (totalAssetBookValue + totalLiabilityBookValue) * multiplier
    Cells(row, 8 + nrBuckets) = (totalAssetNPV + totalLiabilityNPV) * multiplier
    Cells(row, 10 + nrBuckets) = (totalAssetNPV + totalLiabilityNPV - totalAssetBookValue - totalLiabilityBookValue) * multiplier
    
    Set rng = ActiveSheet.Range(Cells(row, 2), Cells(row, 11 + nrBuckets))
    rng.Font.Bold = True

    
    
    row = row + 2
    
    ' Costs
    Dim bucketCosts() As Double
    ReDim bucketCosts(UBound(buckets))
    Dim remainingNotional As Double
    Dim length As Double
    
    
    'Begin, Version 0.19, Change of row + 3 to row + 4/Integration of Modelling OpCosts for Leasing, 09/01/2024, Marie Konrad'
    row = row + 4
    'End, Version 0.19, Change of row + 3 to row + 4/Integration of Modelling OpCosts for Leasing, 09/01/2024, Marie Konrad'
    Dim c As Double
    Cells(row, 2) = "Admin Costs"
    Cells(row, 2).Font.Bold = True
    Cells(row, 7 + nrBuckets) = "Aggregated"
    Cells(row, 7 + nrBuckets).Font.Bold = True
    row = row + 1
    Dim totalCosts As Double
    Dim totalCostsDiscounted As Double
    Dim rowCosts As Double
    Dim rowCostsDiscounted As Double
    For i = 0 To nrAssetTypes - 1
        insType = assetTypes.Keys(i)
        Set insClass = insClasses(insType)
        If insClass.costRatio <> 0 Then
            rowCosts = 0
            rowCostsDiscounted = 0
            Cells(row, 2) = insClass.longName
            remainingNotional = assetNotionals(i)
            
            For j = 0 To nrBuckets - 1
                If j = 0 Then
                    length = bucketLengths(j)
                Else
                    length = bucketLengths(j) - bucketLengths(j - 1)
                End If
                c = -length * remainingNotional * insClass.costRatio
                rowCosts = rowCosts + c
                rowCostsDiscounted = rowCostsDiscounted + discountCurve.getDFByYF((bucketLengths(j))) * c
                Cells(row, 4 + j) = c * multiplier
                bucketCosts(j) = bucketCosts(j) + c
                remainingNotional = remainingNotional - assetCashFlows(j, i, 2)
            Next j
            Cells(row, 7 + nrBuckets) = rowCosts * multiplier
            Cells(row, 8 + nrBuckets) = rowCostsDiscounted * multiplier
            Cells(row, 10 + nrBuckets) = rowCostsDiscounted * multiplier
            Cells(row, 11 + nrBuckets) = insClass.shortName & " Admin Costs"

            totalCosts = totalCosts + rowCosts
            totalCostsDiscounted = totalCosts + rowCostsDiscounted
            row = row + 1
        End If
    Next i
        
        
    ' synthetic costs
    Dim syntheticCosts As Double
    syntheticCosts = ThisWorkbook.Sheets(strConfiguration).Range(strSyntheticCosts).value
        
    Cells(row, 2) = "Synthetic Costs"
    rowCosts = 0
    rowCostsDiscounted = 0
    For j = 0 To nrBuckets - 1
        If j = 0 Then
            length = bucketLengths(j)
        Else
            length = bucketLengths(j) - bucketLengths(j - 1)
        End If
        If totalAssetCf(j) > 0.01 Then
            c = -syntheticCosts * length
            rowCosts = rowCosts + c
            rowCostsDiscounted = rowCostsDiscounted + discountCurve.getDFByYF((bucketLengths(j))) * c
            Cells(row, 4 + j) = c * multiplier
            bucketCosts(j) = bucketCosts(j) + c
        End If
    Next j
    totalCosts = totalCosts + rowCosts
    totalCostsDiscounted = totalCostsDiscounted + rowCostsDiscounted
    Cells(row, 7 + nrBuckets) = rowCosts * multiplier
    Cells(row, 8 + nrBuckets) = rowCostsDiscounted * multiplier
    Cells(row, 10 + nrBuckets) = rowCostsDiscounted * multiplier
    Cells(row, 11 + nrBuckets) = "Other Admin Costs"

        
    row = row + 1
    ' fixed costs
        
        
    Dim fixedCostsThreshold As Double
    Dim fixedCostsAmount As Double
    fixedCostsThreshold = ThisWorkbook.Sheets(strConfiguration).Range(strFixedCostsThreshold).value
    fixedCostsAmount = ThisWorkbook.Sheets(strConfiguration).Range(strFixedCostsAmount).value

        
    Cells(row, 2) = "Fixed Costs"
    rowCosts = 0
    rowCostsDiscounted = 0
    For j = 0 To nrBuckets - 1
        If j = 0 Then
            length = bucketLengths(j)
        Else
            length = bucketLengths(j) - bucketLengths(j - 1)
        End If
        If bucketCosts(j) > -fixedCostsThreshold * length Then
            c = -fixedCostsAmount * length
            Cells(row, 4 + j) = c * multiplier
            bucketCosts(j) = bucketCosts(j) + c
            rowCosts = rowCosts + c
            rowCostsDiscounted = rowCostsDiscounted + discountCurve.getDFByYF((bucketLengths(j))) * c
        End If
    Next j
    totalCosts = totalCosts + rowCosts
    totalCostsDiscounted = totalCostsDiscounted + rowCostsDiscounted
    Cells(row, 7 + nrBuckets) = rowCosts * multiplier
    Cells(row, 8 + nrBuckets) = rowCostsDiscounted * multiplier
    Cells(row, 10 + nrBuckets) = rowCostsDiscounted * multiplier
    Cells(row, 11 + nrBuckets) = "Fixed Costs"
    row = row + 1
        
    Cells(row, 2) = "Total"
    Cells(row, 11 + nrBuckets) = "Total Costs"
           
    For j = 0 To nrBuckets - 1
        Cells(row, 4 + j) = bucketCosts(j) * multiplier
    Next j
    Cells(row, 7 + nrBuckets) = totalCosts * multiplier
    Cells(row, 8 + nrBuckets) = totalCostsDiscounted * multiplier
    Cells(row, 10 + nrBuckets) = totalCostsDiscounted * multiplier
    Set rng = ActiveSheet.Range(Cells(row, 2), Cells(row, 11 + nrBuckets))
    rng.Font.Bold = True
        
    Dim riskAssessment As Double
    row = row + 2
    Cells(row, 11 + nrBuckets) = "Risk Assessment"
    Cells(row, 11 + nrBuckets).Font.Bold = True
    Cells(row, 10 + nrBuckets) = riskAssessment * multiplier
    row = row + 2
    Cells(row, 11 + nrBuckets) = "PV minus Book Value minus Costs minus Risk Assessment"
    Cells(row, 11 + nrBuckets).Font.Bold = True
    Cells(row, 10 + nrBuckets) = (totalAssetNPV + totalLiabilityNPV - totalAssetBookValue - totalLiabilityBookValue + totalCostsDiscounted + riskAssessment) * multiplier

    
        
    
End Sub


Sub printAssetLiabilityCf(assetTypes As Dictionary, cfMatrix() As Double, nrBuckets As Integer, insClasses As Dictionary, sheet As Worksheet, row As Integer, column As Integer, cfElement As Integer, multiplier As Double)
    Dim nrAssetTypes As Integer
    nrAssetTypes = assetTypes.count
    Dim i As Integer
    Dim j As Integer
    Dim insType As InstrumentType
    Dim insClass As InstrumentClass
    Dim cf As Double
    Dim totalCf As Double
    For i = 0 To nrAssetTypes - 1
        insType = assetTypes.Keys(i)
        Set insClass = insClasses(insType)
        sheet.Cells(row + i, column + 1) = insClass.longName
        totalCf = 0
        For j = 0 To nrBuckets - 1
            cf = cfMatrix(j, i, cfElement)
            sheet.Cells(row + i, column + 3 + j) = cf * multiplier
            totalCf = totalCf + cf
        Next j
        sheet.Cells(row + i, column + 4 + nrBuckets) = totalCf * multiplier
    Next i
End Sub

Sub printAssetLiabilityNii(assetTypes As Dictionary, cfMatrix() As Double, nrBuckets As Integer, insClasses As Dictionary, sheet As Worksheet, row As Integer, column As Integer, multiplier As Double)
    Dim nrAssetTypes As Integer
    nrAssetTypes = assetTypes.count
    Dim i As Integer
    Dim j As Integer
    Dim insType As InstrumentType
    Dim insClass As InstrumentClass
    Dim cf As Double
    Dim totalCf As Double
    For i = 0 To nrAssetTypes - 1
        insType = assetTypes.Keys(i)
        Set insClass = insClasses(insType)
        sheet.Cells(row + i, column + 1) = insClass.longName
        totalCf = 0
        For j = 0 To nrBuckets - 1
            cf = cfMatrix(i, j)
            sheet.Cells(row + i, column + 3 + j) = cf * multiplier
            totalCf = totalCf + cf
        Next j
        sheet.Cells(row + i, column + 4 + nrBuckets) = totalCf * multiplier
    Next i
End Sub


Function cfMatrixDiff2(cf1() As Double, cf2() As Double) As Double()

    Dim result() As Double
    ReDim result(UBound(cf1, 1), UBound(cf1, 2))
    Dim i As Double
    Dim j As Double
    Dim k As Double
    For i = 0 To UBound(cf1, 1)
        For j = 0 To UBound(cf1, 2)
            result(i, j) = cf1(i, j) - cf2(i, j)
        Next j
    Next i
    cfMatrixDiff2 = result
End Function
Function cfMatrixDiff3(cf1() As Double, cf2() As Double) As Double()

    Dim result() As Double
    ReDim result(UBound(cf1, 1), UBound(cf1, 2), UBound(cf2, 3))
    Dim i As Double
    Dim j As Double
    Dim k As Double
    For i = 0 To UBound(cf1, 1)
        For j = 0 To UBound(cf1, 2)
            For k = 0 To UBound(cf1, 3)
                result(i, j, k) = cf1(i, j, k) - cf2(i, j, k)
            Next k
        Next j
    Next i
    cfMatrixDiff3 = result
End Function

Sub printSAKIIRRBB(overviewSheetName As String, eveDetailSheetName As String, niiDetailSheetName As String, saki As Boolean, insClasses As Dictionary, _
                   scenAssetCashFlows As Dictionary, assetTypes As Dictionary, scenAssetNii As Dictionary, _
                   scenLiabilityCashFlows As Dictionary, liabilityTypes As Dictionary, scenLiabilityNii As Dictionary, _
                   buckets() As String, scenDiscountCurve As Dictionary, _
                   regulatoryCapital As Double, tier1Capital As Double)

    Dim rng As Range, i As Integer, j As Integer
    Dim nrBuckets As Integer
    nrBuckets = UBound(buckets) + 1
    Dim nrAssetTypes As Integer
    nrAssetTypes = assetTypes.count
    Dim nrLiabilityTypes As Integer
    nrLiabilityTypes = liabilityTypes.count
    Dim row As Integer
    If saki Then
        row = 1
    Else
        row = 20
    End If
    Dim eveRow As Integer
    eveRow = nrAssetTypes + nrLiabilityTypes + 15
    Dim eveRowCf As Integer
    eveRowCf = 1
    Dim niiRow As Integer
    niiRow = 1
    Dim multiplier As Double
    multiplier = 0.000001
    
    
    Dim overviewSheet As Worksheet
    Set overviewSheet = ThisWorkbook.Sheets(overviewSheetName)
    overviewSheet.Cells.Clear
    Dim eveDetailSheet As Worksheet
    Dim niiDetailSheet As Worksheet
    If Not saki Then
        Set eveDetailSheet = ThisWorkbook.Sheets(eveDetailSheetName)
        eveDetailSheet.Cells.Clear
        Set niiDetailSheet = ThisWorkbook.Sheets(niiDetailSheetName)
        niiDetailSheet.Cells.Clear
    End If
    Dim evePerScen As New Dictionary
    Dim niiPerScen As New Dictionary
    
    Dim bucketLengths() As Double
    bucketLengths = getBucketLengths(buckets)
    
    ' print buckets header
    Set rng = overviewSheet.Range(overviewSheet.Cells(row + 1, 4), overviewSheet.Cells(row + 1, 5 + nrBuckets))
    rng.Merge
    rng = "Cash Flows (Interest and Amortization)"
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter

    
    Set rng = overviewSheet.Range(overviewSheet.Cells(row + 2, 4), overviewSheet.Cells(row + 2, 3 + nrBuckets))
    rng = buckets
    overviewSheet.Cells(row + 2, 5 + nrBuckets) = "Total"
    Set rng = overviewSheet.Range(overviewSheet.Cells(row + 2, 4), overviewSheet.Cells(row + 2, 5 + nrBuckets))
    rng.Font.Bold = True
    
    Dim totalAssetCf() As Double
    Dim totalLiabilityCf() As Double
    Dim bucketedAssetNPV() As Double
    Dim bucketedLiabilityNPV() As Double
    Dim bucketedAssetNPVBase() As Double
    Dim bucketedLiabilityNPVBase() As Double
    Dim totalAssetNPV As Double
    Dim totalLiabilityNPV As Double
    Dim totalAssetNPVBase As Double
    Dim totalLiabilityNPVBase As Double
    Dim totalTotalAssetCf As Double
    Dim totalTotalLiabilityCf As Double
        
    Dim bucketedAssetNii() As Double
    Dim bucketedLiabilityNii() As Double
    Dim bucketedAssetNiiBase() As Double
    Dim bucketedLiabilityNiiBase() As Double
    Dim totalAssetNii As Double
    Dim totalLiabilityNii As Double
    Dim totalAssetNiiBase As Double
    Dim totalLiabilityNiiBase As Double
    
    Dim totalCf As Double
    Dim npv As Double
    Dim insType As InstrumentType
    Dim insClass As InstrumentClass
        
    Dim cf As Double
    
    Dim scen As Variant
    Dim discountCurve As clsRateCurve
    Dim baseNPVAssets As Double
    Dim baseNPVLiabilities As Double
    
    overviewSheet.Cells(row + 3, 2) = "Interest Sensitive Assets"
    overviewSheet.Cells(row + 3, 2).Font.Bold = True
    Dim cfx() As Double
    Dim cfx2() As Double
    Dim nii() As Double
    Dim cfDelta() As Double
    Dim niiDelta() As Double
    Dim scenCtr As Integer
    scenCtr = 0
    For Each scen In scenDiscountCurve.Keys
        ReDim totalAssetCf(nrBuckets - 1)
        ReDim totalLiabilityCf(nrBuckets - 1)
        ReDim bucketedAssetNPV(nrBuckets - 1)
        ReDim bucketedLiabilityNPV(nrBuckets - 1)
        totalAssetNPV = 0
        totalLiabilityNPV = 0
        If Not saki Then
            ReDim bucketedAssetNii(nrBuckets - 1)
            ReDim bucketedLiabilityNii(nrBuckets - 1)
            totalAssetNii = 0
            totalLiabilityNii = 0
        End If
        Set discountCurve = scenDiscountCurve(scen)
        For i = 0 To nrAssetTypes - 1
            insType = assetTypes.Keys(i)
            Set insClass = insClasses(insType)
            npv = 0
            For j = 0 To nrBuckets - 1
                cf = scenAssetCashFlows(scen)(j, i, cfTotalCf)
                totalAssetCf(j) = totalAssetCf(j) + cf
                npv = npv + scenAssetCashFlows(scen)(j, i, 5)
                bucketedAssetNPV(j) = bucketedAssetNPV(j) + scenAssetCashFlows(scen)(j, i, cfTotalNpv)
            Next j
            totalAssetNPV = totalAssetNPV + npv
        Next i
        If Not saki Then
            For i = 0 To nrAssetTypes - 1
                For j = 0 To nrBuckets - 1
                    cf = scenAssetNii(scen)(i, j)
                    bucketedAssetNii(j) = bucketedAssetNii(j) + cf
                    totalAssetNii = totalAssetNii + cf
                Next j
            Next i
            If scen = "" Then
                bucketedAssetNiiBase = bucketedAssetNii
                totalAssetNiiBase = totalAssetNii
            End If
        End If
        If scen = "" Then
            bucketedAssetNPVBase = bucketedAssetNPV
            totalAssetNPVBase = totalAssetNPV
            cfx = scenAssetCashFlows(scen)
            Call printAssetLiabilityCf(assetTypes, cfx, nrBuckets, insClasses, overviewSheet, row + 4, 1, cfTotalCf, multiplier)
            overviewSheet.Cells(row + 5 + nrAssetTypes, 2) = "Total Assets"
            overviewSheet.Cells(row + 8 + nrAssetTypes, 2) = "Interest Sensitive Liabilities"
            overviewSheet.Cells(row + 8 + nrAssetTypes, 2).Font.Bold = True
        End If
        If Not saki Then
            eveDetailSheet.Cells(eveRowCf, (nrBuckets + 6) * scenCtr + 5) = scen
            eveDetailSheet.Cells(eveRowCf + 1, (nrBuckets + 6) * scenCtr + 4) = "Interest and Amortization Cash Flows"
            Set rng = eveDetailSheet.Range(eveDetailSheet.Cells(eveRowCf + 2, (nrBuckets + 6) * scenCtr + 4), eveDetailSheet.Cells(eveRowCf + 2, (nrBuckets + 6) * (scenCtr + 1) - 3))
            rng = buckets
            eveDetailSheet.Cells(eveRowCf + 3, (nrBuckets + 6) * scenCtr + 2) = "Interest Sensitive Assets"
            cfx = scenAssetCashFlows(scen)
            Call printAssetLiabilityCf(assetTypes, cfx, nrBuckets, insClasses, eveDetailSheet, eveRowCf + 4, (nrBuckets + 6) * scenCtr + 1, cfTotalCf, multiplier)
            
            
            eveDetailSheet.Cells(eveRow + 1, 2) = scen
            eveDetailSheet.Cells(eveRow + 1, 2).Font.Bold = True
            Set rng = eveDetailSheet.Range(eveDetailSheet.Cells(eveRow + 2, 4), eveDetailSheet.Cells(eveRow + 2, nrBuckets + 3))
            rng = buckets
            eveDetailSheet.Cells(eveRow + 3, 2) = "Interest Sensitive Assets Total"
            Call printAssetLiabilityCf(assetTypes, cfx, nrBuckets, insClasses, eveDetailSheet, eveRow + 4, 1, cfTotalNpv, multiplier)
            
            nii = scenAssetNii(scen)
            niiDetailSheet.Cells(niiRow + 1, 2) = scen
            niiDetailSheet.Cells(niiRow + 1, 2).Font.Bold = True
            Set rng = niiDetailSheet.Range(niiDetailSheet.Cells(niiRow + 2, 4), niiDetailSheet.Cells(niiRow + 2, nrBuckets + 3))
            rng = buckets
            niiDetailSheet.Cells(niiRow + 3, 2) = "Interest Sensitive Assets Total"
            Call printAssetLiabilityNii(assetTypes, nii, nrBuckets, insClasses, niiDetailSheet, niiRow + 4, 1, multiplier)
            
        End If
        For i = 0 To nrLiabilityTypes - 1
            insType = liabilityTypes.Keys(i)
            Set insClass = insClasses(insType)
            npv = 0
            For j = 0 To nrBuckets - 1
                cf = scenLiabilityCashFlows(scen)(j, i, 4)
                totalLiabilityCf(j) = totalLiabilityCf(j) + cf
                npv = npv + scenLiabilityCashFlows(scen)(j, i, 5)
                bucketedLiabilityNPV(j) = bucketedLiabilityNPV(j) + scenLiabilityCashFlows(scen)(j, i, 5)
            Next j
            totalLiabilityNPV = totalLiabilityNPV + npv
        Next i
        If Not saki Then
            For i = 0 To nrLiabilityTypes - 1
                For j = 0 To nrBuckets - 1
                    cf = scenLiabilityNii(scen)(i, j)
                    bucketedLiabilityNii(j) = bucketedLiabilityNii(j) + cf
                    totalLiabilityNii = totalLiabilityNii + cf
                Next j
            Next i
            If scen = "" Then
                bucketedLiabilityNiiBase = bucketedLiabilityNii
                totalLiabilityNiiBase = totalLiabilityNii
            Else
                niiPerScen.Add scen, (totalAssetNiiBase + totalLiabilityNiiBase) - (totalAssetNii + totalLiabilityNii)
            End If
        End If
        If Not saki Then
            eveDetailSheet.Cells(eveRowCf + 8 + nrAssetTypes, (nrBuckets + 6) * scenCtr + 2) = "Interest Sensitive Liabilities"
            cfx = scenLiabilityCashFlows(scen)
            Call printAssetLiabilityCf(liabilityTypes, cfx, nrBuckets, insClasses, eveDetailSheet, eveRowCf + 9 + nrAssetTypes, (nrBuckets + 6) * scenCtr + 1, cfTotalCf, multiplier)
            For j = 0 To nrBuckets - 1
            
            Next j
            
            
            eveDetailSheet.Cells(eveRow + nrAssetTypes + 8, 2) = "Interest Sensitive Liabilities Total"
            Call printAssetLiabilityCf(liabilityTypes, cfx, nrBuckets, insClasses, eveDetailSheet, eveRow + nrAssetTypes + 9, 1, cfTotalNpv, multiplier)
            
            eveDetailSheet.Cells(eveRow + nrAssetTypes + nrLiabilityTypes + 10, 2) = "Total"
            For j = 0 To nrBuckets - 1
                eveDetailSheet.Cells(eveRow + nrAssetTypes + nrLiabilityTypes + 10, j + 4) = (bucketedAssetNPV(j) + bucketedLiabilityNPV(j)) * multiplier
            Next j
            eveDetailSheet.Cells(eveRow + nrAssetTypes + nrLiabilityTypes + 10, nrBuckets + 5) = (totalAssetNPV + totalLiabilityNPV) * multiplier
            
            
            
            nii = scenLiabilityNii(scen)
            niiDetailSheet.Cells(niiRow + nrAssetTypes + 8, 2) = "Interest Sensitive Liabilities Total"
            Call printAssetLiabilityNii(liabilityTypes, nii, nrBuckets, insClasses, niiDetailSheet, niiRow + nrAssetTypes + 9, 1, multiplier)
            niiDetailSheet.Cells(niiRow + nrAssetTypes + nrLiabilityTypes + 10, 2) = "Total"
            For j = 0 To nrBuckets - 1
                niiDetailSheet.Cells(niiRow + nrAssetTypes + nrLiabilityTypes + 10, j + 4) = (bucketedAssetNii(j) + bucketedLiabilityNii(j)) * multiplier
            Next j
            niiDetailSheet.Cells(niiRow + nrAssetTypes + nrLiabilityTypes + 10, nrBuckets + 5) = (totalAssetNii + totalLiabilityNii) * multiplier

            eveRow = eveRow + nrAssetTypes + nrLiabilityTypes + 15
            niiRow = niiRow + nrAssetTypes + nrLiabilityTypes + 15
        End If
        
        If scen = "" Then
            bucketedLiabilityNPVBase = bucketedLiabilityNPV
            totalLiabilityNPVBase = totalLiabilityNPV
            cfx = scenLiabilityCashFlows(scen)
            Call printAssetLiabilityCf(liabilityTypes, cfx, nrBuckets, insClasses, overviewSheet, row + 9 + nrAssetTypes, 1, cfTotalCf, multiplier)
    
            overviewSheet.Cells(row + 10 + nrAssetTypes + nrLiabilityTypes, 2) = "Total Liabilities"
            overviewSheet.Cells(row + 10 + nrAssetTypes + nrLiabilityTypes, 2).Font.Bold = True
        
            For j = 0 To nrBuckets - 1
                overviewSheet.Cells(row + 5 + nrAssetTypes, 4 + j) = totalAssetCf(j) * multiplier
                totalTotalAssetCf = totalTotalAssetCf + totalAssetCf(j)
                overviewSheet.Cells(row + 10 + nrAssetTypes + nrLiabilityTypes, 4 + j) = totalLiabilityCf(j) * multiplier
                totalTotalLiabilityCf = totalTotalLiabilityCf + totalLiabilityCf(j)
            Next j
            overviewSheet.Cells(row + 5 + nrAssetTypes, 5 + nrBuckets) = totalTotalAssetCf * multiplier
            overviewSheet.Cells(row + 10 + nrAssetTypes + nrLiabilityTypes, 5 + nrBuckets) = totalTotalLiabilityCf * multiplier
            Set rng = overviewSheet.Range(overviewSheet.Cells(row + 5 + nrAssetTypes, 2), overviewSheet.Cells(row + 5 + nrAssetTypes, 5 + nrBuckets))
            rng.Font.Bold = True
            Set rng = overviewSheet.Range(overviewSheet.Cells(row + 10 + nrAssetTypes + nrLiabilityTypes, 2), overviewSheet.Cells(row + 10 + nrAssetTypes + nrLiabilityTypes, 5 + nrBuckets))
            rng.Font.Bold = True
        
            row = row + nrAssetTypes + nrLiabilityTypes + 13
            overviewSheet.Cells(row, 2) = "Total"
            For j = 0 To nrBuckets - 1
                overviewSheet.Cells(row, 4 + j) = (totalAssetCf(j) + totalLiabilityCf(j)) * multiplier
            Next j
            overviewSheet.Cells(row, nrBuckets + 5) = (totalTotalAssetCf + totalTotalLiabilityCf) * multiplier
            Set rng = overviewSheet.Range(overviewSheet.Cells(row, 2), overviewSheet.Cells(row, 5 + nrBuckets))
            rng.Font.Bold = True
            
            overviewSheet.Cells(row + 2, 2) = "Regulatory Capital (RC)"
            overviewSheet.Cells(row + 2, 3) = regulatoryCapital * multiplier
            overviewSheet.Cells(row + 3, 2) = "Tier 1 Capital (RC)"
            overviewSheet.Cells(row + 3, 3) = tier1Capital * multiplier
            Set rng = overviewSheet.Range(overviewSheet.Cells(row + 2, 2), overviewSheet.Cells(row + 3, 3))
            rng.Font.Bold = True
            row = row + 5
            If Not saki Then
                eveDetailSheet.Cells(eveRow + 2, 2) = "Regulatory Capital (RC)"
                eveDetailSheet.Cells(eveRow + 2, 3) = regulatoryCapital * multiplier
                eveDetailSheet.Cells(eveRow + 3, 2) = "Tier 1 Capital (RC)"
                eveDetailSheet.Cells(eveRow + 3, 3) = tier1Capital * multiplier
                eveRow = eveRow + 5
                niiDetailSheet.Cells(niiRow + 2, 2) = "Regulatory Capital (RC)"
                niiDetailSheet.Cells(niiRow + 2, 3) = regulatoryCapital
                niiDetailSheet.Cells(niiRow + 3, 2) = "Tier 1 Capital (RC)"
                niiDetailSheet.Cells(niiRow + 3, 3) = tier1Capital * multiplier
                niiRow = niiRow + 5
            End If
        
        End If
        
        If scen = "" Then
            overviewSheet.Cells(row, 1) = "Baseline"
        Else
            overviewSheet.Cells(row, 1) = "Scenario " & scen
        End If
        overviewSheet.Cells(row, 1).Font.Bold = True
        Set rng = overviewSheet.Range(overviewSheet.Cells(row, 4), overviewSheet.Cells(row, 3 + nrBuckets))
        rng = buckets
        overviewSheet.Cells(row, 5 + nrBuckets) = "Total"
        
        
        Set rng = overviewSheet.Range(overviewSheet.Cells(row, 4), overviewSheet.Cells(row, 5 + nrBuckets))
        rng.Font.Bold = True
        overviewSheet.Cells(row + 1, 2) = "Interest Rates"
        overviewSheet.Cells(row + 2, 2) = "Discount Rates"
        overviewSheet.Cells(row + 3, 2) = "Present Value Assets"
        overviewSheet.Cells(row + 4, 2) = "Present Value Liabilities"
        overviewSheet.Cells(row + 6, 2) = "Present Value Total"
        For j = 0 To nrBuckets - 1
            overviewSheet.Cells(row + 1, 4 + j) = discountCurve.getZR(bucketLengths(j))
            overviewSheet.Cells(row + 2, 4 + j) = discountCurve.getDFByYF(bucketLengths(j))
            overviewSheet.Cells(row + 3, 4 + j) = bucketedAssetNPV(j) * multiplier
            overviewSheet.Cells(row + 4, 4 + j) = bucketedLiabilityNPV(j) * multiplier
            overviewSheet.Cells(row + 6, 4 + j) = (bucketedAssetNPV(j) + bucketedLiabilityNPV(j)) * multiplier
        Next j
        overviewSheet.Cells(row + 3, 5 + nrBuckets) = totalAssetNPV * multiplier
        overviewSheet.Cells(row + 4, 5 + nrBuckets) = totalLiabilityNPV * multiplier
        overviewSheet.Cells(row + 6, 5 + nrBuckets) = (totalAssetNPV + totalLiabilityNPV) * multiplier
        If scen = "" Then
            baseNPVAssets = totalAssetNPV
            baseNPVLiabilities = totalLiabilityNPV
        Else
            Set rng = overviewSheet.Range(overviewSheet.Cells(row - 1, 7 + nrBuckets), overviewSheet.Cells(row - 1, 9 + nrBuckets))
            rng.Merge
            rng = "vs Baseline"
            rng.Font.Bold = True
            rng.HorizontalAlignment = xlCenter
            overviewSheet.Cells(row, 7 + nrBuckets) = "H/(L)"
            overviewSheet.Cells(row, 8 + nrBuckets) = "H/(L) / T1"
            overviewSheet.Cells(row, 9 + nrBuckets) = "H/(L) / RC"

            overviewSheet.Cells(row + 3, 7 + nrBuckets) = (totalAssetNPV - baseNPVAssets) * multiplier
            overviewSheet.Cells(row + 4, 7 + nrBuckets) = (totalLiabilityNPV - baseNPVLiabilities) * multiplier
            overviewSheet.Cells(row + 6, 7 + nrBuckets) = (totalAssetNPV + totalLiabilityNPV - (baseNPVAssets + baseNPVLiabilities)) * multiplier
            overviewSheet.Cells(row + 6, 8 + nrBuckets) = (totalAssetNPV + totalLiabilityNPV - (baseNPVAssets + baseNPVLiabilities)) / tier1Capital
            overviewSheet.Cells(row + 6, 9 + nrBuckets) = (totalAssetNPV + totalLiabilityNPV - (baseNPVAssets + baseNPVLiabilities)) / regulatoryCapital
            evePerScen.Add scen, (baseNPVAssets + baseNPVLiabilities) - (totalAssetNPV + totalLiabilityNPV)
        End If
        If Not saki And scen <> "" Then
            If nrAssetTypes > 0 Then
                cfx = scenAssetCashFlows(scen)
                cfx2 = scenAssetCashFlows("")
                cfDelta = cfMatrixDiff3(cfx, cfx2)
            End If
            
            eveDetailSheet.Cells(eveRow - 1, 2) = scen & " vs Baseline"
            eveDetailSheet.Cells(eveRow - 1, nrBuckets + 7) = "H/(L)"
            eveDetailSheet.Cells(eveRow - 1, nrBuckets + 8) = "H/(L) / T1"
            eveDetailSheet.Cells(eveRow - 1, nrBuckets + 9) = "H/(L) / RC"
            eveDetailSheet.Cells(eveRow, 2) = "Interest Sensitive Assets Total"
            Call printAssetLiabilityCf(assetTypes, cfDelta, nrBuckets, insClasses, eveDetailSheet, eveRow + 1, 1, cfTotalNpv, multiplier)
            If nrLiabilityTypes > 0 Then
                cfx = scenLiabilityCashFlows(scen)
                cfx2 = scenLiabilityCashFlows("")
                cfDelta = cfMatrixDiff3(cfx, cfx2)
            End If
            eveDetailSheet.Cells(eveRow + nrAssetTypes + 2, 2) = "Interest Sensitive Liabilities Total"
            Call printAssetLiabilityCf(liabilityTypes, cfDelta, nrBuckets, insClasses, eveDetailSheet, eveRow + nrAssetTypes + 3, 1, cfTotalNpv, multiplier)
            eveDetailSheet.Cells(eveRow + nrAssetTypes + nrLiabilityTypes + 5, 2) = "Total"
            For j = 0 To nrBuckets - 1
                eveDetailSheet.Cells(eveRow + nrAssetTypes + nrLiabilityTypes + 5, j + 4) = (bucketedAssetNPV(j) + bucketedLiabilityNPV(j) - (bucketedAssetNPVBase(j) + bucketedLiabilityNPVBase(j))) * multiplier
            Next j
            eveDetailSheet.Cells(eveRow + nrAssetTypes + nrLiabilityTypes + 5, nrBuckets + 5) = (totalAssetNPV + totalLiabilityNPV - (totalAssetNPVBase + totalLiabilityNPVBase)) * multiplier
            eveDetailSheet.Cells(eveRow + nrAssetTypes + nrLiabilityTypes + 5, nrBuckets + 7) = (totalAssetNPV + totalLiabilityNPV - (totalAssetNPVBase + totalLiabilityNPVBase)) * multiplier
            eveDetailSheet.Cells(eveRow + nrAssetTypes + nrLiabilityTypes + 5, nrBuckets + 8) = (totalAssetNPV + totalLiabilityNPV - (totalAssetNPVBase + totalLiabilityNPVBase)) / tier1Capital
            eveDetailSheet.Cells(eveRow + nrAssetTypes + nrLiabilityTypes + 5, nrBuckets + 9) = (totalAssetNPV + totalLiabilityNPV - (totalAssetNPVBase + totalLiabilityNPVBase)) / regulatoryCapital
            
            eveRow = eveRow + nrAssetTypes + nrLiabilityTypes + 15
            
            
            
            If nrAssetTypes > 0 Then
                cfx = scenAssetNii(scen)
                cfx2 = scenAssetNii("")
                cfDelta = cfMatrixDiff2(cfx, cfx2)
            End If
            
            niiDetailSheet.Cells(niiRow - 1, 2) = scen & " vs Baseline"
            niiDetailSheet.Cells(niiRow - 1, nrBuckets + 7) = "H/(L)"
            niiDetailSheet.Cells(niiRow - 1, nrBuckets + 8) = "H/(L) / T1"
            niiDetailSheet.Cells(niiRow - 1, nrBuckets + 9) = "H/(L) / RC"
            
            niiDetailSheet.Cells(niiRow, 2) = "Interest Sensitive Assets Total"
            Call printAssetLiabilityNii(assetTypes, cfDelta, nrBuckets, insClasses, niiDetailSheet, niiRow + 1, 1, multiplier)
            If nrLiabilityTypes > 0 Then
                cfx = scenLiabilityNii(scen)
                cfx2 = scenLiabilityNii("")
                cfDelta = cfMatrixDiff2(cfx, cfx2)
            End If
            niiDetailSheet.Cells(niiRow + nrAssetTypes + 2, 2) = "Interest Sensitive Liabilities Total"
            Call printAssetLiabilityNii(liabilityTypes, cfDelta, nrBuckets, insClasses, niiDetailSheet, niiRow + nrAssetTypes + 3, 1, multiplier)
            niiDetailSheet.Cells(niiRow + nrAssetTypes + nrLiabilityTypes + 5, 2) = "Total"
            For j = 0 To nrBuckets - 1
                niiDetailSheet.Cells(niiRow + nrAssetTypes + nrLiabilityTypes + 5, j + 4) = (bucketedAssetNii(j) + bucketedLiabilityNii(j) - (bucketedAssetNiiBase(j) + bucketedLiabilityNiiBase(j))) * multiplier
            Next j
            niiDetailSheet.Cells(niiRow + nrAssetTypes + nrLiabilityTypes + 5, nrBuckets + 5) = (totalAssetNii + totalLiabilityNii - (totalAssetNiiBase + totalLiabilityNiiBase)) * multiplier
            niiDetailSheet.Cells(niiRow + nrAssetTypes + nrLiabilityTypes + 5, nrBuckets + 7) = (totalAssetNii + totalLiabilityNii - (totalAssetNiiBase + totalLiabilityNiiBase)) * multiplier
            niiDetailSheet.Cells(niiRow + nrAssetTypes + nrLiabilityTypes + 5, nrBuckets + 8) = (totalAssetNii + totalLiabilityNii - (totalAssetNiiBase + totalLiabilityNiiBase)) / tier1Capital
            niiDetailSheet.Cells(niiRow + nrAssetTypes + nrLiabilityTypes + 5, nrBuckets + 9) = (totalAssetNii + totalLiabilityNii - (totalAssetNiiBase + totalLiabilityNiiBase)) / regulatoryCapital
            
            niiRow = niiRow + nrAssetTypes + nrLiabilityTypes + 15
            
            
            
            
            
        End If
        row = row + 12
    
        scenCtr = scenCtr + 1
    Next scen
    
    
    If Not saki Then
        Dim maxDeltaEve As Double
        Dim maxDeltaNii As Double
        maxDeltaEve = -1E+99
        maxDeltaNii = -1E+99
        row = 2
        overviewSheet.Cells(row, 4) = "EVE"
        overviewSheet.Cells(row, 5) = "NII"
        row = row + 1
        For Each scen In scenDiscountCurve
            If scen <> "" Then
                overviewSheet.Cells(row, 2) = scen
                overviewSheet.Cells(row, 4) = evePerScen(scen) * multiplier
                overviewSheet.Cells(row, 5) = niiPerScen(scen) * multiplier
                If evePerScen(scen) > maxDeltaEve Then
                    maxDeltaEve = evePerScen(scen)
                End If
                If niiPerScen(scen) > maxDeltaNii Then
                    maxDeltaNii = niiPerScen(scen)
                End If
                row = row + 1
            End If
        Next scen
        overviewSheet.Cells(row, 2) = "Highest Position"
        overviewSheet.Cells(row, 4) = maxDeltaEve * multiplier
        overviewSheet.Cells(row, 5) = maxDeltaNii * multiplier
    End If
        
    overviewSheet.Activate
        
    
End Sub

