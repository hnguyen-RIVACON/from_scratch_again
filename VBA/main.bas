Attribute VB_Name = "main"
Option Explicit

Sub main()
Dim curveDic As Scripting.Dictionary
Dim assetsDic As Scripting.Dictionary

' read curve market data and calibrate curves: iterate over curves data as given by ranges "RateCurves" and "Scenarios"
' OR
' read calibrated curves from sheet
' optionally: print curve

    Set curveDic = ReadandCalibrateCurve(True)
    ' Set CurveDic = readCalibratedCurvesFromSheet()

' Create IR rate developments

    CalculatePortfolio "all"
        
    ' read portfolio
        'Set AssetsDic = New Scripting.Dictionary
        'Set AssetsDic = InstrumentFactory()

' iterate over all Curves and print the ir development incl. NPVs
' calculate IR gap based on NPV
' ensure that ir curve calibration can be decoupled from ir development analysis

    ' set Discounting Curve Name
        Dim curveFreq As Integer
        curveFreq = 3
        
    ' calc ir cash flows
        'calcIRCF AssetsDic, curveDic, curveFreq, ""
    ' print ir cash flows
        'printIRCF AssetsDic
    
End Sub

Sub bfa3()
    Dim curveDic As Scripting.Dictionary
    Dim assetsDic As Scripting.Dictionary



    Set curveDic = ReadandCalibrateCurve(True)

    Dim insClasses As Dictionary
    Set insClasses = getInstrumentClasses()
    
    Dim reports As Dictionary
    Set reports = getReports(insClasses)
    Dim reportConfig As Report
    Set reportConfig = reports(BFA3Report)
    
    Dim scen As String
    scen = ""
    
    
    Dim pfAssets As New clsPortfolio
    Set pfAssets = PortfolioFactory(insClasses, reportConfig.getEnabledInstruments(), curveDic, "all", True, False, True)     ' reads all assets
    Dim pfLiabilities As New clsPortfolio
    Set pfLiabilities = PortfolioFactory(insClasses, reportConfig.getEnabledInstruments, curveDic, "all", False, True, True)    ' reads all liabilities
    
    
    Dim discCurve As String
    discCurve = ThisWorkbook.Sheets(strConfiguration).Range(strDiscountCurve).value
    
    Dim row As Range
    ' calc ir cash flows give a scenario
    If curveDic.count = 0 Then
        MsgBox "There are no calibrated curve for the selected scenario. Please check the available curves or re-calibrate the required curves."
        Exit Sub
    End If
    
    Dim buckets() As String
    Dim effectiveBuckets() As String
    buckets = getBuckets(ThisWorkbook.Sheets(strConfiguration).Range(strGAPBuckets))
    
    pfAssets.calcInstrIrCF curveDic, discCurve, scen
    pfLiabilities.calcInstrIrCF curveDic, discCurve, scen
    
    Dim assetDailyCashFlows As Dictionary
    Set assetDailyCashFlows = pfAssets.calcCashFlowMatrix()
    Dim liabilityDailyCashFlows As Dictionary
    Set liabilityDailyCashFlows = pfLiabilities.calcCashFlowMatrix()
    
    Dim assetCashFlows() As Double
    assetCashFlows = pfAssets.aggregateCashFlowMatrix(assetDailyCashFlows, buckets, effectiveBuckets)
    Dim liabilityCashFlows() As Double
    liabilityCashFlows = pfLiabilities.aggregateCashFlowMatrix(liabilityDailyCashFlows, buckets, effectiveBuckets)
    
    ' notionals are required for calculation of costs
    Dim assetNotionals() As Double
    assetNotionals = pfAssets.getTotalNotionals()
    Dim liabilityNotionals() As Double
    liabilityNotionals = pfLiabilities.getTotalNotionals()
    
    Dim assetBookValues() As Double
    assetBookValues = pfAssets.getTotalBookValues()
    Dim liabilityBookValues() As Double
    liabilityBookValues = pfLiabilities.getTotalBookValues()
    
        
    Call printBFA3Report("BFA3", insClasses, assetCashFlows, pfAssets.assetTypeIndices, assetNotionals, assetBookValues, _
                    liabilityCashFlows, pfLiabilities.assetTypeIndices, liabilityNotionals, liabilityBookValues, buckets, curveDic(getScenarioCurveName(discCurve, "")))
    
    
End Sub

Sub RunSakiOverviewReport()
    Dim curveDic As Scripting.Dictionary
    Dim assetsDic As Scripting.Dictionary


    Set curveDic = ReadandCalibrateCurve(True)

    Dim insClasses As Dictionary
    Set insClasses = getInstrumentClasses()
    
    Dim reports As Dictionary
    Set reports = getReports(insClasses)
    Dim reportConfig As Report
    Set reportConfig = reports(saki)
        
    
    Dim pfAssets As New clsPortfolio
    Set pfAssets = PortfolioFactory(insClasses, reportConfig.getEnabledInstruments(), curveDic, "all", True, False, True)     ' reads all assets
    Dim pfLiabilities As New clsPortfolio
    Set pfLiabilities = PortfolioFactory(insClasses, reportConfig.getEnabledInstruments, curveDic, "all", False, True, True)    ' reads all liabilities
    
    Dim discCurve As String
    discCurve = ThisWorkbook.Sheets(strConfiguration).Range(strDiscountCurve).value
    
    Dim row As Range
    ' calc ir cash flows give a scenario
    If curveDic.count = 0 Then
        MsgBox "There are no calibrated curve for the selected scenario. Please check the available curves or re-calibrate the required curves."
        Exit Sub
    End If
    
    Dim buckets() As String
    Dim effectiveBuckets() As String
    buckets = getBuckets(ThisWorkbook.Sheets(strConfiguration).Range(strGAPBuckets))
    Dim regulatoryCapital As Double
    Dim tier1Capital As Double
    regulatoryCapital = ThisWorkbook.Sheets(strConfiguration).Range(strRegulatoryCapital).value
    tier1Capital = ThisWorkbook.Sheets(strConfiguration).Range(strTier1Capital).value
    
    
    Dim assetDailyCashFlows As Dictionary
    Dim liabilityDailyCashFlows As Dictionary
    
    Dim scenarios As New Dictionary
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets(strConfiguration).Range(strScenarios)
    Dim scen As Variant
    For Each scen In rng
        scenarios.Add scen.value, True
    Next scen
    Dim scenAssetCashFlows As New Dictionary
    Dim scenLiabilityCashFlows As New Dictionary
    Dim assetBookValues() As Double
    assetBookValues = pfAssets.getTotalBookValues()
    Dim liabilityBookValues() As Double
    liabilityBookValues = pfLiabilities.getTotalBookValues()
    Dim scenDiscCurve As New Dictionary
    For Each scen In scenarios.Keys
        pfAssets.calcInstrIrCF curveDic, discCurve, (scen)
        pfLiabilities.calcInstrIrCF curveDic, discCurve, (scen)
        Set assetDailyCashFlows = pfAssets.calcCashFlowMatrix()
        Set liabilityDailyCashFlows = pfLiabilities.calcCashFlowMatrix()
    
        scenAssetCashFlows.Add scen, pfAssets.aggregateCashFlowMatrix(assetDailyCashFlows, buckets, effectiveBuckets)
        scenLiabilityCashFlows.Add scen, pfLiabilities.aggregateCashFlowMatrix(liabilityDailyCashFlows, buckets, effectiveBuckets)
        
        scenDiscCurve.Add scen, curveDic(getScenarioCurveName(discCurve, (scen)))
    Next scen
    Dim nii As New Dictionary
        
    Call printSAKIIRRBB("SAKI Overview", "", "", True, insClasses, _
                    scenAssetCashFlows, pfAssets.assetTypeIndices, nii, _
                    scenLiabilityCashFlows, pfLiabilities.assetTypeIndices, nii, _
                    buckets, scenDiscCurve, _
                    regulatoryCapital, tier1Capital)
    
    
End Sub

Sub RunIRRBBReport()
    Dim curveDic As Scripting.Dictionary
    Dim assetsDic As Scripting.Dictionary


    Set curveDic = ReadandCalibrateCurve(True)

    Dim insClasses As Dictionary
    Set insClasses = getInstrumentClasses()
    
    Dim reports As Dictionary
    Set reports = getReports(insClasses)
    Dim reportConfig As Report
    Set reportConfig = reports(irrbb)
        
    
    Dim pfAssets As New clsPortfolio
    Set pfAssets = PortfolioFactory(insClasses, reportConfig.getEnabledInstruments(), curveDic, "all", True, False, True)     ' reads all assets
    Dim pfLiabilities As New clsPortfolio
    Set pfLiabilities = PortfolioFactory(insClasses, reportConfig.getEnabledInstruments, curveDic, "all", False, True, True)    ' reads all liabilities
    
    Dim discCurve As String
    discCurve = ThisWorkbook.Sheets(strConfiguration).Range(strDiscountCurve).value
    
    Dim row As Range
    ' calc ir cash flows give a scenario
    If curveDic.count = 0 Then
        MsgBox "There are no calibrated curve for the selected scenario. Please check the available curves or re-calibrate the required curves."
        Exit Sub
    End If
    
    Dim buckets() As String
    Dim effectiveBuckets() As String
    buckets = getBuckets(ThisWorkbook.Sheets(strConfiguration).Range(strGAPBuckets))
    Dim regulatoryCapital As Double
    Dim tier1Capital As Double
    regulatoryCapital = ThisWorkbook.Sheets(strConfiguration).Range(strRegulatoryCapital).value
    tier1Capital = ThisWorkbook.Sheets(strConfiguration).Range(strTier1Capital).value
    Dim sh() As Date
    Dim shBucket() As String
    ReDim shBucket(0)
    shBucket(0) = ThisWorkbook.Sheets(strConfiguration).Range(strSimulationHorizonNII).value
    sh = pfAssets.getDtBucketsFromString(shBucket)
    Dim simulationHorizon As Date
    simulationHorizon = sh(0)
    

    
    Dim assetDailyCashFlows As Dictionary
    Dim liabilityDailyCashFlows As Dictionary
    
    Dim scenarios As New Dictionary
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets(strConfiguration).Range(strScenarios)
    Dim scen As Variant
    For Each scen In rng
        scenarios.Add scen.value, True
    Next scen
    Dim scenAssetCashFlows As New Dictionary
    Dim scenLiabilityCashFlows As New Dictionary
    Dim assetBookValues() As Double
    assetBookValues = pfAssets.getTotalBookValues()
    Dim liabilityBookValues() As Double
    liabilityBookValues = pfLiabilities.getTotalBookValues()
    Dim scenDiscCurve As New Dictionary
    Dim scenAssetNii As New Dictionary
    Dim assetNii() As Double
    Dim liabilityNii() As Double
    Dim scenLiabilityNii As New Dictionary
    Dim numBuckets() As Double ' buckets as years (in act/365)
    numBuckets = getBucketLengths(buckets)
    For Each scen In scenarios.Keys
        pfAssets.calcInstrIrCF curveDic, discCurve, (scen)
        pfLiabilities.calcInstrIrCF curveDic, discCurve, (scen)
        Set assetDailyCashFlows = pfAssets.calcCashFlowMatrix()
        Set liabilityDailyCashFlows = pfLiabilities.calcCashFlowMatrix()
        'assetNii = pfAssets.calcNiiMatrixtest(numBuckets)
        
        assetNii = pfAssets.calcNiiMatrix(buckets, simulationHorizon, curveDic, discCurve, (scen))
        liabilityNii = pfLiabilities.calcNiiMatrix(buckets, simulationHorizon, curveDic, discCurve, (scen))
        
        scenAssetCashFlows.Add scen, pfAssets.aggregateCashFlowMatrix(assetDailyCashFlows, buckets, effectiveBuckets)
        scenLiabilityCashFlows.Add scen, pfLiabilities.aggregateCashFlowMatrix(liabilityDailyCashFlows, buckets, effectiveBuckets)
        scenAssetNii.Add scen, assetNii
        scenLiabilityNii.Add scen, liabilityNii
        scenDiscCurve.Add scen, curveDic(getScenarioCurveName(discCurve, (scen)))
    Next scen
    
        
    Call printSAKIIRRBB("IRRBB Overview", "IRRBB Delta EVE Detail", "IRRBB Delta NII Detail", False, insClasses, _
                    scenAssetCashFlows, pfAssets.assetTypeIndices, scenAssetNii, _
                    scenLiabilityCashFlows, pfLiabilities.assetTypeIndices, scenLiabilityNii, _
                    buckets, scenDiscCurve, _
                    regulatoryCapital, tier1Capital)
    
    
End Sub


Sub CalibrateCurve(Optional scen As String = "default")
'content:
'Sub calibrates and prints the calibrated curves to sheet
'Parameters:
' scen: string: indicates the applicable scenario, if not provided all curves will be calibrated, if provided empty no scenario will be applied, scen name must extist

Dim curveDic As Scripting.Dictionary
    If scen = "default" Then Set curveDic = ReadandCalibrateCurve(True, ThisWorkbook.Sheets(strConfiguration).Range(strScenario).value) Else Set curveDic = ReadandCalibrateCurve(True, scen)
End Sub
Sub CalculateIRBalance(Optional scen As String = "defaults")
'content:
'Sub calculates the ir balance for the given scenario and prints the balance(s) to sheet
'Parameters:
' scen: string: indicates the applicable scenario, if not provided ir balances will be provided for all scenarios, if provided empty no scenario will be applied, scen name must extist

    Dim curveDic As Scripting.Dictionary, assetsDic As Scripting.Dictionary, discCurve As String, scenBatch As String, row As Range
    Set curveDic = ReadCurvesFromSheet()    ' reads all calibrated curves from sheet
'    Set AssetsDic = InstrumentFactory(curveDic)     ' reads all instruments
    Dim pf As clsPortfolio
    Dim insClasses As Dictionary
    Set insClasses = getInstrumentClasses()
    Dim enableAll As New Dictionary
    Dim k As Variant
    For Each k In insClasses.Keys
        enableAll.Add k, True
    Next k
    Set pf = PortfolioFactory(insClasses, enableAll, curveDic, "")     ' reads all instruments
    Set assetsDic = pf.PFInstruments()
    discCurve = ThisWorkbook.Sheets(strConfiguration).Range(strDiscountCurve).value
    ' calc ir cash flows give a scenario
    If curveDic.count = 0 Then
        MsgBox "There are no calibrated curve for the selected scenario. Please check the available curves or re-calibrated the required curves."
        Exit Sub
    End If
    If assetsDic.count = 0 Then
        MsgBox "There are no instruments. Please check the instrument selection."
        Exit Sub
    End If
    
    If scen = "defaults" Then
        scen = ThisWorkbook.Sheets(strConfiguration).Range(strScenario).value
        calcIRCF assetsDic, curveDic, discCurve, scen
        ' print ir cash flows
        printIRCF assetsDic
    Else
        For Each row In ThisWorkbook.Sheets(strConfiguration).Range(strScenarios).Rows
            scenBatch = row.value
            calcIRCF assetsDic, curveDic, discCurve, scenBatch
            ' print ir cash flows
            printIRCF assetsDic
        Next
    End If
End Sub

Sub CalculatePortfolio(Optional scen As String = "")
'content:
'Sub calculates the ir balance for the given scenario and prints the balance(s) to sheet
'Parameters:
' scen: string: indicates the applicable scenario, if not provided ir balances will be provided for all scenarios, if provided empty no scenario will be applied, scen name must extist

    Dim curveDic As Scripting.Dictionary, pf As New clsPortfolio, discCurve As String, scenBatch As String, row As Range
    Set curveDic = ReadCurvesFromSheet()    ' reads all calibrated curves from sheet
    discCurve = ThisWorkbook.Sheets(strConfiguration).Range(strDiscountCurve).value
    Dim insClasses As Dictionary
    Set insClasses = getInstrumentClasses()
    Dim enableAll As New Dictionary
    Dim k As Variant
    For Each k In insClasses.Keys
        enableAll.Add k, True
    Next k
    Set pf = PortfolioFactory(insClasses, enableAll, curveDic, scen)     ' reads all instruments
    ' calc ir cash flows give a scenario
    If curveDic.count = 0 Then
        MsgBox "There are no calibrated curve for the selected scenario. Please check the available curves or re-calibrate the required curves."
        Exit Sub
    End If
'    If pf.Count = 0 Then
'        MsgBox "There are no instruments. Please check the instrument selection."
'        Exit Sub
'    End If
    
    Dim cashFlows() As Double
    Dim myBuckets1() As String
    Dim mybuckets2() As String
    Dim mybuckets3() As String
    Dim effectiveBuckets() As String
    myBuckets1 = getBuckets(ThisWorkbook.Sheets(strConfiguration).Range(strGAPBuckets))
    Dim bstr As String
    Dim i As Integer
    bstr = "1M"
    For i = 2 To 120
        bstr = bstr + " " + CStr(i) + "M"
    Next i
    mybuckets2 = split(bstr)
    mybuckets3 = split("")

    Dim dailyCashFlows As Dictionary

    If scen = "" Then
        'Worksheets("TEst").Cells(1, 1) = Now()
        scen = ThisWorkbook.Sheets(strConfiguration).Range(strScenario).value
        pf.calcInstrIrCF curveDic, discCurve, scen
    
        Set dailyCashFlows = pf.calcCashFlowMatrix()
        cashFlows = pf.aggregateCashFlowMatrix(dailyCashFlows, myBuckets1, effectiveBuckets)
        Call printDefaultBucketedCashFlowMatrix(scen, 0, cashFlows, pf.AssettypIndices, effectiveBuckets)
        
        Call printBucketedCashFlowMatrix(scen, "IR Gap 1", 0, cashFlows, pf.AssettypIndices, effectiveBuckets)
        cashFlows = pf.aggregateCashFlowMatrix(dailyCashFlows, mybuckets2, effectiveBuckets)
        Call printBucketedCashFlowMatrix(scen, "IR Gap 2", 0, cashFlows, pf.AssettypIndices, effectiveBuckets)
        cashFlows = pf.aggregateCashFlowMatrix(dailyCashFlows, mybuckets3, effectiveBuckets)
        Call printBucketedCashFlowMatrix(scen, "IR Gap 3", 0, cashFlows, pf.AssettypIndices, effectiveBuckets)
        'Worksheets("TEst").Cells(2, 1) = Now()
    
    Else
        'Worksheets("TEst").Cells(1, 1) = Now()
        i = 0
        For Each row In ThisWorkbook.Sheets(strConfiguration).Range(strScenarios).Rows
            scenBatch = row.value
            pf.calcInstrIrCF curveDic, discCurve, scenBatch
            
            Set dailyCashFlows = pf.calcCashFlowMatrix()
                      
            cashFlows = pf.aggregateCashFlowMatrix(dailyCashFlows, myBuckets1, effectiveBuckets)
            Call printDefaultBucketedCashFlowMatrix(scenBatch, i, cashFlows, pf.AssettypIndices, effectiveBuckets)
            
            Call printBucketedCashFlowMatrix(scenBatch, "IR Gap 1", i, cashFlows, pf.AssettypIndices, effectiveBuckets)
            cashFlows = pf.aggregateCashFlowMatrix(dailyCashFlows, mybuckets2, effectiveBuckets)
            Call printBucketedCashFlowMatrix(scenBatch, "IR Gap 2", i, cashFlows, pf.AssettypIndices, effectiveBuckets)
            cashFlows = pf.aggregateCashFlowMatrix(dailyCashFlows, mybuckets3, effectiveBuckets)
            Call printBucketedCashFlowMatrix(scenBatch, "IR Gap 3", i, cashFlows, pf.AssettypIndices, effectiveBuckets)
            
            i = i + 1
        Next
        'Worksheets("TEst").Cells(2, 1) = Now()
    End If
End Sub



