Attribute VB_Name = "mdl_Globalconst"
Option Explicit

'************* Moduel containes Public Const variables ***************************

' Data Sheets
Public Const strConfiguration As String = "Configuration"
Public Const strMarketData As String = "Market_Scenario_Data"
Public Const strWholesaleData As String = "RetailWholesale data"
Public Const strABSData As String = "ABS data"
Public Const strIntercompanyLoanData As String = "Intercompanyloans data"
Public Const strDepositData As String = "Cash_Deposit data"
Public Const strSwapData As String = "Swaps data"
Public Const strIllustrateCurve As String = "Illustrate_Curves"
Public Const strCurveDataCalibrated As String = "CurveDataCalibrated"
Public Const strIRBalance As String = "IR Balance"
Public Const strPFCF As String = "PF Cashflows"
Public Const strIRGAP As String = "PF IR GAPs"

' Tables and Names - Configuration
Public Const strDiscountCurve As String = "DiscountCurve"
Public Const strRefDate As String = "RefDate"
Public Const strScenario As String = "CurrentScenario"
Public Const strCurrentPosition As String = "CurrentPosition"
Public Const strPositions As String = "POSITIONS"
Public Const strRateCurves As String = "RateCurves"
Public Const strScenarios As String = "Scenarios"
Public Const strHolidays As String = "Holidays"
Public Const strAmortizationSchemes As String = "AmortizationSchemes"
Public Const strFloor As String = "Floor"
Public Const strSlope As String = "Slope"
Public Const strRetailAccounts As String = "RetailSubDivs"
Public Const strWholesaleAccounts As String = "WholesaleSubDivs"
Public Const strLeasingAccounts As String = "LeasingSubDivs"
Public Const strCurveInstruments As String = "CurveInstruments"
Public Const strRetailLC As String = "Retail_LC"
Public Const strWholesaleLC As String = "Wholesale_LC"
Public Const strAvailableCurves As String = "AvailableCurves"
Public Const strGAPBuckets As String = "GAPBuckets"
Public Const strDefaultBDC As String = "defaultBDC"
Public Const strDefaultDCC As String = "defaultDCC"
Public Const strOffSet As String = "OffSet"
Public Const strAdminCosts As String = "AdminCosts"
Public Const strFixedCostsThreshold As String = "FixedCostsThreshold"
Public Const strFixedCostsAmount As String = "FixedCostsAmount"
Public Const strSyntheticCosts As String = "SyntheticCosts"
Public Const strEnabledInstruments = "EnabledInstruments"
Public Const strRegulatoryCapital As String = "RegulatoryCapital"
Public Const strTier1Capital = "Tier1Capital"
Public Const strSimulationHorizonNII = "SimulationHorizonNII"

' Tables and Names - Market & Scenario Data
Public Const strEUR1M As String = "EURIBOR_1M"
Public Const strEUR3M As String = "EURIBOR_3M"
Public Const strTenorScenarios As String = "Tenor_Scenarios"
Public Const strParallelUp As String = "ParallelUp"
Public Const strParallelDown As String = "ParallelDown"
Public Const strSteepening As String = "Steepening"
Public Const strFlattening As String = "Flattening"
Public Const strShortTermUp As String = "ShortTermUp"
Public Const strShortTermDown As String = "ShortTermDown"

' Tables and Names - RetailWholesale
Public Const strBSData As String = "BS_DATA"
Public Const strRetailWholesaleData As String = "RetailWholesale_DATA"

' Tables and Names - ABS
Public Const strABSAssetsData As String = "ABS_Assets"
Public Const strABSLiabilitiesData As String = "ABS_Liabilities"

' Tables and Names - Intercompany Loans
Public Const strIntercompanyLoansData As String = "IntercompanyLoans"

' Tables and Names - Cash & Deposit
Public Const strDepositsData As String = "DEPOSITS"
Public Const strECBCash As String = "ECB_Cash"
Public Const strCash As String = "Cash"
Public Const strCashData As String = "CASH_DATA"

' Tables and Names - Swaps
Public Const strSwapsData As String = "SWAPS"

' Tables and Names - IllustrateCurves
Public Const strIllustrationCurve As String = "IllustrationCurve"

' ****************** further constants ****************
Public Const intSizeCurveGrid As Integer = 120
Public Const intDiscCurveFreq As Integer = 3
Public Const intNrOfCFData As Integer = 5

Public Const nrOfCfElements As Integer = 6

' **** cashflow elements
Public Const cfTotalCf = 4
Public Const cfTotalNpv = 5


