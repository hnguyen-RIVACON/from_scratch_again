Attribute VB_Name = "mdl_Setup"
Option Explicit

' This enum lists all instrument types
Public Enum InstrumentType
    ECBCash = 1
    Cash = 2
    RetailCommitment = 3
    Retail = 4
    ECBTender = 5
    Wholesale = 6
    ABSRetainedNotes = 7
    IntercompanyLoans = 8
    ABSSynthLiabilities = 9
    Swap = 10
    DepositFix = 11
    ABSSwapPayLeg = 12
    ABSSwapReceiveLeg = 13
    ALMSwapPayLeg = 14
    ALMSwapReceiveLeg = 15
    Leasing = 16
    DepositFlex = 17
End Enum

Public Enum ReportType
    BFA3Report = 1
    saki = 2
    irrbb = 3
End Enum


Function getAdminCosts(sn As String) As Double
    Dim costRange As Range
    Set costRange = ThisWorkbook.Sheets(strConfiguration).Range(strAdminCosts)
    Dim row As Range
    Dim result As Double
    result = -1#
    For Each row In costRange.Rows
        If row.Cells(1, 1).value = sn Then
            result = row.Cells(1, 2).value
        End If
    Next row
    If result = -1# Then
        MsgBox ("No admin costs specified for " & sn & ", using 0")
        result = 0#
    End If
    getAdminCosts = result
End Function


Function getInstrumentClasses() As Dictionary
    Dim classes As New Dictionary
    Dim c As InstrumentClass
    
    Set c = New InstrumentClass
    c.init ECBCash, "ECBCash", "ECB Cash", True, False
    classes.Add ECBCash, c
    
    Set c = New InstrumentClass
    c.init Cash, "Cash", "Cash", True, False
    classes.Add Cash, c

    Set c = New InstrumentClass
    c.init RetailCommitment, "RetailCommitment", "Retail Commitments", True, False
    c.costRatio = getAdminCosts(c.shortName)
    classes.Add RetailCommitment, c

    Set c = New InstrumentClass
    c.init Retail, "Retail", "Retail Receivables", True, False
    c.costRatio = getAdminCosts(c.shortName)
    classes.Add Retail, c

    Set c = New InstrumentClass
    c.init Wholesale, "Wholesale", "Wholesale Receivables", True, False
    c.costRatio = getAdminCosts(c.shortName)
    classes.Add Wholesale, c

    Set c = New InstrumentClass
    c.init ECBTender, "ECBTender", "ECB Tender", False, True
    classes.Add ECBTender, c
    
    Set c = New InstrumentClass
    c.init ABSRetainedNotes, "ABSRetainedNotes", "ABS (Retained Notes)", True, False
    classes.Add ABSRetainedNotes, c
    
    Set c = New InstrumentClass
    c.init ABSSynthLiabilities, "ABSSynthLiabilities", "ABS", False, True
    classes.Add ABSSynthLiabilities, c
    
    Set c = New InstrumentClass
    c.init IntercompanyLoans, "IntercompanyLoans", "Intercompany Loans", False, True
    classes.Add IntercompanyLoans, c
    
    Set c = New InstrumentClass
    c.init Swap, "Swap", "Swap", True, True
    classes.Add Swap, c
    
    Set c = New InstrumentClass
    c.init ALMSwapPayLeg, "ALMSwapPayLeg", "ALM Swap Pay Leg", False, True
    classes.Add ALMSwapPayLeg, c
    
    Set c = New InstrumentClass
    c.init ALMSwapReceiveLeg, "ALMSwapReceiveLeg", "ALM Swap Receive Leg", True, False
    classes.Add ALMSwapReceiveLeg, c
    
    Set c = New InstrumentClass
    c.init ABSSwapPayLeg, "ABSSwapPayLeg", "ABS Swap Pay Leg", False, True
    classes.Add ABSSwapPayLeg, c
    
    Set c = New InstrumentClass
    c.init ABSSwapReceiveLeg, "ABSSwapReceiveLeg", "ABS Swap Receive Leg", True, False
    classes.Add ABSSwapReceiveLeg, c
    
    Set c = New InstrumentClass
    c.init Leasing, "Leasing", "Lease Assets", True, False
    'Begin, Version 0.19, Inclusion of calculation formula getAdminCosts/Integration of Modelling OpCosts for Leasing, 09/01/2024, Marie Konrad'
    c.costRatio = getAdminCosts(c.shortName)
    'End, Version 0.19, Inclusion of calculation formula getAdminCosts/Integration of Modelling OpCosts for Leasing, 09/01/2024, Marie Konrad'
    classes.Add Leasing, c
    
    Set c = New InstrumentClass
    c.init DepositFix, "DepositFix", "Deposits Festgeld", False, True
    classes.Add DepositFix, c
    
    Set c = New InstrumentClass
    c.init DepositFlex, "DepositFlex", "Deposits Tagesgeld", False, True
    classes.Add DepositFlex, c
    
    Set getInstrumentClasses = classes
End Function

Function getReports(instrumentClasses As Dictionary) As Dictionary
    Dim reports As New Dictionary
    Dim c As Report
    Set c = New Report
    c.init BFA3Report, "BFA3", "BFA3 Report", instrumentClasses
    reports.Add BFA3Report, c
    
    Set c = New Report
    c.init saki, "Saki", "Saki", instrumentClasses
    reports.Add saki, c
    
    Set c = New Report
    c.init irrbb, "IRRBB", "IRRBB", instrumentClasses
    reports.Add irrbb, c
    
    Set getReports = reports
    
End Function





Function getTypeFromShortName(classes As Dictionary, sn As String) As InstrumentType
    Dim k As Variant
    For Each k In classes.Keys
        If classes(k).shortName = sn Then
            getTypeFromShortName = k
            Exit Function
        End If
    Next k
End Function
