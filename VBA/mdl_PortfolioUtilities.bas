Attribute VB_Name = "mdl_PortfolioUtilities"
Option Explicit


'************************* Instrument Factory *****************************************

Public Function PortfolioFactory(insClasses As Dictionary, enabledInstruments As Dictionary, curves As Dictionary, Optional scen As String = "", Optional includeAssets As Boolean = True, Optional includeLiabilities As Boolean = True, Optional splitSwaps As Boolean = False) As clsPortfolio
' Function creates an instance of the clsPortfolio class
' by reading all instruments of the applicable instrument types
' Parameters:
' - scen: parameter that controls if the portfolio contains instruments from only one instrument class (i.e. by default) or from all instrument classes


    Dim assets As New Dictionary
    Dim assetsind As New Dictionary
    Dim assetType As New Dictionary
    Dim pf As clsPortfolio
    Dim rng As Range, rngAsset As Range, i As Integer, index As Integer
    
    If scen = "" Then
        Set rng = ThisWorkbook.Sheets(strConfiguration).Range(strCurrentPosition)
    Else
        Set rng = ThisWorkbook.Sheets(strConfiguration).Range(strPositions)
    End If
    index = 0
    Dim insTypeName As String
    Dim insType As InstrumentType
    Dim insClass As InstrumentClass
    Dim assetIndices As New Dictionary
    Dim count(0 To 1) As Integer
    Dim addClass As Boolean
    For i = 1 To rng.Rows.count
        insTypeName = rng.Cells(i, 1).value
        insType = getTypeFromShortName(insClasses, insTypeName)
        addClass = enabledInstruments(insType)
        Set insClass = insClasses(insType)
        If ((includeAssets And insClass.isAsset) Or (includeLiabilities And insClass.isLiability)) Then
            count(0) = assets.count
            Select Case insType
            Case Retail, RetailCommitment, Wholesale, Leasing
                If enabledInstruments(insType) Then
                    RetailWholesaleFactory assets, assetsind, assetType, index, insTypeName, curves
                End If
            Case Cash, ECBCash, ECBTender
                If enabledInstruments(insType) Then
                    CashFactory assets, assetsind, assetType, index, insTypeName
                End If
            Case ABSRetainedNotes, ABSSynthLiabilities
                If enabledInstruments(insType) Then
                    ABSFactory assets, assetsind, assetType, index, insTypeName
                End If
            Case IntercompanyLoans
                If enabledInstruments(insType) Then
                    IntercompanyLoansFactory assets, assetsind, assetType, index
                End If
            Case Swap
                If splitSwaps Then
                    If includeLiabilities Then
                        ' pay legs ALM swaps
                        If enabledInstruments(ALMSwapPayLeg) Then
                            SwapFactory assets, assetsind, assetType, index, True, True, True
                            count(1) = assets.count
                            assetIndices.Add ALMSwapPayLeg, count
                            count(0) = assets.count
                        End If
                        ' pay legs ABS swaps
                        If enabledInstruments(ABSSwapPayLeg) Then
                            SwapFactory assets, assetsind, assetType, index, True, True, False
                            count(1) = assets.count
                            assetIndices.Add ABSSwapPayLeg, count
                            count(0) = assets.count
                        End If
                    End If
                    If includeAssets Then
                        ' receive legs ALM swaps
                        If enabledInstruments(ALMSwapReceiveLeg) Then
                            SwapFactory assets, assetsind, assetType, index, True, False, True
                            count(1) = assets.count
                            assetIndices.Add ALMSwapReceiveLeg, count
                            count(0) = assets.count
                        End If
                        ' receive legs ABS swaps
                        If enabledInstruments(ABSSwapReceiveLeg) Then
                            SwapFactory assets, assetsind, assetType, index, True, False, False
                            count(1) = assets.count
                            assetIndices.Add ABSSwapReceiveLeg, count
                            count(0) = assets.count
                        End If
                    End If
                    addClass = False
                Else
                    SwapFactory assets, assetsind, assetType, index
                End If
            Case DepositFix, DepositFlex
                If enabledInstruments(insType) Then
                    DepositFactory assets, assetsind, assetType, index, insType, insTypeName
                End If
            Case Else
                MsgBox "The instrument " & insTypeName & " has not been setup yet."
            End Select
            If addClass Then
                count(1) = assets.count
                assetIndices.Add insType, count
            End If
        End If
    
    Next
    
    Set pf = New clsPortfolio
    pf.name = "FBG_Standard"
    pf.PFDate = ThisWorkbook.Sheets(strConfiguration).Range(strRefDate)
    pf.buckets = ThisWorkbook.Sheets(strConfiguration).Range(strGAPBuckets)
    pf.PFInstruments = assets
    pf.PFInstrumentsInd = assetsind
    pf.AssettypIndices = assetType
    Set pf.assetTypeIndices = assetIndices
    
    
    Set PortfolioFactory = pf
    Set pf = Nothing
    Set rng = Nothing
End Function

