VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPortfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ****************** Variables *****************************************
Private strPFName As String
Private dtPFDate As Date

Private dicInstruments As New Scripting.Dictionary             ' key: instrument name, value: instrument class
Private dicInstrumentIndices As New Scripting.Dictionary       ' key: position index of instrument within portfolio, value: instrument name
Private dicAssetTypeIndices As New Scripting.Dictionary        ' key: instrument type (string) (e.g. "Wholesale"), value: last instrument index of this type in dicInstrument
Public assetTypeIndices As Dictionary                         ' key: instrument type (enum InstrumentType), value: int(0 to 1), range of relevant instrument indices

Private strBuckets() As String                                 ' Contains the different IR buckets

' calculated data
Private dblAssetBucketMatrix() As Double                        ' [#buckets][#instrument types]


'***********************************************************************
'               Properties
'***********************************************************************
Property Let name(name As String)
    strPFName = name
End Property
Property Let PFDate(dt As Date)
    dtPFDate = dt
End Property
Property Let PFInstruments(inst As Dictionary)
    Set dicInstruments = inst
End Property
Property Get PFInstruments() As Dictionary
    Set PFInstruments = dicInstruments
End Property
Property Let PFInstrumentsInd(instind As Dictionary)
    Set dicInstrumentIndices = instind
End Property
Property Let AssettypIndices(assettypind As Dictionary)
    Set dicAssetTypeIndices = assettypind
End Property
Property Get AssettypIndices() As Dictionary
    Set AssettypIndices = dicAssetTypeIndices
End Property

Property Let buckets(bckts As Range)
Dim i As Integer
    ReDim strBuckets(bckts.count - 1)
    For i = 0 To bckts.count - 1
        strBuckets(i) = bckts.Cells(i + 1, 1)
    Next
End Property



'***********************************************************************
'               Methods
'***********************************************************************

'***************************** calculate IR cash flows for all instruments of the portfolio ******************************
Sub calcInstrIrCF(curveDic As Dictionary, discCurve As String, scen As String)
' Content:
' CashFlows for all instrumens of the portfolio are calculated. CFs include ir as well as amortization cashflows
' NPV and discounted CFs are calculated as well
' calculated data is stored within the rate curve object
' Parameter: Assets - dictionary with all assets for which ir cf are calculated, key name corresponds each asset's name
' Parameter: Curve - clsRateCurve for discouting and forward calculation

Dim c As Variant
    For Each c In dicInstruments.Keys
        dicInstruments(c).calcCF curveDic, discCurve, scen
    Next
End Sub

'***************************** calc full matrix ********************************************


'***************************** calc full matrix ********************************************




Function calcCashFlowMatrix() As Dictionary

' Routine determines for all instruments from all instrument types cash flow data, orders them by date and aggregates them according to the IR GAP bucket structure
' i.e. routine calclulates data for members dblAssetBucketMatrix
' Uses:
' - members: dicInstruments, dicAssetTypeIndices
' - other methods: printCF method of each instrument class
'On Error GoTo errDetails
    If dicInstruments.count = 0 Then
        Exit Function
    End If
    Dim dataExist As Variant
    Dim ins As Variant ' loop over instruments
    Dim result As New Scripting.Dictionary
    Dim i As Integer, j As Integer
    Dim data() As clsCashFlow
    
    Dim cfData() As Double
    Dim cashFlow As Variant
    Dim insToClass As New Dictionary
    Dim insId As Integer ' index of current instument
    Dim insAssetClassId As Integer ' index of asset class of current instrument
    j = 0
    For i = 0 To dicInstruments.count - 1
        Do While j < dicAssetTypeIndices.count - 1 And i > dicAssetTypeIndices.Items()(j)
            j = j + 1
        Loop
        insToClass.Add i, j
    Next i
    
    Dim bucketEntry() As Double
    ReDim bucketEntry(dicAssetTypeIndices.count * nrOfCfElements - 1)
    
    insId = 0
    For Each ins In dicInstruments.Items
        data = ins.getCF()
        insAssetClassId = insToClass(insId)
                
        For Each cashFlow In data
            cfData = cashFlow.getAll()
            If result.Exists(cashFlow.dt) Then
                dataExist = result(cashFlow.dt)
                For i = 0 To nrOfCfElements - 1
                    dataExist(insAssetClassId * nrOfCfElements + i) = dataExist(insAssetClassId * nrOfCfElements + i) + cfData(i)
                Next
                result(cashFlow.dt) = dataExist
            Else
                For i = 0 To nrOfCfElements - 1
                    bucketEntry(insAssetClassId * nrOfCfElements + i) = cfData(i)
                Next
                result.Add cashFlow.dt, bucketEntry
                ReDim bucketEntry(dicAssetTypeIndices.count * nrOfCfElements - 1)        ' re-initialize to zeros
            End If
        Next
        insId = insId + 1
    Next
    Set calcCashFlowMatrix = funcSortKeysAsc(result)
    
    Exit Function
    
errDetails:
    MsgBox err.Number & ", " & err.Source & ", " & err.Description
    
End Function




Function calcNiiMatrixtest(buckets() As Double) As Double()
    Dim result() As Double
    ReDim result(3)
    calcNiiMatrixtest = result
End Function


Function calcNiiMatrix(buckets() As String, simulationHorizon As Date, curveDic As Dictionary, discCurve As String, scenName As String) As Double()
    
    If assetTypeIndices.count = 0 Then
        Exit Function
    End If
    Dim nrBuckets As Integer
    nrBuckets = UBound(buckets) + 1
    
    Dim bucketDates() As Date
    bucketDates = getDtBucketsFromString(buckets)
    Dim result() As Double
    ReDim result(assetTypeIndices.count - 1, UBound(buckets))
    
    Dim assetTypeResult() As Double
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim tp As Variant
    j = 0
    For Each tp In assetTypeIndices.Keys
        ReDim assetTypeResult(nrBuckets - 1)
        
        For i = assetTypeIndices(tp)(0) To assetTypeIndices(tp)(1) - 1
            dicInstruments.Items(i).addNii assetTypeResult, bucketDates, simulationHorizon, curveDic, discCurve, scenName
        Next i
        For k = 0 To UBound(result, 2)
            result(j, k) = assetTypeResult(k)
        Next k
        j = j + 1
    Next tp
    calcNiiMatrix = result

End Function

    

'***************************** convert strBuckets to dtBuckets ********************************************
Function getDtBuckets() As Date()
' Function calculates period start and end dates given a IR GAP bucket structure, and fills member dtBuckets
' Uses:
' - members: dtPfDate, strBuckets()
' Ouput: date vector with bucket dates
    Dim i As Integer
    Dim dtBuckets() As Date
    ReDim dtBuckets(UBound(strBuckets))
    For i = 0 To UBound(strBuckets)
        Select Case Right(strBuckets(i), 1)
        Case "M", "m"
            dtBuckets(i) = DateAdd("d", -1, DateAdd("m", Left(strBuckets(i), Len(strBuckets(i)) - 1), DateAdd("d", 1, dtPFDate)))
        Case "Y", "y"
            dtBuckets(i) = DateAdd("d", -1, DateAdd("m", Left(strBuckets(i), Len(strBuckets(i)) - 1) * 12, DateAdd("d", 1, dtPFDate)))
        Case "N", "n" ' O/N - Bucket
            dtBuckets(i) = DateAdd("d", 1#, dtPFDate)
        Case Else
            MsgBox "Please double check the notation of the rate buckets in the configuration sheet."
        End Select
    Next
    getDtBuckets = dtBuckets
    
End Function

'***************************** convert buckets (strings) to dtBuckets (dates) ********************************************
Public Function getDtBucketsFromString(buckets() As String) As Date()
' Function calculates period start and end dates given a IR GAP bucket structure, and fills member dtBuckets
' Uses:
' - members: dtPfDate
' Ouput: date vector with bucket dates
    Dim i As Integer
    Dim dtBuckets() As Date
    ReDim dtBuckets(UBound(buckets))
    For i = 0 To UBound(buckets)
        Select Case Right(buckets(i), 1)
        Case "M", "m"
            dtBuckets(i) = DateAdd("d", -1, DateAdd("m", Left(buckets(i), Len(buckets(i)) - 1), DateAdd("d", 1, dtPFDate)))
        Case "Y", "y"
            dtBuckets(i) = DateAdd("d", -1, DateAdd("m", Left(buckets(i), Len(buckets(i)) - 1) * 12, DateAdd("d", 1, dtPFDate)))
        Case "N", "n" ' O/N - Bucket
            dtBuckets(i) = DateAdd("d", 1#, dtPFDate)
        Case Else
            MsgBox "Please double check the notation of the rate buckets in the configuration sheet."
        End Select
    Next
    getDtBucketsFromString = dtBuckets
    
End Function

'******************************* Printing *************************************************



Function aggregateCashFlowMatrix(cashFlows As Dictionary, buckets() As String, effectiveBuckets() As String) As Double()
' Function aggregates cash flows into a given bucket structure
' Input:
'   cashFlows: dictionary of all cash flows, indexed by date
'   buckets: list of bucket names, or empty list -> use all cash flow dates as buckets
' Output:
'   effectiveBuckets: returns the list of bucket descriptions
'   function return value: three-dimensional array result[i][j][k] with
'      i = bucket index
'      j = instrument type index
'      k = cash flow type index
'On Error GoTo errDetails

    If cashFlows Is Nothing Then
        Exit Function
    End If
    Dim dtBuckets() As Date
    Dim bucketIndex As Integer
    Dim instrTypeIndex As Integer
    Dim dblTemp() As Double
    
    Dim cfDate As Variant
    Dim k As Integer
    Dim result() As Double
    

    If UBound(buckets) >= 0 Then ' buckets provided
        dtBuckets = getDtBucketsFromString(buckets)
        effectiveBuckets = buckets
    Else ' no buckets provided -> finest granularity, keep all cash flow dates as buckets
        Dim nEffectiveBuckets As Integer
        nEffectiveBuckets = 0
        For Each cfDate In cashFlows.Keys
            If cfDate > dtPFDate Then
                nEffectiveBuckets = nEffectiveBuckets + 1
            End If
        Next cfDate
        ReDim dtBuckets(0 To nEffectiveBuckets - 1)
        ReDim effectiveBuckets(0 To nEffectiveBuckets - 1)
        k = 0
        For Each cfDate In cashFlows.Keys
            If cfDate > dtPFDate Then
                dtBuckets(k) = cfDate
                effectiveBuckets(k) = CStr(cfDate)
                k = k + 1
            End If
        Next cfDate
    End If
    
    ReDim result(UBound(dtBuckets), dicAssetTypeIndices.count - 1, nrOfCfElements - 1)
    
    bucketIndex = 0                     ' index for target array, dimension bucket
    
    ReDim dblTemp(dicAssetTypeIndices.count * nrOfCfElements - 1)
    For Each cfDate In cashFlows.Keys()
        If cfDate > dtPFDate Then ' only take into account cashflows after portfolio date
            Do While bucketIndex < UBound(dtBuckets) And cfDate > dtBuckets(bucketIndex)
                bucketIndex = bucketIndex + 1
            Loop
            dblTemp = cashFlows(cfDate)
            
            For instrTypeIndex = 0 To dicAssetTypeIndices.count - 1
                For k = 0 To nrOfCfElements - 1
                    result(bucketIndex, instrTypeIndex, k) = result(bucketIndex, instrTypeIndex, k) + dblTemp(instrTypeIndex * nrOfCfElements + k)
                Next k
                
            Next
        End If
    Next
    
    aggregateCashFlowMatrix = result
    Exit Function
    
errDetails:
    MsgBox err.Number & ", " & err.Source & ", " & err.Description
    
End Function

Function getTotalNotionals() As Double()

    Dim result() As Double
    ReDim result(assetTypeIndices.count - 1)
    Dim notional As Double
    Dim i As Integer
    Dim j As Integer
    Dim tp As Variant
    j = 0
    For Each tp In assetTypeIndices.Keys
        notional = 0
        For i = assetTypeIndices(tp)(0) To assetTypeIndices(tp)(1) - 1
            notional = notional + dicInstruments.Items(i).getNominal()
        Next i
        result(j) = notional
        j = j + 1
    Next tp
    getTotalNotionals = result

End Function

Function getTotalBookValues() As Double()

    Dim result() As Double
    If assetTypeIndices.count = 0 Then
        Exit Function
    End If
    ReDim result(assetTypeIndices.count - 1)
    
    Dim bookValue As Double
    Dim i As Integer
    Dim j As Integer
    Dim tp As Variant
    j = 0
    
    
    For Each tp In assetTypeIndices.Keys
        bookValue = 0
        For i = assetTypeIndices(tp)(0) To assetTypeIndices(tp)(1) - 1
            bookValue = bookValue + dicInstruments.Items(i).bookValue
        Next i
        result(j) = bookValue
        j = j + 1
    Next tp
    getTotalBookValues = result

End Function

