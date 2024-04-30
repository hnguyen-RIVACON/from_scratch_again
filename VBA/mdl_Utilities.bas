Attribute VB_Name = "mdl_Utilities"
Option Explicit

Private cfNames() As String

'***************************************************************************************************
'                   Lineare Interpolation
Public Function LIP(xVector() As Double, yVector() As Double, xValue As Double) As Double
' Function linearly interpolates from two double vectors
'
' Parameters:
'   xVector, yVector and xValue
'
' Returns: interpolated value

Dim Dimension As Long, m As Double, n As Double, i As Integer ', xValue As Long

'    If yr = 0 Then yr = year(Value)
'    xValue = Value - DateSerial(yr, 1, 1)

    
    Dimension = UBound(xVector)
    If xValue < xVector(0) Or xValue > xVector(Dimension) Then
        If xValue < xVector(0) Then
            m = (yVector(1) - yVector(0)) / (xVector(1) - xVector(0))
            n = yVector(1) - m * xVector(1)
        Else
            m = (yVector(Dimension) - yVector(Dimension - 1)) / (xVector(Dimension) - xVector(Dimension - 1))
            n = yVector(Dimension) - m * xVector(Dimension)
        End If
        LIP = m * xValue + n
    Else
        For i = 2 To Dimension
            If xValue <= xVector(i - 1) Then Exit For
        Next i
        LIP = yVector(i - 2) _
        + (xValue - xVector(i - 2)) / (xVector(i - 1) - xVector(i - 2)) _
        * (yVector(i - 1) - yVector(i - 2))
    End If
End Function


Function getBucketLengths(buckets() As String) As Double()
' Function calculates bucket lengths from string bucket descriptions
' Ouput: date vector with bucket lengths
    Dim i As Integer
    Dim result() As Double
    ReDim result(UBound(buckets))
    For i = 0 To UBound(buckets)
        Select Case LCase(Right(buckets(i), 1))
        Case "m"
            result(i) = CDbl(Left(buckets(i), Len(buckets(i)) - 1)) / 12#
        Case "y"
            result(i) = CDbl(Left(buckets(i), Len(buckets(i)) - 1))
        Case "n" ' O/N - Bucket
            result(i) = 1# / 365#
        Case Else
            MsgBox "Please double check the notation of the rate buckets in the configuration sheet."
        End Select
    Next
    getBucketLengths = result
    
End Function


'***************************** sub prints data structure on predefined worksheet starting in specified cell ***************************
Public Sub printVecOnSheet(data As Variant, sheetName As String, Optional header As String = "Header", Optional offset As Integer = 0)
' sub prints data structure on predefined worksheet starting in specified cell
Dim i As Integer
    
    ThisWorkbook.Sheets(sheetName).Activate
    ' determine first empty column after considering offset
    ActiveSheet.Cells(1, 1 + offset).Activate
    If ActiveSheet.Cells(1, 1 + offset) <> "" Then
        If ActiveSheet.Cells(1, 2 + offset) <> "" Then
            ActiveSheet.Cells(1, Range(Cells(1, 1 + offset), Cells(1, 1 + offset)).End(xlToRight).column + 1).Activate
        Else
           ActiveSheet.Cells(1, 2 + offset).Activate
        End If
    End If
    ActiveCell = Cells(1, ActiveCell.column).Activate
    'print data
    ActiveSheet.Cells(1, ActiveCell.column) = header
    For i = 0 To UBound(data)
        ActiveSheet.Cells(2 + i, ActiveCell.column) = data(i)
    Next
    
End Sub

Function array_Empty(testArr As Variant) As Boolean
' Function returns true if an array is empty, false elsewise
' parameter: testArr as variant, array to be checked
Dim i As Long, k As Long, flag As Long

    On Error Resume Next
    i = UBound(testArr)
    If err.Number = 0 Then
       flag = 0
       For k = LBound(testArr) To UBound(testArr)
          If IsEmpty(testArr(k)) = False Then
             flag = 1
             array_Empty = False
             Exit For
          End If
        Next k
        If flag = 0 Then array_Empty = True
    Else
        array_Empty = True
    End If

End Function

'************************* Date Generator *****************************************
Public Function dtGenerator(dtLastFixing As Date, dtMat As Date, intFreq As Integer, intDCC As Integer, strBDC As String, Optional intOffSet As Integer = 0, Optional paymentDay As Integer = -1) As Date()
' Function returns for a given fixingdate, maturity date, frequency, daycountconvention, business day convention and settlement offset a date matrix of relevant calculation and cash flow dates
'
' Arguments: 
'   maturity:
'   currentdate:
'   frequency:
'   offset:
'   business day convention:
'   day count convention:
'   paymentDay: day in month on which payments are made, or -1 for maturity
'
' ouput: 
'   Array(1. scheduled dates, 2. fixing dates, 3. payment dates)

   Dim i As Integer, i0 As Integer, dtHypDate As Date, dtSch() As Date, dtFix() As Date, dtPay() As Date, dtMatrix() As Date, intNrDates As Integer
    Dim refDate As Date
    Dim stubAtEnd As Boolean
    stubAtEnd = False
    
    refDate = dtMat
    
    If paymentDay > 0 Then ' short stub at end
        refDate = DateSerial(Year(dtMat), Month(dtMat), paymentDay)
        If refDate > dtMat Then
            refDate = DateSerial(Year(dtMat), Month(dtMat) - 1, paymentDay) ' wrap-around at January handled correctly by DateSerial
        End If
        stubAtEnd = BusinessDate(dtMat, strBDC) > BusinessDate(refDate, strBDC)
    End If
        
    i0 = 0
    
    If intFreq = 0 Then
        intNrDates = 1 'intFreq=0: interest at maturity
    Else
        intNrDates = WorksheetFunction.Max(1, ((Year(refDate) - Year(dtLastFixing)) * 12 + Month(refDate) - Month(dtLastFixing) + (intFreq - 1)) \ intFreq)   ' calculate n0 of month and divide by payment frequency assumingn frequency is given in full months
    End If
    If stubAtEnd Then
        intNrDates = intNrDates + 1
        i0 = 1
    End If
    ReDim dtMatrix(intNrDates, 2)
    
    If stubAtEnd Then
        dtMatrix(intNrDates, 0) = dtMat
        dtMatrix(intNrDates, 1) = BusinessDate(dtMat, strBDC)
        dtMatrix(intNrDates, 2) = BusinessDate(DateAdd("d", intOffSet, dtMat), strBDC)
    End If
    
    i = i0
    dtHypDate = refDate
    
    Do
        If i > i0 Then
            If intFreq = 0 Then
                dtHypDate = dtLastFixing
            Else
                dtHypDate = DateAdd("d", -1, DateAdd("m", -intFreq, DateAdd("d", 1, dtHypDate)))          ' payment frequency is giiven in month with a minimum of 1
            End If
        End If
        dtMatrix(intNrDates - i, 0) = dtHypDate
        dtMatrix(intNrDates - i, 1) = BusinessDate(dtHypDate, strBDC)
        dtMatrix(intNrDates - i, 2) = BusinessDate(DateAdd("d", intOffSet, dtHypDate), strBDC)
        i = i + 1
    Loop While dtHypDate > dtLastFixing
    

    dtGenerator = dtMatrix

End Function
'****************************** Business Day Calculation ***************************
Public Function BusinessDate(dt As Date, strBDC As String) As Date
' Function returns the business date for a given daten and a given business date convention
' function processes conventions ' modified following', 'following', 'actual', 'preceding' and 'modified preceding'
' default is 'actual'
' parameters:
' - date to be checked
' - strBDC: business day convention as string

Dim dtCheck As Date
    dtCheck = dt
    
    Select Case strBDC
    Case "modified following"
        Do While Weekday(dtCheck, vbMonday) > 5 Or isHoliday(dtCheck)
            dtCheck = DateAdd("d", 1, dtCheck)
        Loop
        If Month(dtCheck) > Month(dt) Then dtCheck = BusinessDate(dt, "preceding")
    
    Case "modified preceding"
        Do While Weekday(dtCheck, vbMonday) > 5 Or isHoliday(dtCheck)
            dtCheck = DateAdd("d", -1, dtCheck)
        Loop
        If Month(dtCheck) < Month(dt) Then dtCheck = BusinessDate(dt, "following")
    Case "modified"
        Do While Weekday(dtCheck, vbMonday) > 5 Or isHoliday(dtCheck)
            dtCheck = DateAdd("d", 1, dtCheck)
        Loop
    Case "preceding"
        Do While Weekday(dtCheck, vbMonday) > 5 Or isHoliday(dtCheck)
            dtCheck = DateAdd("d", -1, dtCheck)
        Loop
    Case Else
        'BusinessDate = dtCheck
    End Select
    BusinessDate = dtCheck
End Function

Public Function getDCC(strDCC As String) As Integer
' Function returns an integer to be used in the worksheet function yearfrac to indicate the day count convention to be used
' paraemter strDCC as string: name of the day count convention
Dim dcc As Integer
    Select Case strDCC
    Case "30/360"
        dcc = 4
    Case "act/act", "actual/actual"
        dcc = 1
    Case "act/360", "actual/360"
        dcc = 2
    Case "act/365", "actual/365"
        dcc = 3
    End Select
    getDCC = dcc
End Function



Sub deleteNames(Optional strPart As String = "")
' content:
' deletes all Names in Workbook that contain strPart, if parameter ist not provided do nothing
' parameter:
' strPart: string: string for which each name is searched

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim xName As name
    
    For Each xName In Application.ActiveWorkbook.Names
        If strPart <> "" Then
            If InStr(xName.name, strPart) Then
                xName.Delete
            End If
        End If
    Next
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Public Function funcSortKeysAsc(dctList As Object) As Object
' Sorts an object in ascending order by key
' parameter: dctList as Object, object to be sorted
' Returns: sorted dictionary object
    Dim arrTemp() As Variant
    Dim curKey As Variant
    Dim itX As Integer
    Dim itY As Integer

    'Only sort if more than one item in the dict
    If dctList.count > 1 Then

        'Populate the array
        ReDim arrTemp(dctList.count - 1)
        itX = 0
        For Each curKey In dctList
            arrTemp(itX) = curKey
            itX = itX + 1
        Next

        'Do the sort in the array
        BubbleSort arrTemp
        'Create the new dictionary
        Set funcSortKeysAsc = CreateObject("Scripting.Dictionary")
        Dim c() As Variant
        For itX = 0 To (dctList.count - 1)
            funcSortKeysAsc.Add arrTemp(itX), dctList(arrTemp(itX))

        Next
    Else
        Set funcSortKeysAsc = dctList
    End If
End Function

Sub BubbleSort(MyArray() As Variant)
'Sorts a one-dimensional VBA array from smallest to largest
'using the bubble sort algorithm.
Dim i As Long, j As Long
Dim Temp As Variant
 
For i = LBound(MyArray) To UBound(MyArray) - 1
    For j = i + 1 To UBound(MyArray)
        If MyArray(i) > MyArray(j) Then
            Temp = MyArray(j)
            MyArray(j) = MyArray(i)
            MyArray(i) = Temp
        End If
    Next j
Next i
End Sub

Function isHoliday(dt As Date) As Boolean
' Function returne true if a provided date is a TARGET II holiday, i.e.
' New Year's Eve, God Friday, Easter Monday, Labour Day, Christmas, Go
' parameter: dt date to be checked
Dim bln As Boolean
    
    If Month(dt) = 1 And Day(dt) = 1 Then
        bln = True                                              ' New Year's Eve
    ElseIf Month(dt) = 5 And Day(dt) = 1 Then
        bln = True                                              ' Labour Day
    ElseIf Month(dt) = 12 And (Day(dt) = 26 Or Day(dt) = 25) Then
        bln = True                                              ' Christmas
    ElseIf isEasterHoliday(dt) Then
        bln = True                                              ' God Friday and Easter Monday
    Else
        bln = False
    End If
    isHoliday = bln

End Function

Public Function isEasterHoliday(dt As Date) As Boolean
' Function returns true if a specific date is an Easter Monday or an Easter Friday
' parameter: dt i.e. the date to be checked
Dim x As Integer, EasterDAte As Date
    x = (((255 - 11 * (Year(dt) Mod 19)) - 21) Mod 30) + 21
    EasterDAte = DateSerial(Year(dt), 3, 1) + x + (x > 48) + 6 - ((Year(dt) + Year(dt) / 4 + x + (x > 48) + 1) Mod 7)
    If dt = DateAdd("d", -2, EasterDAte) Or dt = DateAdd("d", 1, EasterDAte) Then isEasterHoliday = True Else isEasterHoliday = False
End Function

Public Function initializeArray(arr() As Variant, val As Variant) As Variant()
' Function initializes an arrray with a specific value
' parameters:
' - arr: array to be initialized
' - val: value to be used
Dim i As Integer

    For i = 0 To UBound(arr)
        arr(i) = val
    Next
    initializeArray = arr
End Function

Public Function CreateCashFlow(dt As Date, interest As Double, amortization As Double, discountFactor As Double) As clsCashFlow

    Set CreateCashFlow = New clsCashFlow
    CreateCashFlow.dt = dt
    CreateCashFlow.interest = interest
    CreateCashFlow.amortization = amortization
    CreateCashFlow.discountFactor = discountFactor

End Function


Public Function getCfNames(i As Integer) As String

    If (Not Not cfNames) = 0 Then
        cfNames = split("IR CF,Disc. IR CF,Am. CF,Disc. Am. CF,CF,Disc. CF", ",")
    End If
    getCfNames = cfNames(i)
End Function

Public Function getBuckets(bckts As Range) As String()
    Dim i As Integer
    Dim result() As String
    ReDim result(bckts.count - 1)
    For i = 0 To bckts.count - 1
        result(i) = bckts.Cells(i + 1, 1)
    Next
    getBuckets = result
End Function



