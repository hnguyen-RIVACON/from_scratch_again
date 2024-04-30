Attribute VB_Name = "CreateTemplates"
Option Explicit



Function AreItemsContinuous(arr() As Variant) As Boolean
' Function: AreItemsContinuous
' Checks if all values in the given array arr() are continuous
'
' Arguments:
'   arr: the array we want to be searched
'
' Returns:
'   Boolean: True or False


    Dim i As Long
    For i = LBound(arr) To UBound(arr) - 1
        If arr(i) + 1 <> arr(i + 1) Then
            AreItemsContinuous = False
            Exit Function
        End If
    Next i
    AreItemsContinuous = True
End Function



Function CreateTemplate(ByVal templateName As String) As TemplateObj
' Creates a TemplateObj depending on which template Identifier is given
' Expected input for the string templateName are:
' "(A1) Umfrage", "(A2) ST-GuV-Kapital", "(A2) ST-Zinsergebnis"
' Returns the appropriate data structure of the desired Template filled
' with example values to be overwritten
'
' Arguments:
'   templateName: the name of the desired Template type
'
' Returns:
'   CreateTemplate: is a TemplateObj with defined structure and dummy values
'
' See Also:
'   [CreateTemplate_Umfrage][CreateTemplate_Umfrage], [CreateTemplate_ST_GuV][CreateTemplate_ST_GuV], [CreateTemplate_ST_Zinsergebnis][CreateTemplate_ST_Zinsergebnis]


    ' Init
    Dim template As TemplateObj
    Set template = New TemplateObj
    
    ' Determine which Template to create
    Debug.Print "Creating TemplateObj for " & templateName
    If templateName = "(A1) Umfrage" Then
        Set template = CreateTemplate_Umfrage(templateName)
        
    ElseIf templateName = "(A2) ST-GuV-Kapital" Then
        Set template = CreateTemplate_ST_GuV(templateName)

    ElseIf templateName = "(A2) ST-Zinsergebnis" Then
        Set template = CreateTemplate_ST_Zinsergebnis(templateName)
    
    Else
        MsgBox "Invalid template name: " & templateName & vbNewLine & _
        "Please select: (A1) Umfrage, (A2) ST-GuV-Kapital, or (A2) ST-Zinsergebnis"
        ' Should we throw an Exception?
    End If
    
    ' Return the templateObj
    Set CreateTemplate = template
    'Debug.Print "Created TemplaltObj: " & templateName
End Function




Function CreateTemplate_Umfrage(ByVal templateName As String) As TemplateObj
' ------------------------------------
' Function: CreateTemplate_Umfrage
'
' Creates and returns a TemplateObj meant for Template "(A1) Umfrage"
' We have predefined which rows and columns are to be mapped to the template
' and are identified by the ZNr and SNr of the given template.
'
' Arguments:
'   templateName: the name of the desired Template type
'
' Returns:
'   TemplateObj: with defined structure and dummy values
'
' See Also:
'   [CreateTemplate][CreateTemplate], [CreateTemplate_ST_GuV][CreateTemplate_ST_GuV], [CreateTemplate_ST_Zinsergebnis][CreateTemplate_ST_Zinsergebnis]

    ' Init
    Dim template As TemplateObj
    Set template = New TemplateObj
    Dim RowOffDict As Object
    Set RowOffDict = CreateObject("Scripting.Dictionary")
    Dim RowDict As Object
    Set RowDict = CreateObject("Scripting.Dictionary")
    'For now, the columns we know we are concerned about are 7-31 inclusive, no break
    
    ' How to handle bitte auswahlen? Is that our purview?
    '    E.g. Row 53 oonwards as well
    'Note that We might need too fill ZNr 49-52 TBD
    
    'Note that row number offset changesss at Znr 28 and 53
    
    'Rows skipped 1,6,7,8,11,12,14,23,27
    
    ' Declaration of OFFSETS
    template.templateName = templateName
    RowOffDict.Add 2, 11 - 2   ' since we do not start with 1 anyway
    RowOffDict.Add 29, 41 - 29 ' since we do not start with 28 anyway
    RowOffDict.Add 53, 69 - 53
    Set template.RowOffDict = RowOffDict
    template.ColOffset = 8 - 1
    
    ' -------------------------------------------------------------------------
    'creating the TemplateRow for 1-27
    'create EXCLUDED rowss
    Dim excludedRow() As Variant
    excludedRow = Array(1, 6, 7, 8, 11, 12, 14, 23, 27)
    
    Dim i As Integer, j As Integer
    'Dim ROW As Object
    Dim row As TemplateRow
    Dim colDict As Object
    For i = 2 To 3

        Set row = New TemplateRow
        
        If Not IsValueInArray(i, excludedRow) Then
            row.template = templateName
            row.RowNumber = i
            row.RowName = "Filler Text" 'Look up way to pull from template? multiple merged rows?
            'row.excludedCol = excludedCol 'this part dooes not have any excloded cols
     
            Set colDict = CreateObject("Scripting.Dictionary")
            ' Populate the the dictionary of the columns values, where the keys are the SNr
            For j = 7 To 31
                ' skipping columns
                'If Not IsValueInArray(j, excludedCol) Then
                '    colDict.Add j, "(" & i & ", " & j & ")"
                'End If
                'colDict.Add j, "(" & i & ", " & j & ")" ' INDEX CHECK
                colDict.Add j, i + j 'integer values
                
            Next j
            
            'Saving dict of col values to row object and then to RowDict of TemplateObj
            Set row.ColValuesDict = colDict
            RowDict.Add i, row
            
        End If
    Next i
              
     
    'Adding the RowDict to the template object
    Set template.RowValuesDict = RowDict
    
    ' Return the templateObj
    Set CreateTemplate_Umfrage = template
    Debug.Print "Created TemplateObj: " & templateName

End Function


Function CreateTemplate_ST_GuV(ByVal templateName As String) As TemplateObj
' Creates and returns a TemplateObj meant for Template "(A2) ST-GuV-Kapital"
' We have predefined which rows and columns are to be mapped to the template
' and are identified by the ZNr and SNr of the given template.
'
' Arguments:
'   templateName: the name of the desired Template type
'
' Returns:
'   TemplateObj: with defined structure and dummy values
'
' See Also:
'   [CreateTemplate][CreateTemplate], [CreateTemplate_Umfrage][CreateTemplate_Umfrage], [CreateTemplate_ST_Zinsergebnis][CreateTemplate_ST_Zinsergebnis]

    ' Init
    Dim template As TemplateObj
    Set template = New TemplateObj
    Dim RowOffDict As Object
    Set RowOffDict = CreateObject("Scripting.Dictionary")
    Dim RowDict As Object
    Set RowDict = CreateObject("Scripting.Dictionary")

    ' Declaration of OFFSETS
    template.templateName = templateName
    RowOffDict.Add 2, 11 - 2   '

    Set template.RowOffDict = RowOffDict
    template.ColOffset = 13 - 5
    
    ' -------------------------------------------------------------------------
    'creating the TemplateRow for 2-3
    'create EXCLUDED rowss
    Dim excludedRow() As Variant
    excludedRow = Array(1)
    
    Dim i As Integer, j As Integer
    'Dim ROW As Object
    Dim row As TemplateRow
    Dim colDict As Object
    For i = 2 To 3 ' 27

        Set row = New TemplateRow
        
        If Not IsValueInArray(i, excludedRow) Then
            row.template = templateName
            row.RowNumber = i
            row.RowName = "Filler Text" 'Look up way to pull from template? multiple merged rows?
            'row.excludedCol = excludedCol 'this part dooes not have any excloded cols
     
            Set colDict = CreateObject("Scripting.Dictionary")
            ' Populate the the dictionary of the columns values, where the keys are the SNr
            For j = 5 To 7
                colDict.Add j, "(" & i & ", " & j & ")" ' INDEX CHECK
                'colDict.Add j, i + j 'working values
                
            Next j
            
            'Saving dict of col values to row object and then to RowDict of TemplateObj
            Set row.ColValuesDict = colDict
            RowDict.Add i, row
            
        End If
    Next i
            

    'Adding the RowDict to the template object
    
    Set template.RowValuesDict = RowDict
    
    ' Return the templateObj
    Set CreateTemplate_ST_GuV = template
    'Debug.Print "Created TemplateObj: " & templateName

End Function


Function CreateTemplate_ST_Zinsergebnis(ByVal templateName As String) As TemplateObj
' ------------------------------------
' Function: CreateTemplate_ST_Zinsergebnis
'
' Creates and returns a TemplateObj meant for Template "(A2) ST-Zinsergebnis"
' We have predefined which rows and columns are to be mapped to the template
' and are identified by the ZNr and SNr of the given template.
'
' Arguments:
'   templateName: the name of the desired Template type
'
' Returns:
'   CreateTemplate_ST_Zinsergebnis: with defined structure and dummy values
'
' See Also:
'   <CreateTemplate>, <CreateTemplate_Umfrage>, <CreateTemplate_ST_GuV>
' ------------------------------------

    ' Init
    Dim template As TemplateObj
    Set template = New TemplateObj
    Dim RowOffDict As Object
    Set RowOffDict = CreateObject("Scripting.Dictionary")
    Dim RowDict As Object
    Set RowDict = CreateObject("Scripting.Dictionary")
    'For now, the columns we know we are concerned about are 12-35 inclusive, no break
    

    ' Declaration of OFFSETS
    template.templateName = templateName
    RowOffDict.Add 1, 11 - 1   '
    RowOffDict.Add 2, 12 - 2   ' since  row 1 is unique
    Set template.RowOffDict = RowOffDict
    template.ColOffset = 7 - 1
    
    ' -------------------------------------------------------------------------
    'creating the TemplateRow for 1-16
    'create EXCLUDED rowss
    Dim excludedRow() As Variant
    excludedRow = Array(7, 9, 15) ' fully skipped rows
    
    'Special treament rows
    Dim uniqueRow() As Variant
    uniqueRow = Array(1, 10, 16)
    
    'create EXCLUDED columns
    Dim excludedCol() As Variant
    Dim excludedColU() As Variant
    excludedCol = Array(3, 4, 11)
    
    Dim i As Integer, j As Integer, k As Integer
    'Dim ROW As Object
    Dim row As TemplateRow
    Dim colDict As Object
    For i = 1 To 16 ' 16

        Set row = New TemplateRow
        
        If Not IsValueInArray(i, excludedRow) Then
            row.template = templateName
            row.RowNumber = i
            row.RowName = "Filler Text" 'Look up way to pull from template? multiple merged rows?
            
            Set colDict = CreateObject("Scripting.Dictionary")
     
            'Now check if it is unique row arrangement
            If Not IsValueInArray(i, uniqueRow) Then
                row.excludedCol = excludedCol 'put here as qunie row have diff excluded col
            
                
                ' Populate the the dictionary of the columns values, where the keys are the SNr
                For j = 12 To 35
                    ' skipping coolumns
                    If Not IsValueInArray(j, excludedCol) Then
                         colDict.Add j, "(" & i & ", " & j & ")" ' INDEX check
                         'colDict.Add j, i + j 'working values
                    End If
                    
                
                Next j
            

            Else ' It is an unique row
            
                ' Checking if the row has special column considerations
                ' Debug.Print "Special Row " & i
                If i = 1 Then
                    excludedColU = Array(1, 3, 4, 5, 7, 9, 11, _
                                        12, 14, 15, 16, 18, 20, 22, 23, _
                                        24, 26, 28, 30, 31, 32, 34, _
                                        36, 37, 38) 'every 2...
                    row.excludedCol = excludedColU
                
                ElseIf i = 10 Then
                    ReDim excludedColU(LBound(excludedColU) To 31)
                    excludedColU = Array(1, 3, 4, 5, 7, 9, 11, _
                                        12, 13, 14, 15, 16, 17, 18, _
                                        20, 21, 22, 23, 24, 25, 26, _
                                        28, 29, 30, 31, 32, 33, 34, _
                                        36, 37, 38)
                                        
                    row.excludedCol = excludedColU
                ElseIf i = 16 Then
                    Dim includeCol() As Variant
                    Dim preArray() As Variant
                    Dim index As Integer
                    includeCol() = Array(7, 8, 11, 36, 37, 38)

                    
                    ReDim excludedColU(1 To 32)
                    index = 1
                    For k = 1 To 38
                        If Not IsValueInArray(k, includeCol) Then
                            excludedColU(index) = k
                            index = index + 1
                        End If
                        
                    Next k
                    
                    ' Use the Filter function to exclude specific elements from the original array
                    'Why dooess it not work?
                    'excludedColU = Filter(preArray, Function(x) Not IsValueInArray(x, includeCol), False)
                    
                   
                    row.excludedCol = excludedColU
                End If
                
                ' Finally output to the row
                For j = 12 To 38 'This is increased to account for row 16
                    ' skipping coolumns
                    If Not IsValueInArray(j, excludedColU) Then
                         colDict.Add j, "(" & i & ", " & j & ")"
                         'colDict.Add j, i + j 'working values
                    End If
                    
                
                Next j
            
            End If
            
            'Saving dict of col values to row object and then to RowDict of TemplateObj
            Set row.ColValuesDict = colDict
            RowDict.Add i, row
        End If
    Next i
         
            
            
    'Adding the RowDict to the template object
    
    Set template.RowValuesDict = RowDict
    
    ' Return the templateObj
    Set CreateTemplate_ST_Zinsergebnis = template
    Debug.Print "Created TemplateObj: " & templateName

End Function



Sub WriteTemplateOut(ByVal dWB As Workbook, dWBName As String, ByVal dWSName As String, template As TemplateObj)
' Given workbook path, workbook name, and template TemplateObj
' write the values out to the correct Template Sheet in correct worksheet
' rowOffset is then a dictionary of the (key) starting row and the (value) offset number
'
' Arguments:
'   dWB: Destination Workbook object
'   dWBName: Destination Workbook name
'   dWSName: Destination Worksheet name
'   template: TemplateObj with values stored to be written into the Destination Workbook
'
' Returns:
'   Nothing: 
'
' See Also:
'   [AreItemsContinuous][], [writeToWorkbookRow][writeToWorkbookRow], [writeToWorkbookCell][writeToWorkbookCell]


    ' init var
    Dim ColOffset As Integer
    Dim RowOffset As Integer
    Dim RowOffDict As Object
    Dim RowOffKeys As Variant
    Dim RowDict As Object
    
    ' in the 1st loop
    Dim rowObj As TemplateRow
    Dim colDict As Object
    
    ' in the second loop
    Dim col As Variant
    
    If Not template.templateName = dWSName Then
        MsgBox "Template mismatch. Trying to write " & template.templateName & " into " & dWSName
        ' raise exception?
    Else
    
        ' obtain dictionary of the rows
        Set RowDict = template.RowValuesDict
        Set RowOffDict = template.RowOffDict
        'RowOffKeys = template.RowOffDict.Keys
        ColOffset = template.ColOffset ' asssumption is that there are no jumps
        
        

        
        Dim row As Variant
        For Each row In RowDict
            'Debug.Print "  Row Key: " & row
            
            ' Obtain the rowObj and the dictionary of column values
            Set rowObj = RowDict(row)
            Set colDict = rowObj.ColValuesDict
            
            
            ' correct for row offset and if jump, and col offset
            ' The assumption is that once it reaches the next offset interval
            ' it never goes backwards, only changes at the next jump forward
            If RowOffDict.Exists(row) Then
                RowOffset = RowOffDict(row)
            End If
            'Debug.Print "  Offset used: " & RowOffset
            
            
            'check if columns are continous
            'if yes, write row out
            
            Dim colKeys() As Variant
            Dim colVals As Variant
            colKeys = colDict.Keys
            colVals = WorksheetFunction.Transpose(colDict.Items)
            
            If AreItemsContinuous(colKeys) Then
                ' writ out by ROW
                 
                writeToWorkbookRow dWB, dWSName, _
                                  row + RowOffset, _
                                  colKeys(0) + ColOffset, _
                                  colVals

            Else
            
                'if not, write by each column and row identifier
                For Each col In colDict
                    ' print it out
                    'Debug.Print "    Column Key: " & col & ", Inner Value: " & colDict(col)
                
                    ' put into template
                    writeToWorkbookCell dWB, dWSName, _
                                    row + RowOffset, col + ColOffset, _
                                    colDict(col)
                Next col
                
            End If
            
        Next row

    End If


    Debug.Print "Wrote to Template " & dWSName; " of " & dWBName
End Sub

