Attribute VB_Name = "Copy_subs"
Option Explicit
' s as source
' d as destination
' wb as workbook
' ws as worksheet



Function IsValueInArray(searchValue As Variant, arr As Variant) As Boolean
' This is a docstring.
'
' Arguments:
'     searchValue: value we want searched for
'     arr: the array we want to be searched
'
' Examples:
'     >>> IsValueInArray(1, Array(1,2,3))
'     True

' This comment is not part of the docs.

    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = searchValue Then
            IsValueInArray = True
            Exit Function
        End If
    Next i
    IsValueInArray = False
End Function




Sub writeToWorkbookCell(ByVal dWB As Workbook, dWSName As String, _
                        ByVal dRowNumber As Integer, _
                        ByVal dColumnStart As Integer, _
                        value2write As String)
' VBA Sub to output to a cell.
' Given a Workbook dWB and Worksheet name dWSName,
' write out the value of value2write into cell
' position (dRowNumber, dColumnStart)
'
' Arguments:
'   dWB: Destination Workbook
'   dWSName: name of the Destination Worksheet
'   dRowNumber: Row number of Destination Cell
'   dColumnStart: Column number of Destination Cell
'   value2write: value we want to be placed into Destination Cell
'



    Dim dWS As Worksheet
    Dim dRange As Range
    

    
    ' Open the destination workbook
    'Set dWB = Workbooks.Open(dWBPath & "\" & dWBName) 'sshould be open already
     
    ' Set the destination worksheet
    Set dWS = dWB.Worksheets(dWSName)
    

    ' Set the destination range
    'Set dRange = dWS.Range(dWS.Cells(dRowNumber, dColumnStart))
    Set dRange = dWS.Cells(dRowNumber, dColumnStart)

    'write to the cell
    dRange.value = value2write

    '' Close the workbooks without saving changes (change this according to your requirement)
    'sourceWorkbook.Close False
    'destinationWorkbook.Close True
    
    'Turn THIS ON WHEN YOU WANT TO SAVE or add IF statement and make it a paremeter
    'dWB.Save
    
    ' Clean up
    Set dRange = Nothing
    Set dWS = Nothing
    Set dWB = Nothing
End Sub


Sub writeToWorkbookRow(ByVal dWB As Workbook, dWSName As String, _
                        ByVal dRowNumber As Integer, _
                        ByVal dColumnStart As Integer, _
                        ByVal Values As Variant)
' ------------------------------------
' Sub: writeToWorkbookRow
'
' VBA Sub to output to a partial row.
' Given a Workbook dWB and Worksheet name dWSName,
' write out the array of valuess in Values starting at the
' position (dRowNumber, dColumnStart)
'
' Arguments:
'   dWB: Destination Workbook
'   dWSName: name of the Destination Worksheet
'   dRowNumber: Row number of Destination Cell
'   dColumnStart: Column number of Destination Cell
'   Values: value we want to be placed into Destination Cell
'
' Returns:
'   Nothing:




    Dim dWS As Worksheet
    Dim dRange As Range
 
    ' Open the destination workbook
    'Set dWB = Workbooks.Open(dWBPath & "\" & dWBName) 'sshould be open already
     
    ' Set the destination worksheet
    Set dWS = dWB.Worksheets(dWSName)
 
    ' Set ouput range
    Set dRange = dWS.Cells(dRowNumber, dColumnStart)
            
    ' Resize the range to match the size of the values array
    Set dRange = dRange.Resize(1, UBound(Values) - LBound(Values) + 1)
    
    ' Write the values to the range
    dRange.value = Application.Transpose(Values)


    '' Close the workbooks without saving changes (change this according to your requirement)
    'sourceWorkbook.Close False
    'destinationWorkbook.Close True
    
    'Turn THIS ON WHEN YOU WANT TO SAVE or add IF statement and make it a paremeter
    'dWB.Save
    
    ' Clean up
    Set dRange = Nothing
    Set dWS = Nothing
    Set dWB = Nothing
End Sub


Sub SaveUpdatedWorkbookAsCopy(ByVal origWBPath As String, ByVal origWBName As String)
' ------------------------------------
' VBA Sub to save a copy
' Given the original Workbook to be copied path as origWBPath with Workbook name origWBName
' Save the changes made into a New Workbook that is a copy with an updated filename
'
' Arguments:
'   origWBPath: Path of Workbook to be copied
'   origWBName: name of the Workbook to be copied
'
' Returns:
'   Nothing:
'
'------------------------------------

    Dim originalWorkbook As Workbook
    Dim originalFilePath As String
    Dim newFileName As String
    Dim newFilePath As String
    
    ' Open the workbook
    Set originalWorkbook = Workbooks.Open(origWBPath & "\" & origWBName)

    ' Get the path and filename of the original workbook
    originalFilePath = originalWorkbook.FullName
    
    ' Modify the filename to create a new name for the copy
    'newFileName
    newFileName = Left(origWBName, Len(origWBName) - 5) & "_autofilled.xlsx" ' Assuming the original workbook is in .xlsx format
    
    
    ' Construct the new file path by replacing the filename in the original path
    newFilePath = Left(originalFilePath, InStrRev(originalFilePath, "\")) & newFileName
    
    ' Check if the file already exists, and delete it if it does
    If Dir(newFilePath) <> "" Then
        On Error Resume Next
        Kill newFilePath
        On Error GoTo 0
    End If
    
    ' Save a copy of the workbook with the updated name
    originalWorkbook.SaveCopyAs newFilePath
    
    ' Close the original workbook without saving
    'originalWorkbook.Close False

End Sub

