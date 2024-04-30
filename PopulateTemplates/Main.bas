Attribute VB_Name = "Main"
Option Explicit

' ------------------------------------
' Sub: Main
'
' Main sub to populate template Worksheets in a known Workbook
' In this case we asssume it will be in the same folder as the IRRBB calculator
' for this test purpose we expect workbook of name LSI_ST2024_1208175_V1.xlsx
'
' Returns:
'   Nothing
'
' See Also:
'   <CreateTemplate>, <WriteTemplateOut>, <SaveUpdatedWorkbookAsCopy>
'------------------------------------
Sub Main()


    Dim dWBPath As String
    Dim dWBName As String
    Dim dWB As Workbook
    
    dWBName = "LSI_ST2024_1208175_V1.xlsx"
    dWBPath = ThisWorkbook.Path
    
    
    Dim templateNames() As Variant
    'templateNames = Array("(A1) Umfrage")
    templateNames = Array("(A1) Umfrage", "(A2) ST-GuV-Kapital", "(A2) ST-Zinsergebnis")
    'All template names
    'templateNames = Array("(A1) Umfrage", "(A2) ST-Befreiung RK", "(A2) ST-GuV-Kapital", _
    '                     "(A2) ST-Zinsergebnis", "(A2) ST-Adressrisiko", "(A2) ST-Marktrisiko")
    
    
   
    ' Creating TemplateObjs to be filled
    Dim templatesDict As Object
    Set templatesDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    Dim name As Variant
    For i = LBound(templateNames) To UBound(templateNames)
        name = templateNames(i)
        Dim template As TemplateObj
        Set template = New TemplateObj
        Set template = CreateTemplate(name)

        templatesDict.Add name, template

    Next i
                          
    
    
    ' Write To templates
    
    ' Open workbook once
    Set dWB = Workbooks.Open(dWBPath & "\" & dWBName)
    Application.ScreenUpdating = False ' for speed performance
    For Each name In templateNames
        WriteTemplateOut dWB, dWBName, name, templatesDict(name)
    Next name
    Application.ScreenUpdating = True
    
    '  and save workbook as a COPY
    Debug.Print "Saving a Copy As " & Left(dWBName, Len(dWBName) - 5) & "_autofilled.xlsx"
    SaveUpdatedWorkbookAsCopy dWBPath, dWBName
    
    Debug.Print "Got to end of Main"



End Sub


