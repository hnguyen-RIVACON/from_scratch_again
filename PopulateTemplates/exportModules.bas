Attribute VB_Name = "exportModules"
Option Explicit

' ------------------------------------
' Sub: prcExort
'
' Module to export all .bas and .cls used for VBA code of a given Workbook Object into a
' destination folder. The folder is hardcoded and can be changed as necessary
'
' Returns:
'   Nothing
'
' See Also:
'   <ImportFilesToRepo>
'------------------------------------
Public Sub prcExort()
    Dim objVBComponent As Object
    Dim objWorkbook As Workbook
    Dim strType As String
    Set objWorkbook = ActiveWorkbook
    
    For Each objVBComponent In objWorkbook.VBProject.VBComponents
        With objVBComponent
            Select Case .Type
            Case 1 'vbext_ct_StdModule
                strType = ".bas"
            Case 2 'vbext_ct_ClassModule
                strType = ".cls"
            Case 3 'vbext_ct_MSForm
                strType = ".frm"
            Case 100 'vbext_ct_Document
                strType = ".clsS"
            End Select
            If strType = ".cls" Or strType = ".bas" Then
                objWorkbook.VBProject.VBComponents(.name).Export _
                "C:\Users\DrHansNguyen\Documents\CLIENT_PROJECTS\Kunden\Ford_Bank\Projekte\2024_IRRBB_STRESSTEST\CODE\ford_lsi_irrbb\PopulateTemplates\" & .name & strType
                '"C:\Users\Anwender\Documents\Projekte\Fordbank\IRRBB_LSI\ford_lsi_irrbb\VBA\" & .name & strType
            End If
            
        End With
        Next
End Sub

