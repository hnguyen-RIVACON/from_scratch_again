Attribute VB_Name = "importModules1"
Option Explicit

' ------------------------------------
' Sub: ImportFilesToRepo
'
' Module to import all .bas and .cls  files used for VBA code of a given Workbook Object
' From a source folder into the given Workbook. The source folder is hardcoded and can be changed as necessary
'
' Returns:
'   Nothing
'
' See Also:
'   <prcExort>
'------------------------------------
Sub ImportFilesToRepo()

    ' Tool for importing VBA files from a given folder destination
    ' Make sure to go to the Toolbar Menu -> Tools -> References -> Select "Microsoft Visual Basic For Applications Extensibility 5.3"
    ' Without this, you will not be able to access the VBA Project of the workbook

    Dim pathName As String: pathName = "C:\Users\DrHansNguyen\Documents\CLIENT_PROJECTS\Kunden\Ford_Bank\Projekte\2024_IRRBB_STRESSTEST\CODE\ford_lsi_irrbb\PopulateTemplates\" '"C:\Users\Anwender\Documents\VBA\"

    'Dir is a function that allows you to iterate through files in a directory
    Dim filePath As Variant: filePath = Dir(pathName)
    Dim vbModule As Object
    Dim conditionMet As Boolean

    Do While Len(filePath) > 0
        conditionMet = False
        'Accounting for modules, classes, and forms
        If Right(filePath, 3) = "bas" Or Right(filePath, 3) = "cls" Or Right(filePath, 3) = "frm" Then

            'Remove the edge case of this file
            If filePath <> "gitConnector.bas" Then

                'Need to remove the existing module
                For Each vbModule In ThisWorkbook.VBProject.VBComponents

                    'If the name of the module matches the name of the file (without its extension)
                    If vbModule.name = Left(filePath, Len(filePath) - 4) Then

                        'Delete the module
                        ThisWorkbook.VBProject.VBComponents.Remove vbModule

                        'Import the new module from the path
                        ActiveWorkbook.VBProject.VBComponents.Import (pathName & filePath)
                        conditionMet = True
                        Exit For
                    End If
                    
                Next
                If conditionMet = False Then
                    Debug.Print filePath
                    ActiveWorkbook.VBProject.VBComponents.Import (pathName & filePath)
                End If
            End If
            
        End If
        
        'This is how Dir iterates to the next file
        filePath = Dir
    Loop
End Sub

