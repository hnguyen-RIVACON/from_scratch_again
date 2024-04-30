VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' cls TemplateRow class MODULE
'Generic Row object for filling Given templates 2024.04.15 LSI Stress testing

' Row object with these properties
' - template identifier
' - row number (Znr) as defined in the template worksheet
' - row name
' - array of excluded columns
' - data dict of columns values
'                - key: col#
'                - value: column value



' Class: TemplateRow
'
' Designed to store a dictionary whose (key, value) pairs are the (column numbers, and values to be written out)
'
' See Also:
'   <TemplateObj>




' Property: m_template
'   Name of the template
Private m_template As String 'FOR WHICH TEMPLATE THIS ROW IS SUPPOSED TO BE FOR

' Property: m_rowNumber
'   Row number of worksheet to be filled
Private m_rowNumber As Integer 'WHICH ROW IT IS SUPPOSED TO FILL

' Property: m_rowName
'   Name of row of worksheet to be filled
Private m_rowName As String 'NAME OF THE ROW

' Property: array_exCol
'   Array of columns to be excluded in the structure
Private array_exCol() As Variant 'COLUMNS TO EXCLUDE IN THE STRUCTURE

' Property: m_template
'   'Dictionary with all the values to be stored, where the keys are the column numbers
Private ColDictionary As Object 'DICTIONARY with all the values to be stored, where the keys are the column numbers


' Properties
Public Property Get template() As String
    template = m_template
End Property

Public Property Let template(value As String)
    m_template = value
End Property

Property Get RowNumber() As Integer
    RowNumber = m_rowNumber
End Property

Property Let RowNumber(value As Integer)
    m_rowNumber = value
End Property

Property Get RowName() As String
    RowName = m_rowName
End Property

Property Let RowName(value As String)
    m_rowName = value
End Property

Property Get excludedCol() As Variant
    excludedCol = array_exCol()
End Property

Property Let excludedCol(value As Variant)
    array_exCol() = value
End Property

Property Get ColValuesDict() As Object
    Set ColValuesDict = ColDictionary
End Property

Property Set ColValuesDict(value As Object)
    Set ColDictionary = value
End Property

' methods
Public Sub DisplayDictionaryKeys()
    If Not ColDictionary Is Nothing Then
        Dim key As Variant
        For Each key In ColDictionary.Keys
            Debug.Print key
        Next key
    Else
        Debug.Print "Dictionary is not set."
        
    End If
    
End Sub

'' Methods
'Public Sub setName(name As String)
'    m_rowName = name
'End Sub
'
'Public Sub setRow(number As Integer)
'    m_rowNumber = number
'End Sub
'
'Public Sub setTemplate(TemplateName As String)
'    m_template = TemplateName
'End Sub
'
'Public Sub setExCol(exCol As Variant)
'    array_exCol = exCol
'End Sub
'
'Public Sub setColDict(colDict As Object)
'    ColDictionary = colDict
'End Sub

'keyword: title
