VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' cls TemplateObj class MODULE
' Template object with these properties
' - template identifier
' - dictionary containg the TemplateRow objects
'                - key: row#
'                - value: TemplateRow object
' - The starting offset for the row identifier TemplateObj
' - The starting offset for the column identifier TemplateObj

' Class: TemplateObj
'
' Designed to store TemplateRow objects
'
' See Also:
'   <TemplateRow>

'-------------------------
' Variables
'-------------------------

' Property: m_template
'   Name of the template
Private m_template As String



' Property: RowDictionary
'   Dictionary with all the TemplateRow objects
'   where the keys are the Row numbers of the given
'   Template
Private RowDictionary As Object

' Property: RowOffset
'   Dictionary with ZNr start (key) and offset from Excel row
Private RowOffset As Object

' Property: ColOffset
'   Offset for first column (SNr)
Private m_ColOffset As Integer



'-------------------------
' Properties
Public Property Get templateName() As String
    templateName = m_template
End Property

Public Property Let templateName(value As String)
    m_template = value
End Property

Public Property Get ColOffset() As Integer
    ColOffset = m_ColOffset
End Property

Public Property Let ColOffset(value As Integer)
    m_ColOffset = value
End Property

Property Get RowValuesDict() As Object
    Set RowValuesDict = RowDictionary
End Property

Property Set RowValuesDict(value As Object)
    Set RowDictionary = value
End Property

Property Get RowOffDict() As Object
    Set RowOffDict = RowOffset
End Property

Property Set RowOffDict(value As Object)
    Set RowOffset = value
End Property

'-------------------------
' Methods

'-------------------------
' Sub: DisplayDictionaryKeys
'   Displays the keys of the given dictionary
'
' Returns:
'   Nothing
'
Public Sub DisplayDictionaryKeys()
    If Not RowDictionary Is Nothing Then
        Dim key As Variant
        For Each key In RowDictionary.Keys
            Debug.Print key
        Next key
    Else
        Debug.Print "Dictionary is not set."
        
    End If
    
End Sub

'keyword: title
