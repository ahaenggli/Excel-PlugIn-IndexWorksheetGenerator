Attribute VB_Name = "GlobalFunction"
Option Explicit

'check whether Excel-GUI is german or not
Public Function isGermanGUI() As Boolean
Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    Case msoLanguageIDSwissGerman, _
            msoLanguageIDGermanLiechtenstein, _
            msoLanguageIDGerman, _
            msoLanguageIDGermanAustria, _
            msoLanguageIDGermanLuxembourg
        isGermanGUI = True
    Case Else
        isGermanGUI = False
    End Select
End Function

'check whether a value is in an array of values or not
Public Function inArray(Value As Variant, Arr As Variant) As Boolean
    Dim tmpValue As Variant
On Error GoTo ErrorHandler: 'array is empty
    For Each tmpValue In Arr
        If tmpValue = Value Then
            inArray = True
            Exit Function
        End If
    Next tmpValue
Exit Function
ErrorHandler:
On Error GoTo 0
    inArray = False
End Function

' replaces empty string with alternative string
Public Function isNull(val1 As String, val2 As String) As String
    If val1 <> "" Then
        isNull = val1
    Else
        isNull = val2
    End If
End Function

'Does the sheet exists in specific workbook?
Public Function worksheetExists(wb As Workbook, sheetToFind As String) As Boolean
    worksheetExists = False
    Dim Sheet As Worksheet
    
    For Each Sheet In wb.Worksheets
        If sheetToFind = Sheet.Name Then
            worksheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

