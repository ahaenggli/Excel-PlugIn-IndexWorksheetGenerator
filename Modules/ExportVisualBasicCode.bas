Attribute VB_Name = "ExportVisualBasicCode"
Public Sub ExportVisualBasicCode(Optional WB As Workbook, Optional directory As String)
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String

    Dim extension As String
        
    If WB Is Nothing Then Set WB = ThisWorkbook
           
    If directory = "" Then directory = WB.path & "\"
    count = 0
    
    If Dir(directory, vbDirectory) = "" Then
          MkDir directory
    End If
    
    Dim exportpath As String
    For Each VBComponent In WB.VBProject.VBComponents
        exportpath = directory
        
        Select Case VBComponent.Type
            Case Document
                extension = ".cls"
                exportpath = directory
            Case ClassModule
                extension = ".cls"
                exportpath = directory & "\Class Modules\"
            Case Form
                extension = ".frm"
                exportpath = directory & "\Forms\"
            Case Module
                extension = ".bas"
                exportpath = directory & "\Modules\"
            Case Else
                extension = ".txt"
                exportpath = directory
        End Select
            
    If Dir(exportpath, vbDirectory) = "" Then
          MkDir exportpath
    End If
    
                
        On Error Resume Next
        Err.Clear
        
        path = exportpath & "\" & VBComponent.Name & extension
            
        Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    
End Sub

Public Sub ExportThis()
    ExportVisualBasicCode
End Sub


