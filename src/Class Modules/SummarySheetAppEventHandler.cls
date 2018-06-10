VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SummarySheetAppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'handles events of application
Private WithEvents AppEvent As Application
Attribute AppEvent.VB_VarHelpID = -1

'handles changes in the summary sheet (update custom properties if edited manually)
Private Sub AppEvent_SheetChange(ByVal Sh As Object, ByVal Target As Range)
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    'is change in the summary sheet?
    If Sh.Name = getSummarySheetName() Then
    'is there already a table? If not -> just exit sub
    If Sh.ListObjects.Count = 0 Then Exit Sub
    
            Dim cl As Range
            If Intersect(Target, Sh.Range(Sh.ListObjects(1).Range.Address)) Is Nothing Then Exit Sub
            For Each cl In Intersect(Target, Sh.Range(Sh.ListObjects(1).Range.Address))
            On Error Resume Next
            'first row are just headers -> skip
            If cl.Row = 1 Then GoTo Skip
            'first column is just hyperlink to other worksheets -> skip
            If cl.Column = 1 Then GoTo Skip

            If Sh.Cells(cl.Row, 1).Value <> "" Then
            setProperty Sh.Parent.Worksheets(Sh.Cells(cl.Row, 1).Value), Sh.Cells(1, cl.Column).Value, cl.Value
            End If
Skip:
            On Error GoTo 0
            Next cl
    End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

'constructor:
Public Sub Class_Initialize()
    'bind events from "Application" to this class var "AppEvent"
    Set AppEvent = Application
End Sub

' New WorkBook added in Application: with it also the first WorkSheet was created
Private Sub AppEvent_NewWorkbook(ByVal wb As Workbook)
    'set a custom property for the first worksheet
    setProperty wb.Sheets(1), getWorksheetCreatedDatePropName(), Now
End Sub

' new WorkSheet added
Private Sub AppEvent_WorkbookNewSheet(ByVal wb As Workbook, ByVal Sh As Object)
    'set a custom property for the new worksheet
    setProperty Sh, getWorksheetCreatedDatePropName(), Now
     
    If worksheetExists(wb, getSummarySheetName()) Then
        Call generateSummaryWorksheet
    End If
    
End Sub



