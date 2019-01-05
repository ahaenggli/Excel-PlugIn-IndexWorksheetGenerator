VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TocSheetAppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'handles events of application
Private WithEvents AppEvent As Application
Attribute AppEvent.VB_VarHelpID = -1

'handles changes in the Toc sheet (update custom properties if edited manually)
Private Sub AppEvent_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    'not a worksheet?
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    'not Toc page?
    If Sh.Name <> getTocSheetName() Then Exit Sub
    'is there already a table? If not -> just exit sub
    If Sh.ListObjects.count = 0 Then Exit Sub
    
    If isF5 Then Exit Sub
    
    Dim cl As Range
    
    Dim arrIdxCols As Variant
    Dim arrCusProp As Variant
    Dim newCusProp As String
    
    arrIdxCols = getTocColumns()
    arrCusProp = getTocCustomProperties()
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
     
    'not in table-area (updated cell) and (number of table-columns is equal Toc-arry)
    If Intersect(Target, Sh.Range(Sh.ListObjects(1).Range.Address)) Is Nothing And Sh.ListObjects(1).ListColumns.count = UBound(arrIdxCols) + 1 Then Exit Sub
            
    'delete Toc props in available sheets
    Dim rw As Range
    Dim vl As Variant
    For Each rw In Sh.ListObjects(1).Range.Rows
        If worksheetExists(ActiveWorkbook, rw.Columns(1).Value) Then
            For Each vl In arrIdxCols
              If CStr(vl) <> getWorksheetCreatedDatePropName() Then Call setProperty(ActiveWorkbook.Worksheets(rw.Columns(1).Value), CStr(vl), "")
            Next vl
        End If
    Next rw
     
    Dim x As Integer
    x = 0
    'resize
    ReDim arrIdxCols(Sh.ListObjects(1).HeaderRowRange.Cells.count - 1)
    'Loop through each cell in Range and store value in Array
    For Each rw In Sh.ListObjects(1).HeaderRowRange.Cells
      arrIdxCols(x) = rw.Value
      x = x + 1
    Next rw
                  
            For Each cl In Sh.ListObjects(1).Range
                On Error Resume Next
                    'first row are headers -> skip
                    If cl.Row = 1 Then GoTo Skip
                    'first column is just hyperlink to other worksheets -> skip
                    If cl.Column = 1 Then GoTo Skip
        
                    If Sh.Cells(cl.Row, 1).Value <> "" Then
                        setProperty Sh.Parent.Worksheets(Sh.Cells(cl.Row, 1).Value), Sh.Cells(1, cl.Column).Value, cl.Value
                    End If
    
Skip:
                On Error GoTo 0
            Next cl
            

For Each vl In arrCusProp
    If inArray(vl, arrIdxCols) Then
        newCusProp = newCusProp + CStr(vl) + ";"
    Else
        If Not inArray(vl, arrIdxCols) And inArray(vl, getTocColumns()) Then
            x = getTocInArray(vl, getTocColumns())
            If x > -1 And x <= UBound(arrIdxCols) Then newCusProp = newCusProp + arrIdxCols(x) + ";"
        Else
            newCusProp = newCusProp + CStr(vl) + ";"
        End If
    End If
newCusProp = Replace(newCusProp, ";;", ";")
Next vl

If Right(newCusProp, 1) = ";" Then newCusProp = Left(newCusProp, Len(newCusProp) - 1)

setProperty Sh, "TocColumns", Join(arrIdxCols, ";")
setProperty Sh, "TocCustomProperties", newCusProp


Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

'constructor:
Public Sub Class_Initialize()
    'bind events from "Application" to this class var "AppEvent"
    Set AppEvent = Application
End Sub

' New WorkBook added in Application: with it also the first WorkSheet was created
Private Sub AppEvent_NewWorkbook(ByVal WB As Workbook)
    Dim cPrpNm As String
    cPrpNm = getWorksheetCreatedDatePropName()
    
   'set a custom property for the first worksheet
   If cPrpNm <> "" Then
     setProperty WB.Sheets(1), cPrpNm, Now
   End If
   
   setProperty WB.Sheets(1), "isToc", "0"
End Sub

' new WorkSheet added
Private Sub AppEvent_WorkbookNewSheet(ByVal WB As Workbook, ByVal Sh As Object)
If Not TypeOf Sh Is Worksheet Then Exit Sub
Dim cPrpNm As String
cPrpNm = getWorksheetCreatedDatePropName()

    'set a custom property for the new worksheet
   If cPrpNm <> "" Then
    setProperty Sh, cPrpNm, Now
   End If
   
   setProperty Sh, "isToc", "0"
   
    If worksheetExists(WB, getTocSheetName()) Then
        Call generateTocWorksheet
    End If
    
End Sub

'refresh table of contents sheet if activated
Private Sub AppEvent_SheetActivate(ByVal Sh As Object)
    'not a worksheet?
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    'not a workbook?
    If Not TypeOf Sh.Parent Is Workbook Then Exit Sub
    
    'call just once
    Dim idx As String
    idx = getTocSheetName()
    
    'there is no sheet for table of contents?
    If Not worksheetExists(Sh.Parent, idx) Then Exit Sub
    
    If Sh.Name = idx Then
     Call generateTocWorksheet
    End If
End Sub