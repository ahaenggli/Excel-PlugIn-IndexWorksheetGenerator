Attribute VB_Name = "TocSheetExtension"
Option Explicit
Public isF5 As Boolean

'handles click on F5-Key
Public Sub handleF5Click()
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    isF5 = True
    If ActiveWorkbook.ActiveSheet.Name <> getTocSheetName() Then
        Call ShowPropEditForm
    Else
        Call generateTocWorksheet
    End If
    isF5 = False
End Sub

'get the name for the worksheet created field (custom property)
Public Function getWorksheetCreatedDatePropName() As String
    Dim prop As String
    prop = ""
    
    On Error Resume Next
    If prop = "" Then prop = isNull(getProperty(getTocSheet(), "WorksheetCreatedDatePropName"), getProperty(ThisWorkbook.Worksheets(1), "WorksheetCreatedDatePropName"))
    If prop = "" And isGermanGUI() Then prop = "Datum"
    If prop = "" And Not isGermanGUI() Then prop = "Created"

    If Err.Number > 0 Then
        prop = "Created"
        Err.Clear
    End If
    On Error GoTo 0
    
    getWorksheetCreatedDatePropName = prop
End Function

'get name of properties which are shown in the Toc sheet
Public Function getTocColumns() As Variant
    Dim props As String
    props = ""
    
    On Error Resume Next
    If props = "" Then props = isNull(getProperty(getTocSheet(), "TocColumns"), getProperty(ThisWorkbook.Worksheets(1), "TocColumns"))
    If props = "" And isGermanGUI() Then props = "Blatt;Datum;Beschreibung;Verantwortlich;ToDo;Status;Info"
    If props = "" And Not isGermanGUI() Then props = "Worksheet;Created;Description;Responsible;ToDo;Status;Info"

    If Err.Number > 0 Then
        props = "Worksheet;Created;Description;Responsible;ToDo;Status;Info"
        Err.Clear
    End If
    On Error GoTo 0
            
    'first column has to be for the hyperlink to the other worksheets, first array entry should not be an existing custom property
    If (inArray(Split(props, ";")(0), getTocCustomProperties()) Or Split(props, ";")(0) = getWorksheetCreatedDatePropName()) Then
        props = ";" & props
    End If
    
    getTocColumns = Split(props, ";")
End Function

'get name of custom proprties which should be created in all worksheets
Public Function getTocCustomProperties() As Variant
    Dim props As String
    props = ""
    
    On Error Resume Next
    If props = "" Then props = isNull(getProperty(getTocSheet(), "TocCustomProperties"), getProperty(ThisWorkbook.Worksheets(1), "TocCustomProperties"))
    If props = "" And isGermanGUI() Then props = "Beschreibung;Verantwortlich;ToDo;Status;Info;Datum"
    If props = "" And Not isGermanGUI() Then props = "Description;Responsible;ToDo;Status;Info;Created"
    
    If Err.Number > 0 Then
        props = "Description;Responsible;ToDo;Status;Info;Created"
        Err.Clear
    End If
    On Error GoTo 0
    
    getTocCustomProperties = Split(props, ";")
End Function

'set flag for Toc sheet
Public Sub setTocSheetFlag(ws As Worksheet)
    Dim Sheet As Worksheet

    For Each Sheet In ActiveWorkbook.Worksheets
        setProperty Sheet, "isToc", "0"
    Next Sheet

    setProperty ws, "isToc", "1"
End Sub

'get the defined name for the Toc worksheet
Public Function getTocSheetName() As String
    Dim ws As Worksheet
    Dim sumsheet As String
    sumsheet = ""
    
    For Each ws In ActiveWorkbook.Worksheets
        If (getProperty(ws, "isToc") = "1") Then
            sumsheet = ws.Name
            Exit For
        End If
    Next ws
    
    On Error Resume Next
    If sumsheet = "" Then sumsheet = getProperty(ThisWorkbook.Worksheets(1), "TocWorksheetName")
    If sumsheet = "" And isGermanGUI() Then sumsheet = "Uebersicht"
    If sumsheet = "" And Not isGermanGUI() Then sumsheet = "Toc"
               
    If Err.Number > 0 Then
        sumsheet = "Toc"
        Err.Clear
    End If
    On Error GoTo 0
    
    getTocSheetName = sumsheet
End Function

'returns ref to Toc sheet if exists, else nothing
Public Function getTocSheet() As Worksheet
    Dim idx As String
    idx = getTocSheetName()
    If worksheetExists(ActiveWorkbook, idx) Then
        Set getTocSheet = ActiveWorkbook.Worksheets(idx)
    Else
        Set getTocSheet = Nothing
    End If
End Function

'genereates new sheet for overview
Public Sub generateTocWorksheet()
    Dim Sh As Worksheet
    Dim Newsh As Worksheet
    
    Dim Basebook As Workbook
    Dim Basesheet As Worksheet
    
    Dim RwNum, ColNum As Integer
    Dim col As Variant
    
    Dim TableStyle As String
    Dim TocSheetName As String
    Dim TocColumns As Variant
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    Application.DisplayAlerts = False
    Set Basebook = ActiveWorkbook
    Set Basesheet = ActiveWorkbook.ActiveSheet
    
    TocSheetName = getTocSheetName()
    TocColumns = getTocColumns()
    
    If Not worksheetExists(ActiveWorkbook, TocSheetName) Then
        'Add a worksheet with the name "Toc-Sheet"
        Set Newsh = Basebook.Worksheets.Add(Before:=Basebook.Worksheets(1))
        Newsh.Name = TocSheetName
     Else
        Set Newsh = Basebook.Worksheets(TocSheetName)
    End If
    
    If Newsh.ListObjects.count > 0 Then
        TableStyle = Newsh.ListObjects(1).TableStyle
        Newsh.ListObjects(1).Delete
    End If
    
    Newsh.Cells.Clear
    Newsh.Cells.Delete

    
    Call setTocSheetFlag(Newsh)
        
    Application.DisplayAlerts = True
  
    'Add headers
    With Newsh.Range(Newsh.Cells(1, 1), Newsh.Cells(1, 1 + UBound(TocColumns)))
        .Value = TocColumns
        .Font.Bold = True
        .Font.Size = 12
    End With

    'The links to the first sheet will start in row 2
    RwNum = 1

    For Each Sh In Basebook.Worksheets
        If Sh.Name <> Newsh.Name And Sh.Visible Then
            ColNum = 1
            RwNum = RwNum + 1
                       
            'Create a link to the sheet in the A column
            Newsh.Hyperlinks.Add Anchor:=Newsh.Cells(RwNum, 1), Address:="", SubAddress:="'" & Sh.Name & "'!A1", ScreenTip:="", TextToDisplay:=Sh.Name

            For Each col In TocColumns
                If CStr(col) <> "" And CStr(col) <> CStr(TocColumns(0)) Then
                    ColNum = ColNum + 1
                    Newsh.Cells(RwNum, ColNum) = getProperty(Sh, CStr(col))
                End If
            Next col
            
        End If
    Next

    Dim tbl As ListObject
    Dim rng As Range

    Set rng = Newsh.UsedRange
    Set tbl = Newsh.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.TableStyle = isNull(TableStyle, "TableStyleMedium15")
    tbl.Name = TocSheetName
    With rng
     With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
    Newsh.UsedRange.ColumnWidth = 250
    Newsh.UsedRange.RowHeight = 250
    Newsh.UsedRange.HorizontalAlignment = xlLeft
    Newsh.UsedRange.VerticalAlignment = xlTop
    
    Newsh.UsedRange.Columns.AutoFit
    
    For Each rng In Newsh.UsedRange.Columns
        If rng.ColumnWidth > 75 Then
            rng.ColumnWidth = 75
            rng.WrapText = True
        End If
    Next rng
    
    Newsh.UsedRange.Rows.AutoFit
    
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    
    Basebook.Activate
    Basesheet.Activate
End Sub

