Attribute VB_Name = "IndexSheetExtension"
Option Explicit
Public isF5 As Boolean

'handles click on F5-Key
Public Sub handleF5Click()
    isF5 = True
    If ActiveWorkbook.ActiveSheet.Name <> getIndexSheetName() Then
        Call ShowPropEditForm
    Else
        Call generateIndexWorksheet
    End If
    isF5 = False
End Sub

'get the name for the worksheet created field (custom property)
Public Function getWorksheetCreatedDatePropName() As String
    Dim prop As String
    prop = ""
    
    On Error Resume Next
    If prop = "" Then prop = isNull(getProperty(getIndexSheet(), "WorksheetCreatedDatePropName"), getProperty(ThisWorkbook.Worksheets(1), "WorksheetCreatedDatePropName"))
    If prop = "" And isGermanGUI() Then prop = "Datum"
    If prop = "" And Not isGermanGUI() Then prop = "Created"

    If Err.Number > 0 Then
        prop = "Created"
        Err.Clear
    End If
    On Error GoTo 0
    
    getWorksheetCreatedDatePropName = prop
End Function

'get name of properties which are shown in the index sheet
Public Function getIndexColumns() As Variant
    Dim props As String
    props = ""
    
    On Error Resume Next
    If props = "" Then props = isNull(getProperty(getIndexSheet(), "IndexColumns"), getProperty(ThisWorkbook.Worksheets(1), "IndexColumns"))
    If props = "" And isGermanGUI() Then props = "Blatt;Datum;Beschreibung;Verantwortlich;ToDo;Status;Info"
    If props = "" And Not isGermanGUI() Then props = "Worksheet;Created;Description;Responsible;ToDo;Status;Info"

    If Err.Number > 0 Then
        props = "Worksheet;Created;Description;Responsible;ToDo;Status;Info"
        Err.Clear
    End If
    On Error GoTo 0
            
    'first column has to be for the hyperlink to the other worksheets, first array entry should not be an existing custom property
    If (inArray(Split(props, ";")(0), getIndexCustomProperties()) Or Split(props, ";")(0) = getWorksheetCreatedDatePropName()) Then
        props = ";" & props
    End If
    
    getIndexColumns = Split(props, ";")
End Function

'get name of custom proprties which should be created in all worksheets
Public Function getIndexCustomProperties() As Variant
    Dim props As String
    props = ""
    
    On Error Resume Next
    If props = "" Then props = isNull(getProperty(getIndexSheet(), "IndexCustomProperties"), getProperty(ThisWorkbook.Worksheets(1), "IndexCustomProperties"))
    If props = "" And isGermanGUI() Then props = "Beschreibung;Verantwortlich;ToDo;Status;Info;Datum"
    If props = "" And Not isGermanGUI() Then props = "Description;Responsible;ToDo;Status;Info;Created"
    
    If Err.Number > 0 Then
        props = "Description;Responsible;ToDo;Status;Info;Created"
        Err.Clear
    End If
    On Error GoTo 0
    
    getIndexCustomProperties = Split(props, ";")
End Function

'set flag for index sheet
Public Sub setIndexSheetFlag(ws As Worksheet)
    Dim Sheet As Worksheet

    For Each Sheet In ActiveWorkbook.Worksheets
        setProperty Sheet, "isIndex", "0"
    Next Sheet

    setProperty ws, "isIndex", "1"
End Sub

'get the defined name for the index worksheet
Public Function getIndexSheetName() As String
    Dim ws As Worksheet
    Dim sumsheet As String
    sumsheet = ""
    
    For Each ws In ActiveWorkbook.Worksheets
        If (getProperty(ws, "isIndex") = "1") Then
            sumsheet = ws.Name
            Exit For
        End If
    Next ws
    
    On Error Resume Next
    If sumsheet = "" Then sumsheet = getProperty(ThisWorkbook.Worksheets(1), "IndexWorksheetName")
    If sumsheet = "" And isGermanGUI() Then sumsheet = "Uebersicht"
    If sumsheet = "" And Not isGermanGUI() Then sumsheet = "Index"
               
    If Err.Number > 0 Then
        sumsheet = "Index"
        Err.Clear
    End If
    On Error GoTo 0
    
    getIndexSheetName = sumsheet
End Function

'returns ref to index sheet if exists, else nothing
Public Function getIndexSheet() As Worksheet
    Dim idx As String
    idx = getIndexSheetName()
    If worksheetExists(ActiveWorkbook, idx) Then
        Set getIndexSheet = ActiveWorkbook.Worksheets(idx)
    Else
        Set getIndexSheet = Nothing
    End If
End Function

'genereates new sheet for overview
Public Sub generateIndexWorksheet()
    Dim Sh As Worksheet
    Dim Newsh As Worksheet
    
    Dim Basebook As Workbook
    Dim Basesheet As Worksheet
    
    Dim RwNum, ColNum As Integer
    Dim col As Variant
    
    Dim TableStyle As String
    Dim IndexSheetName As String
    Dim IndexColumns As Variant
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    Application.DisplayAlerts = False
    Set Basebook = ActiveWorkbook
    Set Basesheet = ActiveWorkbook.ActiveSheet
    
    IndexSheetName = getIndexSheetName()
    IndexColumns = getIndexColumns()
    
    If Not worksheetExists(ActiveWorkbook, IndexSheetName) Then
        'Add a worksheet with the name "Index-Sheet"
        Set Newsh = Basebook.Worksheets.Add(Before:=Basebook.Worksheets(1))
        Newsh.Name = IndexSheetName
     Else
        Set Newsh = Basebook.Worksheets(IndexSheetName)
    End If
    
    If Newsh.ListObjects.Count > 0 Then
        TableStyle = Newsh.ListObjects(1).TableStyle
        Newsh.ListObjects(1).Delete
    End If
    
    Newsh.Cells.Clear
    Newsh.Cells.Delete

    
    Call setIndexSheetFlag(Newsh)
        
    Application.DisplayAlerts = True
  
    'Add headers
    With Newsh.Range(Newsh.Cells(1, 1), Newsh.Cells(1, 1 + UBound(IndexColumns)))
        .Value = IndexColumns
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

            For Each col In IndexColumns
                If CStr(col) <> "" And CStr(col) <> CStr(IndexColumns(0)) Then
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
    tbl.Name = IndexSheetName
    With rng
     With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
    Newsh.UsedRange.ColumnWidth = 250
    Newsh.UsedRange.RowHeight = 250
    Newsh.UsedRange.Columns.AutoFit
    Newsh.UsedRange.Rows.AutoFit
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    
    Basebook.Activate
    Basesheet.Activate
End Sub

