Attribute VB_Name = "SummarySheetExtension"
Option Explicit

'handles click on F5-Key
Public Sub handleF5Click()
    If Application.ActiveSheet.Name <> getSummarySheetName() Then
        ShowPropEditForm
    Else
        Call generateSummaryWorksheet
    End If
End Sub

'get the name for the worksheet created field (custom property)
Public Function getWorksheetCreatedDatePropName() As String
    Dim prop As String
    prop = ""
    
    On Error Resume Next
        If getProperty(ThisWorkbook.Worksheets(1), "WorksheetCreatedDatePropName") = "" Then
            If isGermanGUI() Then
                prop = "Datum"
            Else
                prop = "Created"
            End If
        Else
            prop = getProperty(ThisWorkbook.Worksheets(1), "WorksheetCreatedDatePropName")
        End If
    If Err.Number > 0 Then
        prop = "Created"
        Err.Clear
    End If
    On Error GoTo 0
    
    getWorksheetCreatedDatePropName = prop
End Function

'get name of properties which are shown in the summary sheet
Public Function getSummaryColumns() As Variant
    Dim props As String
    props = ""
    
    On Error Resume Next
        If getProperty(ThisWorkbook.Worksheets(1), "SummaryColumns") = "" Then
            If isGermanGUI() Then
                props = "Tabelle;Datum;Beschreibung;Verantwortlich;ToDo;Status;Info"
            Else
                props = "Worksheet;Created;Description;Responsible;ToDo;Status;Info"
            End If
        Else
            props = getProperty(ThisWorkbook.Worksheets(1), "SummaryColumns")
        End If
    If Err.Number > 0 Then
        props = "Worksheet;Created;Description;Responsible;ToDo;Status;Info"
        Err.Clear
    End If
    On Error GoTo 0
            
    'first column has to be for the hyperlink to the other worksheets, first array entry should not be an existing custom property
    If (inArray(Split(props, ";")(0), getSummaryCustomProperties()) Or Split(props, ";")(0) = getWorksheetCreatedDatePropName()) Then
        props = ";" & props
    End If
    
    getSummaryColumns = Split(props, ";")
    
End Function

'get name of custom proprties which should be created in all worksheets
Public Function getSummaryCustomProperties() As Variant
    Dim props As String
    props = ""
    
    On Error Resume Next
        If getProperty(ThisWorkbook.Worksheets(1), "SummaryCustomProperties") = "" Then
            If isGermanGUI() Then
                props = "Beschreibung;Verantwortlich;ToDo;Status;Info;Datum"
            Else
                props = "Description;Responsible;ToDo;Status;Info;Created"
            End If
        Else
            props = getProperty(ThisWorkbook.Worksheets(1), "SummaryCustomProperties")
        End If
    If Err.Number > 0 Then
        props = "Description;Responsible;ToDo;Status;Info;Created"
        Err.Clear
    End If
    On Error GoTo 0
    
    getSummaryCustomProperties = Split(props, ";")
End Function

'get the defined name for the summary worksheet
Public Function getSummarySheetName() As String
    Dim sumsheet As String
    
    On Error Resume Next
        If getProperty(ThisWorkbook.Worksheets(1), "SummaryWorksheetName") = "" Then
            If isGermanGUI() Then
                sumsheet = "Uebersicht"
            Else
                sumsheet = "summary"
            End If
        Else
            sumsheet = getProperty(ThisWorkbook.Worksheets(1), "SummaryWorksheetName")
        End If
    If Err.Number > 0 Then
        sumsheet = "summary"
        Err.Clear
    End If
    On Error GoTo 0
    getSummarySheetName = sumsheet
End Function


'genereates new sheet for overview
Public Sub generateSummaryWorksheet()
    Dim Sh As Worksheet
    Dim Newsh As Worksheet
    
    Dim Basebook As Workbook
    Dim Basesheet As Worksheet
    
    Dim myCell As Range
    Dim RwNum, ColNum As Integer
    Dim col As Variant
        
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    Application.DisplayAlerts = False
    Set Basebook = ActiveWorkbook
    Set Basesheet = ActiveWorkbook.ActiveSheet
    
    If Not worksheetExists(ActiveWorkbook, getSummarySheetName()) Then
        'Add a worksheet with the name "Summary-Sheet"
        Set Newsh = Basebook.Worksheets.Add(Before:=Basebook.Worksheets(1))
        Newsh.Name = getSummarySheetName()
     Else
        Set Newsh = Basebook.Worksheets(getSummarySheetName())
    End If
    
    Newsh.Cells.Clear
    Newsh.Cells.Delete
    If Newsh.ListObjects.Count > 0 Then
        Newsh.ListObjects(0).Delete
    End If
    
    Application.DisplayAlerts = True
  
    'Add headers
    With Newsh.Range(Newsh.Cells(1, 1), Newsh.Cells(1, 1 + UBound(getSummaryColumns())))
        .Value = getSummaryColumns()
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
            Newsh.Hyperlinks.Add Anchor:=Newsh.Cells(RwNum, 1), Address:="", SubAddress:="'" & Sh.Name & "'!A1", ScreenTip:="Direkt zur Liste springen", TextToDisplay:=Sh.Name

            For Each col In getSummaryColumns()
                If CStr(col) <> "" And CStr(col) <> CStr(getSummaryColumns(0)) Then
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
    tbl.TableStyle = "TableStyleMedium15"
    tbl.Name = getSummarySheetName()
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

