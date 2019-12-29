VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TocSheetExtensionForm 
   Caption         =   "edit custom values for index sheet"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6435
   OleObjectBlob   =   "TocSheetExtensionForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "TocSheetExtensionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOk_Click()
    Call saveSettings
    Unload Me
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub UserForm_Activate()
    txtSumTitel.Text = getTocSheetName()
    txtProperties.Text = Join(getTocCustomProperties(), ";")
    txtSummaryColumns.Text = Join(getTocColumns(), ";")
    txtWorkSheetCreatedDate.Text = getWorksheetCreatedDatePropName()
    txtCallToc.Text = getGlobalTocHandlerPropName()
    
     If Not worksheetExists(ActiveWorkbook, getTocSheetName()) Then
        Me.cbSetDefault.Value = True
     End If
End Sub

Private Sub saveSettings()
    
    If Not worksheetExists(ActiveWorkbook, getTocSheetName()) And Me.cbSetDefault.Value = False Then
        Call generateTocWorksheet
    End If
    
    If worksheetExists(ActiveWorkbook, getTocSheetName()) And Me.cbSetDefault.Value = False Then
        If txtSumTitel.Text <> "" Then setProperty ActiveWorkbook.Worksheets(1), "TocWorksheetName", txtSumTitel.Text
        If txtProperties.Text <> "" Then setProperty ActiveWorkbook.Worksheets(1), "TocCustomProperties", txtProperties.Text
        If txtSummaryColumns.Text <> "" Then setProperty ActiveWorkbook.Worksheets(1), "TocColumns", txtSummaryColumns.Text
        
        setProperty ActiveWorkbook.Worksheets(1), "WorksheetCreatedDatePropName", txtWorkSheetCreatedDate.Text
            
        On Error Resume Next
            Application.DisplayAlerts = False
            ActiveWorkbook.Save
            Application.DisplayAlerts = True
        On Error GoTo 0
    End If
    
    If Me.cbSetDefault.Value = True Then
    
    'save "global"-properties in ThisWorkbook.Worksheets(1)
    ' -> ThisWorkbook is where the code is saved (xlam-file)
    ' -> even a xlam file has at least one sheet
    ' -> here it's named "TocConfig"
    If txtSumTitel.Text <> "" And txtSumTitel.Text <> getTocSheetName() Then setProperty ThisWorkbook.Worksheets(1), "TocWorksheetName", txtSumTitel.Text
    If txtProperties.Text <> "" And txtProperties.Text <> Join(getTocCustomProperties(), ";") Then setProperty ThisWorkbook.Worksheets(1), "TocCustomProperties", txtProperties.Text
    If txtSummaryColumns.Text <> "" And txtSummaryColumns.Text <> Join(getTocColumns(), ";") Then setProperty ThisWorkbook.Worksheets(1), "TocColumns", txtSummaryColumns.Text
    If txtWorkSheetCreatedDate.Text <> getWorksheetCreatedDatePropName() Then setProperty ThisWorkbook.Worksheets(1), "WorksheetCreatedDatePropName", txtWorkSheetCreatedDate.Text
        
    If txtCallToc.Text <> "" And txtCallToc.Text <> getGlobalTocHandlerPropName() Then
        Application.OnKey getGlobalTocHandlerPropName()
        setProperty ThisWorkbook.Worksheets(1), "GlobalTocHandlerPropName", txtCallToc.Text
        Application.OnKey getGlobalTocHandlerPropName(), "handleF5Click"
    End If
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    End If
    
End Sub

Private Sub UserForm_Terminate()
    ' Call saveSettings
End Sub
