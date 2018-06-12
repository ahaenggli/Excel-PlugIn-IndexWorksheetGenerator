VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'' Source can be found at https://github.com/ahaenggli/Excel-PlugIn-IndexWorksheetGenerator

Option Explicit

'event handler for application
Private IndexSheetExtension_AppEventHandler As IndexSheetAppEventHandler

' Sub is called whenever AddIn is loaded
Private Sub Workbook_Open()
    'To combine keys with    Precede the key code by
    'SHIFT   + (plus sign)
    'CTRL    ^ (caret)
    'ALT     % (percent sign)
    
    '' CTRL + Shift + A
    'Application.OnKey "^+{A}", "tstBox"
     Application.OnKey "{F5}", "handleF5Click"
     isF5 = False
    'init application wide event handler
     Set IndexSheetExtension_AppEventHandler = New IndexSheetAppEventHandler
       
End Sub


