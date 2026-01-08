Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()

''    Call Refresh_Server
'    Call Refresh_Local
''    MsgBox "Server and local databases have been updated."
'

    Application.AutoCorrect.AutoFillFormulasInLists = False
End Sub

'Private Sub showribbon()
'    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"
''    ActiveWindow.DisplayHeadings = True
'End Sub

'Private Sub Workbook_Open()
'    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",False)"
'
'    Application.ScreenUpdating = False
'    ActiveWindow.Visible = False
''    SplashUserForm.Show
'    Windows(ThisWorkbook.Name).Visible = True
'    Application.ScreenUpdating = True
'End Sub
