Attribute VB_Name = "db_RFDS"
Attribute VB_Base = "0{0E05C03D-D150-422E-B626-4510B50AB2F3}{1ECF3BCC-25E5-43E3-B57E-5E373173D8B3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub btnATT_Click() 'btnATT_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False

Dim Answer As String

Answer = MsgBox("Import RFDS as PDF?" & vbNewLine & vbNewLine & "If so, select 'Yes'. Otherwise, select 'No' to import a converted RFDS Excel Spreadsheet.", vbYesNoCancel, "Import RFDS")
If Answer = vbYes Then
    Unload Me
    RFDS_RunForCarrier "ATT"
ElseIf Answer = vbNo Then
    'Call Import_Antenna_Data_APD
    Call ImportRFDS_ATT
    Unload Me
Else
    Unload Me
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
End Sub

Private Sub btnTMO_Click() 'btnTMO_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False

Dim Answer As String

Answer = MsgBox("Import RFDS as PDF?" & vbNewLine & vbNewLine & "If so, select 'Yes'. Otherwise, select 'No' to import a converted RFDS Excel Spreadsheet.", vbYesNoCancel, "Import RFDS")
If Answer = vbYes Then
    Unload Me
    RFDS_RunForCarrier "TMO"
ElseIf Answer = vbNo Then
    Call ImportRFDS_TMO
    Unload Me
Else
    Unload Me
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
End Sub

Private Sub btnVZW_Click() 'btnVZW_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False

Dim Answer As String

Answer = MsgBox("Import RFDS as PDF?" & vbNewLine & vbNewLine & "If so, select 'Yes'. Otherwise, select 'No' to import a converted RFDS Excel Spreadsheet.", vbYesNoCancel, "Import RFDS")
If Answer = vbYes Then
    Unload Me
    RFDS_RunForCarrier "VZW"
ElseIf Answer = vbNo Then
    Call ImportRFDS_VZW
    Unload Me
Else
    Unload Me
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
End Sub

Private Sub RFDS_ATT_Click()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False

Dim Answer As String

Answer = MsgBox("Import RFDS as PDF?" & vbNewLine & vbNewLine & "If so, select 'Yes'. Otherwise, select 'No' to import a converted RFDS Excel Spreadsheet.", vbYesNoCancel, "Import RFDS")
If Answer = vbYes Then
    Unload Me
    RFDS_RunForCarrier "ATT"
Else
    Call ImportRFDS_ATT
    Unload Me
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
End Sub

Private Sub RFDS_TMO_Click()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False

Dim Answer As String

Answer = MsgBox("Import RFDS as PDF?" & vbNewLine & vbNewLine & "If so, select 'Yes'. Otherwise, select 'No' to import a converted RFDS Excel Spreadsheet.", vbYesNoCancel, "Import RFDS")
If Answer = vbYes Then
    Unload Me
    RFDS_RunForCarrier "TMO"
Else
    Call ImportRFDS_TMO
    Unload Me
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
End Sub

Private Sub RFDS_VZW_Click()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False

Dim Answer As String

Answer = MsgBox("Import RFDS as PDF?" & vbNewLine & vbNewLine & "If so, select 'Yes'. Otherwise, select 'No' to import a converted RFDS Excel Spreadsheet.", vbYesNoCancel, "Import RFDS")
If Answer = vbYes Then
    Unload Me
    RFDS_RunForCarrier "VZW"
Else
    Call Import_Antenna_Data_Finalized
    Unload Me
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
End Sub



'''
'''Private Sub RFDS_ATT_Click()
'''Application.ScreenUpdating = False
'''Application.Calculation = xlCalculationManual
'''Application.EnableEvents = False
'''Application.DisplayStatusBar = False
'''
'''    Call Import_Antenna_Data_APD
'''    Unload Me
'''
'''Application.ScreenUpdating = True
'''Application.Calculation = xlCalculationAutomatic
'''Application.EnableEvents = True
'''Application.DisplayStatusBar = True
'''End Sub
'''
'''Private Sub RFDS_TMO_Click()
'''Application.ScreenUpdating = False
'''Application.Calculation = xlCalculationManual
'''Application.EnableEvents = False
'''Application.DisplayStatusBar = False
'''
'''    Call TMO_Import_Sector_Antennas_To_DiscreteLoads
'''    Unload Me
'''
'''Application.ScreenUpdating = True
'''Application.Calculation = xlCalculationAutomatic
'''Application.EnableEvents = True
'''Application.DisplayStatusBar = True
'''End Sub
'''
'''Private Sub RFDS_VZW_Click()
'''Application.ScreenUpdating = False
'''Application.Calculation = xlCalculationManual
'''Application.EnableEvents = False
'''Application.DisplayStatusBar = False
'''
'''    Call Import_Antenna_Data_Finalized
'''    Unload Me
'''
'''Application.ScreenUpdating = True
'''Application.Calculation = xlCalculationAutomatic
'''Application.EnableEvents = True
'''Application.DisplayStatusBar = True
'''End Sub

Private Sub UserForm_Initialize()
    Dim xlApp As Application
    Dim wndLeft As Long, wndTop As Long
    Dim wndWidth As Long, wndHeight As Long

    Sheets("Discrete Loads").Range("A4:AU53").ClearContents
    
    Set xlApp = Application

    ' Get Excel window dimensions
    With xlApp
        wndLeft = .Left
        wndTop = .Top
        wndWidth = .Width
        wndHeight = .Height
    End With

    ' Position the form in the center of the Excel window
    Me.StartUpPosition = 0 ' manual
    Me.Left = wndLeft + (wndWidth - Me.Width) / 2
    Me.Top = wndTop + (wndHeight - Me.Height) / 2

End Sub
