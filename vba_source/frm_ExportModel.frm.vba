Attribute VB_Name = "frm_ExportModel"
Attribute VB_Base = "0{0C7A5312-03D7-4604-9522-69063A2C2BB4}{88377BB9-058C-4A47-A24B-BC7A21DE44AB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub cmbAdd_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    With Sheets("Meta")
        .Range("E19") = tbUSLabel
        .Range("E20") = tbSILabel
    End With
Call cmdAddModel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.tbUSLabel = WorksheetFunction.Substitute(Sheets("Meta").Range("E21"), "USName=", "")
    Me.tbSILabel = WorksheetFunction.Substitute(Sheets("Meta").Range("E22"), "SIName=", "")
'    With Sheets("Code")
'        Me.tbUSLabel = Format(.Range("BB41").Value, "0.0") & " ft " & .Range("MountType") & "_" & Format(.Range("Q31"), "0.00") & " ft^2 _" & .Range("Carrier") & "_" & .Range("H20") & "_" & Format(Date, "YYYYMMDD") & "_" & Format(Now, "hhmmss")
'        Me.tbSILabel = Format(.Range("BB41").Value * 0.3048, "0.00") & " m " & .Range("MountType") & "_" & Format(.Range("Q31") * 0.3048 * 0.3048, "0.00") & " m^2 _" & .Range("Carrier") & "_" & .Range("H20") & "_" & Format(Date, "YYYYMMDD") & "_" & Format(Now, "hhmmss")
'    End With
End Sub
