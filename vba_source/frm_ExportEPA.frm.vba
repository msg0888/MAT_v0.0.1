Attribute VB_Name = "frm_ExportEPA"
Attribute VB_Base = "0{14DAB7A2-3729-4861-99B6-A102677A528A}{D9EB7045-51E3-4A63-B199-3B8B223CED72}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub cmbContinue_Click()
    If cbAppendModel = True Then
        Sheets("Meta").Range("E18") = True
    ElseIf cbAppendModel = False Then
        Sheets("Meta").Range("E18") = False
        Unload Me
    End If

    If cbCreateModel = True Then
        frm_ExportModel.Show
    ElseIf cbCreateModel = False Then
        Unload Me
    End If
    
    Export_Loading_tnxTower
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
