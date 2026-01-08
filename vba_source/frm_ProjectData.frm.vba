Attribute VB_Name = "frm_ProjectData"
Attribute VB_Base = "0{9D2248C9-DF19-4B31-BAD4-3071A630221B}{17106106-99F3-489C-994F-4CB905A96AA2}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CommandButton1_Click()
    If boxProjectManager = "Please Select an Entity Name" Or boxProjectManager = "" Then
        Range("Client") = "Network Building + Consulting, LLC"
    Else
        Range("Client") = boxProjectManager
    End If
    Range("Carrier") = boxCarrier
    Call Select_File
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    boxProgramManager.AddItem "Mastec Network Solutions"
    boxProgramManager.AddItem "NB+C Engineering Services, LLC"
    boxProgramManager.AddItem "Network Building + Consulting, LLC"
    boxProgramManager.AddItem "TKK Engineering DPC"
    boxProgramManager.AddItem "TKK Engineering P.C."
    
    boxCarrier.AddItem "AT&T"
    boxCarrier.AddItem "Dish"
    boxCarrier.AddItem "Shentel"
    boxCarrier.AddItem "T-Mobile"
    boxCarrier.AddItem "Verizon"
    boxCarrier.AddItem "Ericsson"
    boxCarrier.AddItem "Motorola"
    boxCarrier.AddItem "Other (Overwrite)"
End Sub

