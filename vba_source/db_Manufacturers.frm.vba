Attribute VB_Name = "db_Manufacturers"
Attribute VB_Base = "0{2262E2EC-97D0-454C-825E-5C50A5A5BB79}{C03FAB77-6558-4A36-B886-9A893622F706}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' NOTE: Intentionally NO Option Explicit

Public PickedArcFullPath As String
Public PickerCancelled As Boolean

Private mFolderPath As String

Public Sub BeginPick(ByVal folderPath As String, ByVal promptText As String)
    mFolderPath = folderPath
    PickedArcFullPath = ""
    PickerCancelled = True

    On Error Resume Next
    If Len(promptText) > 0 Then Me.Caption = promptText
    On Error GoTo 0

    PopulateFileList

    If Me.cmbFiles.ListCount > 0 Then
        Me.cmbFiles.ListIndex = 0
    End If
End Sub

Private Sub PopulateFileList()
    Dim f As String
    Me.cmbFiles.Clear

    If Len(mFolderPath) = 0 Then Exit Sub

    f = Dir(UCase(mFolderPath) & "\*.arc", vbNormal)
    Do While Len(f) > 0
        ' store WITHOUT .arc (matches your earlier pattern)
        Me.cmbFiles.AddItem Left$(f, Len(f) - 4)
        f = Dir()
    Loop
End Sub

Private Sub cmdPickOK_Click()
    Dim baseName As String
    baseName = Trim$(Me.cmbFiles.Value)

    If Len(baseName) = 0 Then
        MsgBox "Please select a database.", vbExclamation, "Select Database"
        Exit Sub
    End If

    PickedArcFullPath = mFolderPath & "\" & baseName & ".arc"
    PickerCancelled = False
    Me.Hide
End Sub

Private Sub cmdPickCancel_Click()
    PickedArcFullPath = ""
    PickerCancelled = True
    Me.Hide
End Sub

'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    If CloseMode = vbFormControlMenu Then
'        Cancel = True
'        cmdPickCancel_Click
'    End If
'End Sub
'
'' Intentionally no Option Explicit (per your preference)
'
'Public SelectedArcPath
'
'Private mFolder

Public Sub InitPicker(ByVal folderPath As String, ByVal manufacturer As String)
    mFolder = folderPath

    On Error Resume Next
    Me.Caption = "Select .arc Database"
    Me.lblPrompt.Caption = "No database found for: " & manufacturer & vbCrLf & _
                           "Select an alternate .arc database:"
    On Error GoTo 0

    PopulateArcList
End Sub

Private Sub PopulateArcList()
    Dim f
    cmbFiles.Clear

    f = Dir(mFolder & "\*.arc", vbNormal)
    Do While Len(f) > 0
        cmbFiles.AddItem UCase$(f)   ' display only filename
        f = Dir()
    Loop

    If cmbFiles.ListCount > 0 Then cmbFiles.ListIndex = 0
End Sub


Private Sub cmdOK_Click()
    If cmbFiles.ListIndex < 0 Then
        MsgBox "Please select a database.", vbExclamation
        Exit Sub
    End If

    SelectedArcPath = cmbFiles.Value
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    SelectedArcPath = ""
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        cmdCancel_Click
    End If
End Sub


