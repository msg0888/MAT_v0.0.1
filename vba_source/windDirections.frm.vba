Attribute VB_Name = "windDirections"
Attribute VB_Base = "0{7AE21860-46EE-49A4-A572-012A5A7F5FCF}{480E7E4B-521C-48B8-A56F-B784A02CD3A3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'=== in the UserForm code-behind ===

Private Sub Directions_Change()
    Dim sel As String
    sel = Me.Directions.Value

    '1) clear all 3 families
    ClearCheckboxFamily "noice_Deg"
    ClearCheckboxFamily "ice_Deg"
    ClearCheckboxFamily "main_Deg"

    '2) apply mapping
    Select Case sel
        Case "Basic 3"
            CheckSet "noice_Deg", Array(0, 120, 240)
            CheckSet "ice_Deg", Array(0, 120, 240)
            CheckSet "main_Deg", Array(0, 120, 240)

        Case "Basic 4"
            CheckSet "noice_Deg", Array(0, 90, 180, 270)
            CheckSet "ice_Deg", Array(0, 90, 180, 270)
            CheckSet "main_Deg", Array(0, 90, 180, 270)

        Case "30° Increment"
            CheckSet "noice_Deg", Array(0, 30, 60, 90, 120, 150, 180, 210, 240, 270, 300, 330)
            CheckSet "ice_Deg", Array(0, 30, 60, 90, 120, 150, 180, 210, 240, 270, 300, 330)
            CheckSet "main_Deg", Array(0, 30, 60, 90, 120, 150, 180, 210, 240, 270, 300, 330)

        Case "45° Increment"
            CheckSet "noice_Deg", Array(0, 45, 90, 135, 180, 225, 270, 315)
            CheckSet "ice_Deg", Array(0, 45, 90, 135, 180, 225, 270, 315)
            CheckSet "main_Deg", Array(0, 45, 90, 135, 180, 225, 270, 315)

        Case "All"
            CheckSet "noice_Deg", Array(0, 30, 45, 60, 90, 120, 135, 150, 180, 210, 225, 240, 270, 300, 315, 330)
            CheckSet "ice_Deg", Array(0, 30, 45, 60, 90, 120, 135, 150, 180, 210, 225, 240, 270, 300, 315, 330)
            CheckSet "main_Deg", Array(0, 30, 45, 60, 90, 120, 135, 150, 180, 210, 225, 240, 270, 300, 315, 330)

        Case Else
            'nothing checked (or add defaults here)
    End Select
End Sub

Private Sub ClearCheckboxFamily(ByVal prefix As String)
    Dim ctl As MSForms.Control
    For Each ctl In Me.Controls
        If typeName(ctl) = "CheckBox" Then
            If LCase$(Left$(ctl.Name, Len(prefix))) = LCase$(prefix) Then
                ctl.Value = False
            End If
        End If
    Next ctl
End Sub

Private Sub CheckSet(ByVal prefix As String, ByVal degrees As Variant)
    Dim i As Long, nm As String
    For i = LBound(degrees) To UBound(degrees)
        nm = prefix & CStr(degrees(i))
        If HasControl(nm) Then Me.Controls(nm).Value = True
    Next i
End Sub

Private Function HasControl(ByVal controlName As String) As Boolean
    On Error GoTo Nope
    Dim tmp As Object
    Set tmp = Me.Controls(controlName)
    HasControl = True
    Exit Function
Nope:
    HasControl = False
End Function




Private Sub UserForm_Initialize()
    Me.Directions.list = Array("Basic 3", "Basic 4", "30° Increment", "45° Increment", "All", "Custom")
    Directions = "All"
End Sub

Private Sub windApply_Click()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False

    Const START_ROW As Long = 6
    Const END_ROW   As Long = 137

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("RISA-3D")

    '--- collect checked degrees from the three checkbox families
    Dim noIceDeg As Object, iceDeg As Object, mainDeg As Object
    Set noIceDeg = GetCheckedDegrees("noice_Deg")
    Set iceDeg = GetCheckedDegrees("ice_Deg")
    Set mainDeg = GetCheckedDegrees("main_Deg")

    Dim anyMain As Boolean
    anyMain = (mainDeg.Count > 0)

    '--- clear output col O first
    ws.Range("O" & START_ROW & ":O" & END_ROW).ClearContents

    Dim r As Long
    For r = START_ROW To END_ROW

        Dim s As String
        s = CStr(ws.Cells(r, "N").Value2)  ' works even if Column N is formula

        If Len(Trim$(s)) = 0 Then
            ' leave blank
        ElseIf InStr(1, s, "Lv", vbTextCompare) > 0 Then
            '--- Lv rows: TRUE if ANY main direction box is checked
            If anyMain Then ws.Cells(r, "O").Value2 = True

        Else
            Dim deg As Variant
            deg = ExtractDeg(s) ' returns Empty if not found

            If IsEmpty(deg) Then
                ' no degree found; leave blank
            Else
                Dim kind As String
                kind = RowKindFromText(s) ' "noice" / "ice" / "main" / ""

                Select Case kind
                    Case "noice"
                        If noIceDeg.Exists(CStr(deg)) Then ws.Cells(r, "O").Value2 = True
                    Case "ice"
                        If iceDeg.Exists(CStr(deg)) Then ws.Cells(r, "O").Value2 = True
                    Case "main"
                        If mainDeg.Exists(CStr(deg)) Then ws.Cells(r, "O").Value2 = True
                    Case Else
                        ' unknown row type; leave blank
                End Select
            End If
        End If
    Next r

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
End Sub

'========================
' Helpers
'========================

'Returns a Dictionary of checked degree strings for controls like:
'  noice_Deg0, noice_Deg30, noice_Deg45, ... etc
Private Function GetCheckedDegrees(ByVal prefix As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If typeName(ctrl) = "CheckBox" Then
            If LCase$(Left$(ctrl.Name, Len(prefix))) = LCase$(prefix) Then
                If ctrl.Value = True Then
                    Dim d As String
                    d = Mid$(ctrl.Name, Len(prefix) + 1) ' suffix after prefix
                    d = Trim$(d)
                    If Len(d) > 0 Then
                        If Not dict.Exists(d) Then dict.Add d, True
                    End If
                End If
            End If
        End If
    Next ctrl

    Set GetCheckedDegrees = dict
End Function

'Classify the row based on tokens in Column N text
Private Function RowKindFromText(ByVal s As String) As String
    'order matters (Wi/Wm contain W)
    If InStr(1, s, "Wi", vbTextCompare) > 0 Then
        RowKindFromText = "ice"
    ElseIf InStr(1, s, "Wm", vbTextCompare) > 0 Then
        RowKindFromText = "main"
    ElseIf InStr(1, s, "Eh", vbTextCompare) > 0 Then
        'new case: earthquake horizontal component -> tie to main_Deg selections
        RowKindFromText = "main"
    ElseIf InStr(1, s, "W", vbTextCompare) > 0 Then
        RowKindFromText = "noice"
    Else
        RowKindFromText = vbNullString
    End If
End Function

'Extracts the first number that appears before "deg" (e.g., "30 deg", "(270 deg)", "| 315 deg")
Private Function ExtractDeg(ByVal s As String) As Variant
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = "(\d{1,3})\s*deg"

    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        ExtractDeg = CStr(m.SubMatches(0))  ' return as string key ("30", "270", etc.)
    Else
        ExtractDeg = Empty
    End If
End Function
