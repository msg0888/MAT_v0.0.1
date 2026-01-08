Attribute VB_Name = "db_DiscreteLoads"
Attribute VB_Base = "0{AFFFA99D-9440-4EA4-BD41-36A6625ADF6D}{E5328130-8C0E-4BEB-9EAE-44C9733E449F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



'———————————————————————————
' Module-level variables
' Change "Control" to whatever sheet holds AT3
Private Const FolderStoreSheet As String = "Meta"
Private Const FolderStoreCell  As String = "Filepath_APPURT"
Private mFolderPath As String
Private mData       As Object   ' late-bound Scripting.Dictionary
'———————————————————————————


' Populate cmbFiles based on mFolderPath
Private Sub PopulateFileList()
    Dim fName    As String
    cmbFiles.Clear
    If mFolderPath = "" Then Exit Sub
    fName = Dir(mFolderPath & "\*.arc")
    Do While fName <> ""
        ' strip “.arc” and force uppercase
        cmbFiles.AddItem UCase(Left$(fName, Len(fName) - 4))
        fName = Dir()
    Loop
End Sub

Private Sub Filepaths_Click()
    frm_Filepaths.Show
End Sub

Private Sub CommandButton1_Click()
    frm_Filepaths.Show
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub txtValue7_Change()
    If txtValue7.Value = 0 Then
        txtValue7.Value = txtThirdDiv.Value
    End If
End Sub

Private Sub UserForm_Initialize()

    Dim xlApp As Application
    Dim wndLeft As Long, wndTop As Long
    Dim wndWidth As Long, wndHeight As Long
    On Error Resume Next

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

    Dim shp As Object
    ' Read last-used folder from the cell
    With ThisWorkbook.Worksheets(FolderStoreSheet)
        mFolderPath = .Range(FolderStoreCell).Value
    End With

    ' If it exists on disk, populate files; otherwise clear path
    If Len(mFolderPath) > 0 Then
        If Dir(mFolderPath, vbDirectory) <> "" Then
            PopulateFileList
            cmdBrowseFolder.Caption = "Change Folder…"
        Else
            mFolderPath = ""
        End If
    End If

    Me.txtValue4.list = Array("Flat", "Round", "Generic")
    Me.AType.list = Array("Antenna", "TME (Front)", "TME (Back)", "Round (Front)", "Round (Back)", "Mount", "Generic")
    Me.ATable.list = Array("Proposed", "Existing", "Removed")
    Me.ASectors.Value = Range("QtySectors").Value
    With shGeometry
        Me.AMembersA.list = .Range("E6:E" & WorksheetFunction.Max(.Range("G:G")) + 6).Value
        Me.AMembersB.list = .Range("E6:E" & WorksheetFunction.Max(.Range("G:G")) + 6).Value
        Me.AMembersC.list = .Range("E6:E" & WorksheetFunction.Max(.Range("G:G")) + 6).Value
        Me.AMembersD.list = .Range("E6:E" & WorksheetFunction.Max(.Range("G:G")) + 6).Value
    End With
    Me.AAttach.list = Array(1, 2)
'    Me.txtValue1.Value = 0
'    Me.txtValue2.Value = 0
'    Me.txtValue3.Value = 0
'    Me.txtValue7.Value = 0
'    Me.txtValue4.Value = "Flat"
'    Me.ASectors.Value = 3
'    Me.AThetaa.Value = 0
'    Me.AThetab.Value = 120
'    Me.AThetac.Value = 240

End Sub

Private Sub cmdBrowseFolder_Click()

    Dim fd As FileDialog, sel As String

    ' Pick folder
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = "Select Folder Containing .arc Files"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        sel = .SelectedItems(1)
    End With

    ' Save and use the new folder
    mFolderPath = sel
    With ThisWorkbook.Worksheets(FolderStoreSheet)
        .Range(FolderStoreCell).Value = mFolderPath
    End With

    ' Refresh file list
    PopulateFileList
    cmdBrowseFolder.Caption = "Change tnxTower Database Folder…"
End Sub

Private Sub cmbAdd_Click()
'Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Dim Answer As String
Dim nMembers As String

nMembers = WorksheetFunction.Max(Sheets("Geometry").Range("G:G")) + 7

With Sheets("Discrete Loads")
    For r = 4 To 53
        If .Cells(r, 1) = "" Then
            .Cells(r, 1).Value = cmbFiles.Value
            .Cells(r, 2).Value = cmbUSNames.Value
            .Cells(r, 3).Value = ATable.Value
            Select Case AType.Value
                Case "TME (Front)":  .Cells(r, 4).Value = "TME"
                Case "TME (Back)":  .Cells(r, 4).Value = "TME"
                Case "Round (Front)":  .Cells(r, 4).Value = "Round"
                Case "Round (Back)":  .Cells(r, 4).Value = "Round"
                Case Else: .Cells(r, 4).Value = AType.Value
            End Select
            If txtValue4 = 3 Then
                .Cells(r, 5) = "Round"
            ElseIf txtValue4 = 2 Then
                .Cells(r, 5) = "Round"
            ElseIf txtValue4 = 1 Then
                .Cells(r, 5) = "Flat"
            ElseIf txtValue4 = 0 Then
                .Cells(r, 5) = "Generic"
            Else
                .Cells(r, 5) = txtValue4.Value
            End If
            .Cells(r, 6).Value = AAttach.Value
            .Cells(r, 7) = 0
            .Cells(r, 8) = AElevation.Value
            If AType.Value = "TME (Back)" Or AType.Value = "Round (Back)" Then
                .Cells(r, 9) = -8
            Else
                .Cells(r, 9) = 8
            End If
            .Cells(r, 10) = 1
            .Cells(r, 11) = 1

            '= Sector A ==========================================================================='
            If AMembersA <> "" Then
                .Cells(r, 12) = AThetaa.Value
                .Cells(r, 13) = AQtya.Value
                .Cells(r, 14) = AMembersA
                .Cells(r, 15) = ATopA.Value
                If ABottomA.Value <> "" Then
                    .Cells(r, 16) = AMembersA
                    .Cells(r, 17) = ABottomA.Value
                End If
                .Cells(r, 18) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(O" & r & "/(XLOOKUP(IF(ISNUMBER(N" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "N" & r & "," & Chr(34) & "$B" & Chr(34) & "), N" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
                .Cells(r, 19) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",Q" & r & "/(XLOOKUP(IF(ISNUMBER(P" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "P" & r & "," & Chr(34) & "$B" & Chr(34) & "), P" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
            Else
                .Cells(r, 12) = ""
                .Cells(r, 13) = ""
                .Cells(r, 14) = ""
                .Cells(r, 15) = ""
                .Cells(r, 16) = ""
                .Cells(r, 17) = ""
            End If

            '= Sector B ==========================================================================='
            If AMembersB <> "" Then
                .Cells(r, 20) = AThetab.Value
                .Cells(r, 21) = AQtyb.Value
                .Cells(r, 22) = AMembersB
                .Cells(r, 23) = ATopB.Value
                If ABottomA.Value <> "" Then
                    .Cells(r, 24) = AMembersB
                    .Cells(r, 25) = ABottomB.Value
                End If
                .Cells(r, 26) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(W" & r & "/(XLOOKUP(IF(ISNUMBER(V" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "V" & r & "," & Chr(34) & "$B" & Chr(34) & "), V" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
                .Cells(r, 27) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",Y" & r & "/(XLOOKUP(IF(ISNUMBER(X" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "X" & r & "," & Chr(34) & "$B" & Chr(34) & "), X" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
            Else
                .Cells(r, 20) = ""
                .Cells(r, 21) = ""
                .Cells(r, 22) = ""
                .Cells(r, 23) = ""
                .Cells(r, 24) = ""
                .Cells(r, 25) = ""
            End If

            '= Sector C ==========================================================================='
            If AMembersC <> "" Then
                .Cells(r, 28) = AThetac.Value
                .Cells(r, 29) = AQtyc.Value
                .Cells(r, 30) = AMembersC
                .Cells(r, 31) = ATopC.Value
                If ABottomA.Value <> "" Then
                    .Cells(r, 32) = AMembersC
                    .Cells(r, 33) = ABottomC.Value
                End If
                .Cells(r, 34) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(AE" & r & "/(XLOOKUP(IF(ISNUMBER(AD" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AD" & r & "," & Chr(34) & "$B" & Chr(34) & "), AD" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
                .Cells(r, 35) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",AG" & r & "/(XLOOKUP(IF(ISNUMBER(AF" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AF" & r & "," & Chr(34) & "$B" & Chr(34) & "), AF" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
            Else
                .Cells(r, 28) = ""
                .Cells(r, 29) = ""
                .Cells(r, 30) = ""
                .Cells(r, 31) = ""
                .Cells(r, 32) = ""
                .Cells(r, 33) = ""
            End If

            '= Sector D ==========================================================================='
            If AMembersD <> "" Then
                .Cells(r, 36) = AThetad.Value
                .Cells(r, 37) = AQtyd.Value
                .Cells(r, 38) = AMembersD
                .Cells(r, 39) = ATopD.Value
                If ABottomA.Value <> "" Then
                    .Cells(r, 40) = AMembersD
                    .Cells(r, 41) = ABottomD.Value
                End If
                .Cells(r, 42) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(AM" & r & "/(XLOOKUP(IF(ISNUMBER(AL" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AL" & r & "," & Chr(34) & "$B" & Chr(34) & "), AL" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
                .Cells(r, 43) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",AO" & r & "/(XLOOKUP(IF(ISNUMBER(AN" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AN" & r & "," & Chr(34) & "$B" & Chr(34) & "), AN" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
            Else
                .Cells(r, 36) = ""
                .Cells(r, 37) = ""
                .Cells(r, 38) = ""
                .Cells(r, 39) = ""
                .Cells(r, 40) = ""
                .Cells(r, 41) = ""
            End If

            '= Dimensions ========================================================================='
            .Cells(r, 44) = txtValue1.Value
            .Cells(r, 45) = txtValue2.Value
            .Cells(r, 46) = txtValue3.Value
            .Cells(r, 47) = txtValue7.Value

            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Exit Sub
        End If
    Next r
End With
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
End Sub

Private Sub cmdAddEntry_Click()
    Dim fso        As Object
    Dim ts         As Object
    Dim i          As Long
    Dim zeros      As String
    Dim valuesLine As String
    Dim filePath   As String

    ' -- Make sure we have a folder, file and USName selected --
    If mFolderPath = "" Or cmbFiles.Value = "" Or cmbUSNames.Value = "" Then
        MsgBox "Please select a folder, file and USName first.", vbExclamation
        Exit Sub
    End If

    filePath = mFolderPath & "\" & cmbFiles.Value & ".arc"

    ' -- Build the 15 zeros string --
    zeros = ""
    For i = 1 To 15
        zeros = zeros & "0 "
    Next i

    ' -- Build the Values= line --
    valuesLine = "Values=" & zeros
    For i = 1 To 7
        With WorksheetFunction
            valuesLine = valuesLine & .Substitute(.Substitute(.Substitute(.Substitute(.Substitute(Me.Controls("txtValue" & i).Value, "Flat", 1), "Round", 3), "Generic", 1), "CFD", 1), "CCI", 1) & " "
        End With
    Next i
    valuesLine = Trim(valuesLine)

    ' -- Open for append and write the three lines --
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' 8 = ForAppending, True = create file if it doesn't exist
    Set ts = fso.OpenTextFile(filePath, 8, True)

    ts.WriteLine          ' blank line before new entry
    ts.WriteLine "USName=" & cmbUSNames.Value
    ts.WriteLine "SIName=" & cmbUSNames.Value
    ts.WriteLine valuesLine

    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    MsgBox "New entry added to " & cmbFiles.Value & ".arc", vbInformation

    Unload Me
    db_DiscreteLoads.Show
End Sub


Private Sub cmbFiles_Change()
    Dim ts            As Object
    Dim fso           As Object
    Dim line          As String
    Dim nums          As Variant
    Dim currentUSName As String
    Dim i             As Long

    ' ——— Initialize/clear the dictionary each time a new file is chosen ———
    Set mData = CreateObject("Scripting.Dictionary")

    ' Clear the USName dropdown and the value TextBoxes
    cmbUSNames.Clear
    For i = 1 To 7
        Me.Controls("txtValue" & i).Value = ""
    Next i

    ' Open the selected .arc file (we re-append “.arc” here)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(mFolderPath & "\" & cmbFiles.Value & ".arc", 1)

    Do While Not ts.AtEndOfStream
        line = ts.ReadLine

        ' Capture each USName=
        If LCase(Left$(line, 7)) = "usname=" Then
            currentUSName = Mid$(line, 8)
        End If

    ' …inside your loop in cmbFiles_Change…
    If LCase(Left$(line, 7)) = "values=" Then
        nums = Split(Mid$(line, 8), " ")
        ' only store if we have at least 8 values
        If UBound(nums) + 1 >= 8 Then
            ' store the *entire* nums array
            mData(currentUSName) = nums
            cmbUSNames.AddItem currentUSName
        End If
    End If

    Loop

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub


Private Sub cmbUSNames_Change()
    Dim nums As Variant
    Dim i    As Long

    ' make sure this key actually exists
    If Not mData.Exists(cmbUSNames.Value) Then Exit Sub

    nums = mData(cmbUSNames.Value)

    ' 1st value ÷ 144 ' CaAa Front 0" Ice
    Me.txtFirstDiv.Value = val(nums(0)) / 144

    ' 2nd value ÷ 144 ' CaAa Front 1/2" Ice
    Me.txtSecondDiv.Value = val(nums(1)) / 144

    ' 3rd value ' Weight 0" Ice
    Me.txtThirdDiv.Value = val(nums(2))

    ' 4th value ' Weight 1/2 Ice
    Me.txtFourthDiv.Value = val(nums(3))

    ' 5th value ÷ 144 ' CaAa Front 1" Ice
    Me.txtFifthDiv.Value = val(nums(4)) / 144

    ' 6th value ÷ 144 ' CaAa Front 2" Ice
    Me.txtSixthDiv.Value = val(nums(5)) / 144

    ' 7th value ÷ 144 ' CaAa Front 4" Ice
    Me.txtSeventhDiv.Value = val(nums(6)) / 144

    ' 8th value ÷ 144 ' CaAa Side 0" Ice
    Me.txtEighthDiv.Value = val(nums(7)) / 144

    ' 9th value ÷ 144 ' CaAa Side 1/2" Ice
    Me.txtNinthDiv.Value = val(nums(8)) / 144

    ' 10th value ÷ 144 ' CaAa Side 1" Ice
    Me.txtTenthDiv.Value = val(nums(9)) / 144

    ' 11th value ÷ 144 ' CaAa Side 2" Ice
    Me.txtEleventhDiv.Value = val(nums(10)) / 144

    ' 12th value ÷ 144 ' CaAa Side 4" Ice
    Me.txtTwelfthDiv.Value = val(nums(11)) / 144

    ' 13th value ' Weight 1" Ice
    Me.txt13thDiv.Value = val(nums(12))

    ' 14th value ' Weight 2" Ice
    Me.txt14thDiv.Value = val(nums(13))

    ' 15th value ' Weight 4" Ice
    Me.txt15thDiv.Value = val(nums(14))

    ' last 7 values
    For i = 1 To 7
        ' UBound(nums)-7+1 is the first of the last 7
        Me.Controls("txtValue" & i).Value = nums(UBound(nums) - 7 + i)
    Next i

    ' 4/7 values
    If Me.txtValue4.Value = 0 Then
        Me.Controls("txtValue" & 4).Value = "Generic"
    ElseIf Me.txtValue4.Value = 1 Then
        Me.Controls("txtValue" & 4).Value = "Flat"
    ElseIf Me.txtValue4.Value = 2 Then
        Me.Controls("txtValue" & 4).Value = "Flat"
    ElseIf Me.txtValue4.Value = 3 Then
        Me.Controls("txtValue" & 4).Value = "Round"
    End If

End Sub


Private Sub cmdBrowse_Click()
    Dim filePath As String
    Dim fso      As Object
    Dim ts       As Object
    Dim txtLine  As String
    Dim nums     As Variant
    Dim i        As Long

    ' 1) Pick an .arc file
    filePath = Application.GetOpenFilename( _
                 FileFilter:="tnxTower Database Files (*.arc), *.arc", _
                 title:="Select a tnxTower Database File")
    If filePath = "False" Then Exit Sub

    ' 2) Clear old contents
    Me.txtUSName.Value = ""
    For i = 1 To 7
        Me.Controls("txtValue" & i).Value = ""
    Next i

    ' 3) Open the file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)    ' ForReading

    Do While Not ts.AtEndOfStream
        txtLine = ts.ReadLine

        ' 4) Capture USName= value
        If LCase(Left$(txtLine, 7)) = "usname=" Then
            Me.txtUSName.Value = Mid$(txtLine, 8)
        End If

        ' 5) Parse Values= last seven numbers
        If LCase(Left$(txtLine, 7)) = "values=" Then
            nums = Split(Mid$(txtLine, 8), " ")
            If UBound(nums) + 1 >= 7 Then
                For i = 1 To 7
                    Me.Controls("txtValue" & i).Value = nums(UBound(nums) - 7 + i)
                Next i
            End If
        End If
    Loop

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub


