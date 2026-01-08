Attribute VB_Name = "db_Dishes"
Attribute VB_Base = "0{BB0732C0-8C63-40C7-846F-84797535AED5}{2AF2E0C0-DD19-499B-A58B-0E5E9571AF74}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False




'———————————————————————————
' Module-level variables
Private Const FolderStoreSheet As String = "Meta"
Private Const FolderStoreCell  As String = "Filepath_DISH"
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

' On form load, try to read the saved folder and auto-populate
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
    
    '    ' 1) read last folder (default = "")
'    mFolderPath = GetSetting("ArcImporter", "Settings", "FolderPath", "")
'
'    ' 2) if it exists on disk, populate files; otherwise clear path
'    If Len(mFolderPath) > 0 Then
'        If Dir(mFolderPath, vbDirectory) <> "" Then
'            PopulateFileList
'            cmdBrowseFolder.Caption = "Change Folder…"
'        Else
'            mFolderPath = ""
'        End If
'    End If
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
    
    Me.txtVal1.list = Array("Dish", "Radome", "Shroud", "Grid", "Conical Horn", "Passive Reflector")
    Me.AType.list = Array("Dish", "Generic")
    Me.ATable.list = Array("Proposed", "Existing", "Removed")
    Me.AAttach.list = Array(1, 2)
    With shGeometry
        Me.AMembersA.list = .Range("E6:E" & WorksheetFunction.Max(.Range("G:G")) + 6).Value
        Me.AMembersB.list = .Range("E6:E" & WorksheetFunction.Max(.Range("G:G")) + 6).Value
        Me.AMembersC.list = .Range("E6:E" & WorksheetFunction.Max(.Range("G:G")) + 6).Value
        Me.AMembersD.list = .Range("E6:E" & WorksheetFunction.Max(.Range("G:G")) + 6).Value
    End With
'    Me.txtVal4.Value = 0
'    Me.txtVal5.Value = 0
'    Me.txtVal1.Value = "Dish"
'    Me.ASectors.Value = 3
'    Me.AThetaa.Value = 0
'    Me.AThetab.Value = 120
'    Me.AThetac.Value = 240
'    Me.AThetad.Value = 0
End Sub

' Browse/Change folder button
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
Dim nMembers As String

nMembers = WorksheetFunction.Max(Sheets("Geometry").Range("G:G")) + 7

With Sheets("Discrete Loads")
    For r = 4 To 53
    If txtVal4 = 0 Then
        Answer = MsgBox("The selected dish does not appear to have dimensions. Please select a dish with a Diameter.", vbOKCancel)
        If vbCancel Then
            Exit Sub
        End If
    End If

    If .Cells(r, 1).Value = "" Then
        .Cells(r, 1).Value = cmbFiles.Value
        .Cells(r, 2).Value = cmbUSNames.Value
        .Cells(r, 3).Value = ATable.Value
        .Cells(r, 4).Value = "Dish"
        If txtVal1 = 5 Then
            .Cells(r, 5) = "Passive Reflector"
        ElseIf txtVal1 = 4 Then
            .Cells(r, 5) = "Conical Horn"
        ElseIf txtVal1 = 3 Then
            .Cells(r, 5) = "Grid"
        ElseIf txtVal1 = 2 Then
            .Cells(r, 45) = "Shroud"  '"Parabloloid w/ Shroud (HP)"
        ElseIf txtVal1 = 1 Then
            .Cells(r, 5) = "Radome"  '"Parabloloid w/ Radome"
        ElseIf txtVal1 = 0 Then
            .Cells(r, 5) = "Dish"    '"Parabloloid w/o Radome"
        Else
            .Cells(r, 5) = txtVal1.Value
        End If
        .Cells(r, 6).Value = 1
        .Cells(r, 7).Value = AElevation.Value
        .Cells(r, 8).Value = 0
        .Cells(r, 9).Value = 8
        .Cells(r, 10).Value = 1
        .Cells(r, 11).Value = 1
        
        .Cells(r, 44).Formula2 = txtVal4.Value
        .Cells(r, 45).Formula2 = txtVal4.Value
        .Cells(r, 46).Formula2 = txtVal4.Value / 2
        .Cells(r, 47).Formula2 = txtVal5.Value
    
            If AMembersA <> "" Then
                .Cells(r, 12) = AThetaa.Value
                .Cells(r, 13) = AQtya.Value
                .Cells(r, 14) = AMembersA
                .Cells(r, 15) = ATopA.Value
'                .Cells(r, 16) = AMembersA.Value
'                .Cells(r, 17) = ABottomA.Value
                .Cells(r, 18) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(O" & r & "/(XLOOKUP(IF(ISNUMBER(N" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "N" & r & "," & Chr(34) & "$B" & Chr(34) & "), N" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
                .Cells(r, 19) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",Q" & r & "/(XLOOKUP(IF(ISNUMBER(P" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "P" & r & "," & Chr(34) & "$B" & Chr(34) & "), P" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
            End If
            
            If AMembersB <> "" Then
                .Cells(r, 20) = AThetab.Value
                .Cells(r, 21) = AQtyb.Value
                .Cells(r, 22) = AMembersB
                .Cells(r, 23) = ATopB.Value
'                .Cells(r, 24) = AMembersB
'                .Cells(r, 25) = ABottomB.Value
                .Cells(r, 26) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(W" & r & "/(XLOOKUP(IF(ISNUMBER(V" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "V" & r & "," & Chr(34) & "$B" & Chr(34) & "), V" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
                .Cells(r, 27) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",Y" & r & "/(XLOOKUP(IF(ISNUMBER(X" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "X" & r & "," & Chr(34) & "$B" & Chr(34) & "), X" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
            End If
            
            If AMembersC <> "" Then
                .Cells(r, 28) = AThetac.Value
                .Cells(r, 29) = AQtyc.Value
                .Cells(r, 30) = AMembersC
                .Cells(r, 31) = ATopC.Value
'                .Cells(r, 32) = AMembersC
'                .Cells(r, 33) = ABottomC.Value
                .Cells(r, 34) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(AE" & r & "/(XLOOKUP(IF(ISNUMBER(AD" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AD" & r & "," & Chr(34) & "$B" & Chr(34) & "), AD" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
                .Cells(r, 35) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",AG" & r & "/(XLOOKUP(IF(ISNUMBER(AF" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AF" & r & "," & Chr(34) & "$B" & Chr(34) & "), AF" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
            End If
            
            If AMembersD <> "" Then
                .Cells(r, 36) = AThetad.Value
                .Cells(r, 37) = AQtyd.Value
                .Cells(r, 38) = AMembersD
                .Cells(r, 39) = ATopD.Value
'                .Cells(r, 40) = AMembersD
'                .Cells(r, 42) = ABottomD.Value
                .Cells(r, 42) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(AM" & r & "/(XLOOKUP(IF(ISNUMBER(AL" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AL" & r & "," & Chr(34) & "$B" & Chr(34) & "), AL" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
                .Cells(r, 43) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",AO" & r & "/(XLOOKUP(IF(ISNUMBER(AN" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AN" & r & "," & Chr(34) & "$B" & Chr(34) & "), AN" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
            End If
                                            
    Exit Sub
    End If
    Next r
End With
End Sub


Private Sub cmdAddEntry_Click()
    Dim fso        As Object
    Dim ts         As Object
    Dim i          As Long
    Dim zeros      As String
    Dim valuesLine As String
    Dim filePath   As String
    Dim dishtype   As String

    ' -- Make sure we have a folder, file and USName selected --
    If mFolderPath = "" Or cmbFiles.Value = "" Or cmbUSNames.Value = "" Then
        MsgBox "Please select a folder, file and USName first.", vbExclamation
        Exit Sub
    End If

    filePath = mFolderPath & "\" & cmbFiles.Value & ".arc"

'    ' -- Build the 12 zeros string --
'    zeros = ""
'    For i = 2 To 4
'        zeros = zeros & "0 "
'    Next i
'
    If txtVal1 = "Dish" Then
        dishtype = "0 "
    ElseIf txtVal1 = "Radome" Then
        dishtype = "1 "
    ElseIf txtVal1 = "Shroud" Then
        dishtype = "2 "
    ElseIf txtVal1 = "Grid" Then
        dishtype = "3 "
    ElseIf txtVal1 = "Conical Horn" Then
        dishtype = "4 "
    ElseIf txtVal1 = "Passive Reflector" Then
        dishtype = "5 "
    End If
    
    ' -- Build the Values= line --
'    valuesLine = "Values=" & zeros
'    For i = 4 To 5
    valuesLine = "Values=" & "0 0 0 " & Me.Controls("txtVal" & 4).Value & " "
    valuesLine = valuesLine & Me.Controls("txtVal" & 5).Value & " " & "0 0 0 0 0 0 0 "
'    Next i
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
    db_Dishes.Show
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
'    For i = 4 To 5
'        Me.Controls("txtVal" & i).Value = ""
'    Next i
    
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
    
    ' 1st value ' Dish Type
    Me.txtVal1.Value = val(nums(0))
    
    ' 2nd value / 144 ' CaAa Front 0" Ice
    Me.txtVal2.Value = val(nums(2)) / 144
    
    ' 4th value ' Diameter inches
    Me.txtVal4.Value = val(nums(3))
    
    ' 5th value ' Weight 0" Ice
    Me.txtVal5.Value = val(nums(4))
    
    ' 6th value ' Weight 1/2" Ice
    Me.txtVal6.Value = val(nums(5))
    
    ' 10th value ' Weight 2" Ice
    Me.txtVal10.Value = val(nums(9))
    
    ' 11th value ' Weight 4" Ice
    Me.txtVal11.Value = val(nums(10))
    
    ' 12th value ' Weight 4" Ice
    Me.txtVal12.Value = val(nums(11))
    
'    ' last 7 values
'    For i = 1 To 7
'        ' UBound(nums)-7+1 is the first of the last 7
'        Me.Controls("txtValue" & i).Value = nums(UBound(nums) - 7 + i)
'    Next i
'
    ' 4/7 values
    If txtVal1 = 5 Then
        Me.Controls("txtVal" & 1).Value = "Passive Reflector"
    ElseIf txtVal1 = 4 Then
        Me.Controls("txtVal" & 1).Value = "Conical Horn"
    ElseIf txtVal1 = 3 Then
        Me.Controls("txtVal" & 1).Value = "Grid"
    ElseIf txtVal1 = 2 Then
        Me.Controls("txtVal" & 1).Value = "Shroud"
    ElseIf txtVal1 = 1 Then
        Me.Controls("txtVal" & 1).Value = "Radome"
    ElseIf txtVal1 = 0 Then
        Me.Controls("txtVal" & 1).Value = "Dish"
    Else
        Me.Controls("txtVal" & 1).Value = txtVal1.Value
    End If
    
End Sub


Private Sub UserForm_Click()

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
    For i = 1 To 12
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




