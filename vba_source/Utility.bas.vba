Attribute VB_Name = "Utility"
'Option Explicit
''This module contains all of the procedures to copy
''point/distibuted loads along with BLCs and LCs.


Sub Updates_Off()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
End Sub


Sub Updates_On()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
End Sub


Sub Append_Locations()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False
    
Dim r As Long
Dim nMembers As String
nMembers = WorksheetFunction.Max(Sheets("Geometry").Range("G:G")) + 7
With Sheets("Discrete Loads")
    For r = 4 To 3 + WorksheetFunction.CountA(.Range("A4:A53"))
        .Cells(r, 18).Formula2 = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(O" & r & "/(XLOOKUP(IF(ISNUMBER(N" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "N" & r & "," & Chr(34) & "$B" & Chr(34) & "), N" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
        .Cells(r, 19).Formula2 = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",Q" & r & "/(XLOOKUP(IF(ISNUMBER(P" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "P" & r & "," & Chr(34) & "$B" & Chr(34) & "), P" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
        .Cells(r, 26).Formula2 = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(W" & r & "/(XLOOKUP(IF(ISNUMBER(V" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "V" & r & "," & Chr(34) & "$B" & Chr(34) & "), V" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
        .Cells(r, 27).Formula2 = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",Y" & r & "/(XLOOKUP(IF(ISNUMBER(X" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "X" & r & "," & Chr(34) & "$B" & Chr(34) & "), X" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
        .Cells(r, 34).Formula2 = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(AE" & r & "/(XLOOKUP(IF(ISNUMBER(AD" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AD" & r & "," & Chr(34) & "$B" & Chr(34) & "), AD" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
        .Cells(r, 35).Formula2 = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",AG" & r & "/(XLOOKUP(IF(ISNUMBER(AF" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AF" & r & "," & Chr(34) & "$B" & Chr(34) & "), AF" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
        .Cells(r, 42).Formula2 = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(AM" & r & "/(XLOOKUP(IF(ISNUMBER(AL" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AL" & r & "," & Chr(34) & "$B" & Chr(34) & "), AL" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
        .Cells(r, 43).Formula2 = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",AO" & r & "/(XLOOKUP(IF(ISNUMBER(AN" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AN" & r & "," & Chr(34) & "$B" & Chr(34) & "), AN" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
    Next r
End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
End Sub


Sub Append_DiscreteDimensions()
    Call Fill_Dimensions_From_Arc
End Sub


Function LastAuthor()
    LastAuthor = ActiveWorkbook.BuiltinDocumentProperties("Last Author")
End Function


Sub cmdAddModel()
    Dim fso        As Object
    Dim ts         As Object
    Dim i          As Long
    Dim zeros      As String
    Dim valuesLine As String
    Dim filePath   As String
    
    ' -- Make sure we have a folder, file and USName selected --
    Set mFolderPath = Sheets("Meta").Range("Filepath_APPURT")
    If mFolderPath = "" Then 'Or cmbFiles.Value = "" Or cmbUSNames.Value = "" Then
        MsgBox "Please select a folder, file and USName first.", vbExclamation
        Exit Sub
    End If
    
    filePath = mFolderPath & "\" & "nb+c_mat_tower mounts.arc"
    
    ' -- Build the 15 zeros string --
'    zeros = ""
'    For i = 1 To 15
'        zeros = zeros & "0 "
'    Next i
    
    ' -- Build the Values= line --
'    valuesLine = "Values=" & zeros
'    For i = 1 To 7
'        valuesLine = valuesLine & WorksheetFunction.Substitute(WorksheetFunction.Substitute(WorksheetFunction.Substitute(Me.Controls("txtValue" & i).Value, "Flat", 1), "Round", 3), "Generic", 1) & " "
'    Next i
'    valuesLine = Trim(valuesLine)
    
    ' -- Open for append and write the three lines --
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' 8 = ForAppending, True = create file if it doesn't exist
    Set ts = fso.OpenTextFile(filePath, 8, True)
    
    ts.WriteLine          ' blank line before new entry
    ts.WriteLine Sheets("Meta").Range("E19")
    ts.WriteLine Sheets("Meta").Range("E21")
    ts.WriteLine Sheets("Meta").Range("E22")
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    MsgBox "New entry added to " & "nb+c_mat_tower mounts.arc", vbInformation
    
'    Unload Me
'    db_DiscreteLoads.Show
End Sub


Sub DeleteProjectData()
Dim F0 As Range, f1 As Range, f2 As Range
With Sheets("Risa File")
    Set F0 = .Columns("A").Find(What:="[GLOBAL_PARAMETERS]", LookIn:=xlValues, LookAt:=xlWhole)
    Set f1 = .Columns("A").Find(What:="[.PROJECT_DESCRIPTION]", LookIn:=xlValues, LookAt:=xlWhole)
    Set f2 = .Columns("A").Find(What:="[.END_PROJECT_DESCRIPTION]", LookIn:=xlValues, LookAt:=xlWhole)
    If Not f1 Is Nothing And Not f2 Is Nothing Then .Rows(f1.row & ":" & f2.row).Delete
'    If F1 Is Nothing And F2 Is Nothing Then
'    F0.Select
    F0.offset(1, 0).EntireRow.Resize(21).Insert
        F0.offset(2, 0) = "[.PROJECT_DESCRIPTION]"
        
        F0.offset(4, 0) = "[..MODEL_TITLE] <1>"
        F0.offset(5, 0) = " "
        F0.offset(6, 0) = "[..END_MODEL_TITLE]"
        
        F0.offset(8, 0) = "[..COMPANY_NAME] <1>"
        If Left(Sheets("Code").Range("B2"), 3) = "TKK" Then
            F0.offset(9, 0) = "TKK Engineering"
        Else
            F0.offset(9, 0) = "NB+C ES" 'Range("LegalEntity")
        End If
        F0.offset(10, 0) = "[..END_COMPANY_NAME]"
        
        F0.offset(12, 0) = "[..DESIGNER_NAME] <1>"
        F0.offset(13, 0) = "NB+C Engineer" 'LastAuthor
        F0.offset(14, 0) = "[..END_DESIGNER_NAME]"
        
        F0.offset(16, 0) = "[..JOB_NUMBER] <1>"
        F0.offset(17, 0) = " "
        F0.offset(18, 0) = "[..END_JOB_NUMBER]"
    
        F0.offset(20, 0) = "[.END_PROJECT_DESCRIPTION]"
End With
End Sub

Sub Save()
    ActiveWorkbook.Save
End Sub

Sub Clear_CodeTab()
Dim lastRow As Long

    With shCode
        .Range("C2:D7").ClearContents
        .Range("C16:D19").ClearContents
        .Range("C21:D22").ClearContents
        .Range("C24:D28").ClearContents
        .Range("C30:D31").ClearContents
        .Range("C33:D35").ClearContents
'        .Range("C41:D44").ClearContents
        .Range("H16:I21").ClearContents
        .Range("H23:I25").ClearContents
        .Range("H27:I32").ClearContents
    End With
End Sub

Sub Clear_GeometryTab()
Dim lastRow As Long

    With shGeometry
    lastRow = .Range("R6").End(xlDown).row
        .Range("B6:V" & lastRow).ClearContents
        .Range("X6:AA31").ClearContents
        .Range("AD6:AF31").ClearContents
    End With
End Sub

Sub Clear_MaintenanceTab()
Dim lastRow As Long

    With shMaintenance_Loads
        .Range("B7:B14").ClearContents
        .Range("E7:E14").ClearContents
        .Range("H7:K10").ClearContents
    End With
End Sub

Sub Clear_PointLoadTablesTab()
Dim lastRow As Long

    With shPointLoadTables
    lastRow = .Range("A4").End(xlDown).row
        .Range("A4:HP" & lastRow).ClearContents
    End With
End Sub

Sub Clear_DistributedLoadTablesTab()
Dim lastRow As Long

    With shDistributedLoadTables
    lastRow = .Range("A4").End(xlDown).row
        .Range("A4:HV99999").ClearContents
    End With
End Sub

Sub Clear_AreaLoadTablesTab()
Dim lastRow As Long

    With shAreaLoadTables
    lastRow = .Range("A4").End(xlDown).row
        .Range("A4:O99999").ClearContents
    End With
End Sub

Sub Clear_CodeCheck_BatchTab()
Dim lastRow As Long

    With CodeCheck_Batch
    lastRow = .Range("A3").End(xlDown).row
        .Range("A4:S99999").ClearContents
        .Range("T4:T6").ClearContents
    End With
End Sub

Sub Clear_NodeReactions_BatchTab()
Dim lastRow As Long

    With NodeReactions_Batch
    lastRow = .Range("A4").End(xlDown).row
        .Range("A4:I99999").ClearContents
    End With
End Sub

Sub Clear_MemberEndRxns_BatchTab()
Dim lastRow As Long

    With MemberEndRxns_Batch
    lastRow = .Range("A4").End(xlDown).row
        .Range("A4:J99999").ClearContents
    End With
End Sub

Sub Clear_ResultsTab()
Dim lastRow As Long

    With shResults
    lastRow = .Range("A3").End(xlDown).row
        .Range("A3:R99999").ClearContents
        .Range("T3:T5").ClearContents
    End With
End Sub

Sub Clear_RisaFileTab()
Dim lastRow As Long

    With Sheets("Risa File")
        .Range("A1:DZ99999").ClearContents
    End With
End Sub

Sub Clear_AllTabs()

    Call Clear_CodeTab
    Call Clear_GeometryTab
    Call Clear_Discrete_Loads
    Call Clear_Dish_Loads
    Call Clear_MaintenanceTab
    Call Clear_PointLoadTablesTab
    Call Clear_DistributedLoadTablesTab
    Call Clear_AreaLoadTablesTab
    Call Clear_CodeCheck_BatchTab
    Call Clear_NodeReactions_BatchTab
    Call Clear_MemberEndRxns_BatchTab
    Call Clear_RisaFileTab

End Sub

Sub Clear_Discrete_Loads()
Dim lastRow As Long

    With shDiscrete_Loads
        .Range("A4:CG99999").ClearContents
    End With
    
End Sub

Sub Clear_Dish_Loads()
Dim lastRow As Long

    With shDish_Loads
        .Range("A4:CB99999").ClearContents
    End With

End Sub

Sub Activate_CodeTab()
    shCode.Activate
End Sub

Sub Activate_GeometryTab()
    shGeometry.Activate
End Sub

Sub Activate_DiscreteLoadsTab()
    shDiscrete_Loads.Activate
End Sub

Sub Activate_DishLoadsTab()
    shDish_Loads.Activate
End Sub

Sub Activate_MaintenanceLoadsTab()
    shMaintenance_Loads.Activate
End Sub

Sub Activate_RISA3DTab()
    shRISA_3D.Activate
End Sub

Sub Activate_ResultsTab()
    shResults.Activate
End Sub

Sub Activate_PlacementDiagramsTab()
    shPlacement_Diagrams.Activate
End Sub

Sub ShowHide_LoadTables()
    If shPointLoadTables.Visible = True And shDistributedLoadTables.Visible = True And shAreaLoadTables.Visible = True Then
        Sheets("Point Load Tables").Visible = False
        Sheets("Distributed Load Tables").Visible = False
        Sheets("Area Load Tables").Visible = False
    ElseIf shPointLoadTables.Visible = False And shDistributedLoadTables.Visible = False And shAreaLoadTables.Visible = False Then
        Sheets("Point Load Tables").Visible = True
        Sheets("Distributed Load Tables").Visible = True
        Sheets("Area Load Tables").Visible = True
    End If

End Sub

Sub Refresh_Query()
    Sheets("Database_APPURT").Select
    ActiveWorkbook.Connections("Query - Database_APPURT").Refresh
    Sheets("EPA Input").Select
'    Application.CalculateUntilAsyncQueriesDone
End Sub

Sub Refresh_Local()

On Error Resume Next

Application.DisplayAlerts = False
Application.ScreenUpdating = False

    With ThisWorkbook
        For Each objConnection In .Connections
            'Get current background-refresh value
            bBackground = objConnection.OLEDBConnection.BackgroundQuery
            'Temporarily disable background-refresh
            objConnection.OLEDBConnection.BackgroundQuery = False
            'Refresh this connection
            objConnection.Refresh
            'Set background-refresh value back to original value
            objConnection.OLEDBConnection.BackgroundQuery = bBackground
        Next
        'Save the updated Data
        .Save
    End With

Application.DisplayAlerts = True
Application.ScreenUpdating = True

'MsgBox "Local appurtenance database has been updated."
End Sub

Sub Get_Connection_Names()
    With ThisWorkbook
     'Check if there is any conection
        If .Connections.Count = 0 Then Exit Sub
     'Print the numer of conection (item number) and its name
        For X = 1 To .Connections.Count
            Debug.Print X & ": " & .Connections.Item(X).Name
        Next X
    End With
End Sub

Sub UpdateConnectionbyName()
    Dim ConectionName As String
    ConectionName = "Query - Database_APPURT"
    With ThisWorkbook
     'Check if there is any conection
        If .Connections.Count = 0 Then Exit Sub
     'Check the connections names if one macht with the one we want.
        For X = 1 To .Connections.Count
            If .Connections.Item(X).Name = ConectionName Then .Connections.Item(X).Refresh
        Next X
    End With
End Sub

Sub Refresh_Server()
    Dim wb As Workbook
    Dim myfilename As String
    Dim Answer As Integer
    Dim PauseTime, Start, Finish, TotalTime
    
    myfilename = "Z:\~Engineering Standards Library\STRUCTURAL\Analysis_Tower\tnxTower\Database_APPURT.xlsm"
    Set wb = Workbooks.Open(myfilename)
    ActiveWorkbook.Connections("Connection").Refresh
    Application.CalculateUntilAsyncQueriesDone
    wb.Save
    wb.Close
    
    Answer = MsgBox("Server and local appurtenance databases have been updated.", vbOKOnly + vbDefaultButton2, "Databases Updated!")
     
End Sub

Sub Refresh_Database()
    Call Refresh_Server
    Call Refresh_Local
End Sub

Sub OpenLoadingProfiles()
    
Workbooks.Open "C:\Users\mgirgis\Documents\~Development\Mount Analysis Tool v3.0.0\Loading Profiles.xlsx"

End Sub

Sub CloseLoadingProfiles()

Workbooks("Loading Profiles.xlsx").Close SaveChanges:=True

End Sub

Sub SaveLoading()
Dim r As Long
Dim lastRow As Long
    
    Range("A4:BQ4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    shLoading_Profiles.Activate
    If Range("D4") <> "" Then
        Range("D4").Select
        Selection.End(xlDown).Select
        Selection.offset(1, 0).Select
        ActiveSheet.Paste
        shDiscrete_Loads.Activate
    Else
        Range("D4").Select
        ActiveSheet.Paste
        shDiscrete_Loads.Activate
        Application.CutCopyMode = False
        Range("A4").Select
    End If
    
'    ThisWorkbook.Sheets("Discrete Loads").Rows("4:4").Select
'    ThisWorkbook.shDiscrete_Loads.Rows("4:4").Select
'    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy Workbooks("Loading Profiles.xlsx").Worksheets("Sheet1").Range("D4")
    
End Sub

Sub LoadingProfile_Export()

Dim r As Long
Dim lastRow As Long
    
    ThisWorkbook.Sheets("Discrete Loads").Range("A4:BH4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks.Open "C:\Users\mgirgis\Documents\~Development\Mount Analysis Tool v3.0.0\Loading Profiles.xlsx"

    If Range("D4") <> "" Then
        Range("D4").Select
        Selection.End(xlDown).Select
        Selection.offset(1, 0).Select
        ActiveSheet.Paste
        Workbooks("Loading Profiles.xlsx").Worksheets("Sheet1").Range("A4").Select
        For r = 4 To 503
        If Range("D" & r) <> "" And Range("A" & r) = "" And Range("B" & r) = "" And Range("C" & r) = "" Then
            Range("A" & r) = ThisWorkbook.Sheets("Code").Range("Carrier").Value
            Range("B" & r) = ThisWorkbook.Sheets("Code").Range("Carrier_Project").Value
            Range("C" & r) = Format(Date, "YYYY-MM-DD") & "_" & Format(Now, "hh-mm-ss")
        Else
        End If
        Next r
        shDiscrete_Loads.Activate
    Else
        Range("D4").End(xlDown).Select
        ActiveSheet.Paste
        shDiscrete_Loads.Activate
        Application.CutCopyMode = False
        Workbooks("Loading Profiles.xlsx").Worksheets("Sheet1").Range("A4").Select
    End If
    
    Selection.Copy Workbooks("Loading Profiles.xlsx").Worksheets("Sheet1").Range("D4")
    
    Workbooks("Loading Profiles.xlsx").Close SaveChanges:=True
    
    shDiscrete_Loads.Range("A4").Select

End Sub

Sub ImportPreviousMA()
    Dim fileName As Variant
    Dim Results_Header As Range               'Results
    Dim Results() As String                   'Results
    Dim Results_Section() As String
    Dim R_Section() As String
    Dim r As Long, n As Long, t As Byte
    Dim a As Byte, b As Byte, c As Byte
    Dim m As Double, lastRow As Double
    Dim v As Integer
    Dim Temp As Worksheet               'Tempoaray file imported from RISA 3D
    Dim Openbook As Workbook
    Dim Project As String, Folder As String
    Dim cnt As Integer
    Dim Results_Count As Integer
    Dim Answer As String
    Dim FlatFile As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    
OpenFile:
    fileName = Application.GetOpenFilename("NB+C MAT v3.0+ (*.xlsm),*.xlsm")
        If fileName = False Then    'If user hits "Cancel" button
            Exit Sub
        End If
    Application.Workbooks.OpenText fileName, dataType:=xlDelimited, Space:=False, Comma:=False, Semicolon:=False
    Set Openbook = ActiveWorkbook
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("C2:D2").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("C2:D2").PasteSpecial (xlPasteValues)
    
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("C4:D6").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("C4:D6").PasteSpecial (xlPasteValues)
    
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("C16:D18").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("C16:D18").PasteSpecial (xlPasteValues)
    
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("C19:D19").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("C19:D19").PasteSpecial (xlPasteValues)
    
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("C21:D28").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("C21:D28").PasteSpecial (xlPasteValues)
    
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("C30:D31").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("C30:D31").PasteSpecial (xlPasteValues)
    
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("C33:D35").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("C33:D35").PasteSpecial (xlPasteValues)
    
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("C41:D44").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("C41:D44").PasteSpecial (xlPasteValues)
    
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("H16:I21").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("H16:I21").PasteSpecial (xlPasteValues)
    
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("H23:I25").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("H23:I25").PasteSpecial (xlPasteValues)
    
        Openbook.Sheets("Code").Activate
        Openbook.Sheets("Code").Range("H27:H32").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Code").Range("H27:H32").PasteSpecial (xlPasteValues)
        
        Openbook.Sheets("Geometry").Activate
        Openbook.Sheets("Geometry").Range("A6:V99999").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Geometry").Range("A6:V99999").PasteSpecial (xlPasteValues)
        
        Openbook.Sheets("Geometry").Activate
        Openbook.Sheets("Geometry").Range("X6:AA31").Copy
        ThisWorkbook.Activate
        ThisWorkbook.Sheets("Geometry").Range("X6:AA31").PasteSpecial (xlPasteValues)
        
        ThisWorkbook.Sheets("Risa File").Select
        ThisWorkbook.Sheets("Risa File").Cells.Select
        Selection.ClearContents
        Openbook.Sheets("Risa File").Activate
        Openbook.Sheets("Risa File").Range("A1").Select
        Openbook.Sheets("Risa File").Cells.Select
        Openbook.Sheets("Risa File").Cells.Copy
        ThisWorkbook.Sheets("Risa File").Range("A1").PasteSpecial (xlPasteValues)
        ThisWorkbook.Sheets("Risa File").Range("A1").Select
        
        ThisWorkbook.Sheets("Code").Activate
        ThisWorkbook.Sheets("Code").Range("C2").Select
        
    Openbook.Close
    ThisWorkbook.Sheets("Code").Range("C2").Select
    MsgBox "Previous MAT v3+ import is completed! Please update the Discrete Loads and Dish Loads tabs as needed.", vbOKOnly, "MAT v3+ Imported"
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


Function CleanFileName(s As String) As String
    Dim invalidChars As Variant
    Dim c As Variant
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each c In invalidChars
        s = Replace(s, c, "_")
    Next c
    CleanFileName = s
End Function


Sub Open_RFDS_SQL_Database()

    Dim url As String
    url = "https://agreeable-rock-0a9d2580f.2.azurestaticapps.net/"

    ThisWorkbook.FollowHyperlink _
        Address:=url, _
        NewWindow:=True

End Sub


