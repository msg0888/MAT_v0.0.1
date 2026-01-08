Attribute VB_Name = "Import_ProjectData"
Option Explicit

Dim fso As Scripting.FileSystemObject
Dim Project_Data As Worksheet
Dim fileName As Variant
Dim Openbook As Workbook
Dim Answer As String

Sub ProjectData()
On Error Resume Next
frm_ProjectData.Show

With shCode
'    .Range("Client") = WorksheetFunction.VLookup("Client", Sheets("Project_Data").Range("B1:D100"), 3, False)
    .Range("Carrier_SiteName") = WorksheetFunction.VLookup(Range("Carrier") & " Site Name", Sheets("Project_Data").Range("B1:D100"), 3, False)
    .Range("Carrier_SiteNumber") = WorksheetFunction.VLookup(Range("Carrier") & " Site ID", Sheets("Project_Data").Range("B1:D100"), 3, False)
    .Range("z") = WorksheetFunction.VLookup("Antenna Rad Center", Sheets("Project_Data").Range("B1:D100"), 3, False)
    .Range("TowerHeight") = WorksheetFunction.VLookup("Overall Tower Height (feet)", Sheets("Project_Data").Range("B1:D100"), 3, False)
    .Range("TowerType") = WorksheetFunction.Substitute(WorksheetFunction.VLookup("Structure Type", Sheets("Project_Data").Range("B1:D100"), 3, False), " Tower", "")
End With

'Call Import
Sheets("Project_Data").Visible = False
shCode.Activate
shCode.Range("ProjectNumber").Select

End Sub

Sub Select_File()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next
 
Sheets("Project_Data").Range("A1:F99999").ClearContents
 
Set fso = New FileSystemObject
Set Project_Data = Sheets("Project_Data")
On Error GoTo 0

If Not Project_Data Is Nothing Then
    Sheets("Project_Data").Delete  'if user selects "yes', this deletes the previously used file
    GoTo OpenFile
End If

OpenFile:
fileName = Application.GetOpenFilename("AutoCAD (*.prp),*.prp")
    If fileName = False Then    'If user hits "Cancel" button
        Exit Sub
    End If
Application.Workbooks.OpenText fileName, dataType:=xlDelimited, Space:=False

'    With FileName
'        FileSelected = .SelectedItems(1)
'    End With
Meta.Range("ProjectFilepath") = fso.GetAbsolutePathName(fileName)
Set Openbook = ActiveWorkbook
Openbook.Sheets(1).Cells.Copy
ThisWorkbook.Activate
ThisWorkbook.Sheets.Add After:=ThisWorkbook.Worksheets(Sheets.Count)
ActiveSheet.Name = "Project_Data"
ActiveSheet.Range("A1").PasteSpecial (xlPasteValues)
Range("A1").Select
ActiveSheet.Visible = True
Application.CutCopyMode = False
Openbook.Close False

Call Import_Data
Application.ScreenUpdating = True

End Sub

Sub Import_Data()

    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Project_Data").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), dataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=False, Tab:=True, Semicolon _
        :=False, Comma:=False, Space:=False, Other:=True, OtherChar:="""", _
        FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)), _
        TrailingMinusNumbers:=True
    Selection.TextToColumns Destination:=Range("A1"), dataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=False, Tab:=True, Semicolon _
        :=False, Comma:=False, Space:=False, Other:=False, OtherChar:="""", _
        FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Columns("D:D").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("A:E").EntireColumn.AutoFit
    Range("A1").Select
End Sub
Sub Macro4()
'
' Macro4 Macro
'

'
    Columns("B:D").Select
    Columns("B:D").EntireColumn.AutoFit
End Sub


