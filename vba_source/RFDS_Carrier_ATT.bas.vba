Attribute VB_Name = "RFDS_Carrier_ATT"
Option Explicit

' ======================================================================
' IMPORT FROM RF_Inventory INTO Discrete Loads
'
' Changes vs APD version:
'   - Source sheet: "RF_Inventory" OR "RF Inventory"
'   - Condition source: "Status" header (instead of APD_Section)
'   - Sector placement comes from "Sec-Pos":
'       * Sector = text before "-" (A-1 -> A)
'       * Q treated as A
'       * If Sec-Pos blank or "-" => omit row
'   - Type:
'       * If Equipment Type = "ANTENNA" => "Antenna"
'       * Else => "TME"
'   - Sec-Pos written to:
'       * A/Q => N and P
'       * B   => V and X
'       * C   => AD and AF
'       * D   => AL and AN
' ======================================================================
Public Sub ImportRFDS_ATT()

    Dim fd As FileDialog
    Dim srcPath As String
    Dim wbSrc As Workbook, wsSrc As Worksheet, wsDest As Worksheet

    Dim az(1 To 4) As Variant  ' A,B,C,D azimuths

    '======== Select Source Workbook ========
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "Select AT&T Mobility Converted RFDS"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
    If fd.Show <> -1 Then Exit Sub
    srcPath = fd.SelectedItems(1)

'    ' Ask once up-front
'    If Not PromptSectorAzimuths(az) Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wbSrc = Workbooks.Open(srcPath)

    ' Try both names
    Set wsSrc = Nothing
    On Error Resume Next
    Set wsSrc = wbSrc.Sheets("RF_Inventory")
    If wsSrc Is Nothing Then Set wsSrc = wbSrc.Sheets("RF Inventory")
    On Error GoTo 0

    If wsSrc Is Nothing Then
        MsgBox "Sheet 'RF_Inventory' or 'RF Inventory' not found.", vbCritical
        wbSrc.Close False
        GoTo CleanExit
    End If

    Set wsDest = ThisWorkbook.Sheets("Discrete Loads")

    ' Process sheet
    Import_RFInv_Data wsSrc, wsDest, az

    wbSrc.Close False

    Call Fill_Dimensions_From_Arc
    Call FillInTheBlanks_ATT
    MsgBox "AT&T RF Inventory appurtenances imported into Discrete Loads.", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


' ======================================================================
' CORE IMPORT FUNCTION
' ======================================================================
Private Sub Import_RFInv_Data(wsSrc As Worksheet, wsDest As Worksheet, ByRef az() As Variant)

    Dim headerRow As Long
    Dim vendorCol As Long, modelCol As Long, statusCol As Long
    Dim secPosCol As Long, eqTypeCol As Long
    Dim lastRow As Long
    Dim i As Long

    Dim vendorVal As String, modelVal As String
    Dim statusVal As String, secPosVal As String, eqTypeVal As String
    Dim sectorLetter As String, secPosOut As String
    Dim groupKey As String

    Dim groups As Object              ' key -> Variant( vendor, model, status, condition, type, secA, secB, secC, secD )
    Dim rec As Variant
    Dim destRow As Long
    Dim k As Variant

    '======== Find header row ========
    headerRow = FindHeaderRowMulti(wsSrc, Array("Vendor", "Model", "Status", "Sec-Pos", "Equipment Type"))
    If headerRow = 0 Then
        MsgBox "Cannot find required headers: Vendor / Model / Status / Sec-Pos / Equipment Type", vbCritical
        Exit Sub
    End If

    vendorCol = FindHeader(wsSrc, "Vendor", headerRow)
    modelCol = FindHeader(wsSrc, "Model", headerRow)
    statusCol = FindHeader(wsSrc, "Status", headerRow)
    secPosCol = FindHeader(wsSrc, "Sec-Pos", headerRow)
    eqTypeCol = FindHeader(wsSrc, "Equipment Type", headerRow)

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, vendorCol).End(xlUp).row

    Set groups = CreateObject("Scripting.Dictionary")

    '======== BUILD GROUPS ========
    For i = headerRow + 1 To lastRow

        vendorVal = Trim$(CStr(wsSrc.Cells(i, vendorCol).Value))
        modelVal = Trim$(CStr(wsSrc.Cells(i, modelCol).Value))
        statusVal = Trim$(CStr(wsSrc.Cells(i, statusCol).Value))
        eqTypeVal = Trim$(CStr(wsSrc.Cells(i, eqTypeCol).Value))
        secPosVal = ExtractSecPosToken(Trim$(CStr(wsSrc.Cells(i, secPosCol).Value)), eqTypeVal)


        ' Must have at least vendor or model
        If vendorVal = "" And modelVal = "" Then GoTo NextI

        ' Skip if Sec-Pos blank or "-"
        If secPosVal = "" Or secPosVal = "-" Then GoTo NextI

        sectorLetter = GetSectorFromSecPos(secPosVal)
        If sectorLetter = "" Then GoTo NextI

        ' Treat Q as A
        If sectorLetter = "Q" Then sectorLetter = "A"

        secPosOut = secPosVal

        ' Group key includes STATUS (per your request)
        groupKey = UCase$(vendorVal) & "|" & UCase$(modelVal) & "|" & UCase$(statusVal)

        If Not groups.Exists(groupKey) Then
            ' rec:
            ' 0 vendor
            ' 1 model
            ' 2 status
            ' 3 condition (mapped)
            ' 4 type ("Antenna" or "TME")
            ' 5 secA
            ' 6 secB
            ' 7 secC
            ' 8 secD
            rec = Array( _
                vendorVal, _
                modelVal, _
                statusVal, _
                MapStatusToCondition(statusVal), _
                IIf(UCase$(eqTypeVal) = "ANTENNA", "Antenna", "TME"), _
                vbNullString, vbNullString, vbNullString, vbNullString _
            )
            groups.Add groupKey, rec
        Else
            rec = groups(groupKey)

            ' If any row in the group is ANTENNA, promote Type to "Antenna"
            If UCase$(eqTypeVal) = "ANTENNA" Then rec(4) = "Antenna"

            ' (Condition comes from Status already in key; no need to remap)
        End If

        ' Fill sector slot if empty (do NOT overwrite existing)
        Select Case sectorLetter
            Case "A"
                If Len(rec(5)) = 0 Then rec(5) = secPosOut
            Case "B"
                If Len(rec(6)) = 0 Then rec(6) = secPosOut
            Case "C"
                If Len(rec(7)) = 0 Then rec(7) = secPosOut
            Case "D"
                If Len(rec(8)) = 0 Then rec(8) = secPosOut
        End Select

        groups(groupKey) = rec

NextI:
    Next i

    '======== WRITE OUTPUT (one row per group) ========
    destRow = 4

    For Each k In groups.Keys
        rec = groups(k)

        wsDest.Cells(destRow, "A").Value = rec(0)          ' Vendor
        wsDest.Cells(destRow, "B").Value = rec(1)          ' Model
        wsDest.Cells(destRow, "C").Value = rec(3)          ' Condition
        wsDest.Cells(destRow, "D").Value = rec(4) '"Antenna" 'rec(4)          ' Type
        wsDest.Cells(destRow, "E").Value = "Flat"          ' Default shape
        Select Case rec(4)                                 ' Default attach
            Case "Antenna": wsDest.Cells(destRow, "F").Value = 2
            Case Else: wsDest.Cells(destRow, "F").Value = 1
        End Select
        wsDest.Cells(destRow, "G").Value = 0               ' Default Vertical Offset
            
'        ' G = Centerline - Code!H16 (Antenna Summary only)
'        If hasCenterline And centerlineCol > 0 Then
'            Dim clVal As Variant
'            clVal = wsSrc.Cells(i, centerlineCol).Value
'            If IsNumeric(clVal) Then
'                wsDest.Cells(destRow, "G").Value = CDbl(clVal) - ThisWorkbook.Worksheets("Code").Range("H16").Value
'            Else
'                wsDest.Cells(destRow, "G").Value = 0
'            End If
'        Else
'            wsDest.Cells(destRow, "G").Value = 0
'        End If
        
        wsDest.Cells(destRow, "H").Value = 0
        wsDest.Cells(destRow, "I").Value = 8
        wsDest.Cells(destRow, "J").Value = 1
        wsDest.Cells(destRow, "K").Value = 1
        
        

        ' Place Sec-Pos by sector (if present)
        If Len(rec(5)) > 0 Then PlaceSecPos wsDest, destRow, "A", rec(5), az
        If Len(rec(6)) > 0 Then PlaceSecPos wsDest, destRow, "B", rec(6), az
        If Len(rec(7)) > 0 Then PlaceSecPos wsDest, destRow, "C", rec(7), az
        If Len(rec(8)) > 0 Then PlaceSecPos wsDest, destRow, "D", rec(8), az

        destRow = destRow + 1
    Next k
End Sub


' ======================================================================
' Sector placement for Sec-Pos
' A/Q => N and P
' B   => V and X
' C   => AD and AF
' D   => AL and AN
' ======================================================================
Private Sub PlaceSecPos(ws As Worksheet, ByVal rowNum As Long, ByVal sectorLetter As String, ByVal secPos As String, ByRef az() As Variant)

    Select Case sectorLetter
        Case "A"
            ws.Cells(rowNum, "N").Value = secPos
            ws.Cells(rowNum, "P").Value = secPos
            ws.Cells(rowNum, "L").Value = IIf(IsNumeric(az(1)), az(1), vbNullString)
        Case "B"
            ws.Cells(rowNum, "V").Value = secPos
            ws.Cells(rowNum, "X").Value = secPos
            ws.Cells(rowNum, "T").Value = IIf(IsNumeric(az(2)), az(2), vbNullString)
        Case "C"
            ws.Cells(rowNum, "AD").Value = secPos
            ws.Cells(rowNum, "AF").Value = secPos
            ws.Cells(rowNum, "AB").Value = IIf(IsNumeric(az(3)), az(3), vbNullString)
        Case "D"
            ws.Cells(rowNum, "AL").Value = secPos
            ws.Cells(rowNum, "AN").Value = secPos
            ws.Cells(rowNum, "AJ").Value = IIf(IsNumeric(az(4)), az(4), vbNullString)
        Case Else
            ' do nothing
    End Select

End Sub


' ======================================================================
' Helpers
' ======================================================================
Private Function GetSectorFromSecPos(ByVal secPos As String) As String
    Dim s As String, p As Long
    s = UCase$(Trim$(secPos))

    If s = "" Or s = "-" Then
        GetSectorFromSecPos = ""
        Exit Function
    End If

    p = InStr(1, s, "-", vbTextCompare)
    If p > 1 Then
        GetSectorFromSecPos = Left$(s, p - 1) ' text before hyphen
    Else
        ' If no hyphen, fallback to first character
        GetSectorFromSecPos = Left$(s, 1)
    End If

    ' Normalize to single letter if they give something longer
    If Len(GetSectorFromSecPos) > 1 Then
        GetSectorFromSecPos = Left$(GetSectorFromSecPos, 1)
    End If

    ' Only accept A/B/C/D/Q
    If InStr(1, "ABCDQ", GetSectorFromSecPos, vbBinaryCompare) = 0 Then
        GetSectorFromSecPos = ""
    End If
End Function


Private Function MapStatusToCondition(ByVal statusVal As String) As String
    Select Case UCase$(Trim$(statusVal))
        Case "FINAL", "PROPOSED", "PROPOSED INSTALL"
            MapStatusToCondition = "Proposed"
        Case "EXISTING", "AS-BUILT"
            MapStatusToCondition = "Existing"
        Case "REMOVED", "REMOVE", "DECOMMISSION", "PROPOSED DECOMMISSION"
            MapStatusToCondition = "Removed"
        Case Else
            MapStatusToCondition = "Existing" ' or "" if you prefer
    End Select
End Function


' Finds a header row (1..20) that contains all required headers (partial match OK)
Private Function FindHeaderRowMulti(ws As Worksheet, headers As Variant) As Long
    Dim r As Long, h As Variant, f As Range, ok As Boolean

    For r = 1 To 20
        ok = True
        For Each h In headers
            Set f = ws.Rows(r).Find(CStr(h), , xlValues, xlPart)
            If f Is Nothing Then
                ok = False
                Exit For
            End If
        Next h

        If ok Then
            FindHeaderRowMulti = r
            Exit Function
        End If
    Next r
End Function


Private Function FindHeader(ws As Worksheet, headerText As String, headerRow As Long) As Long
    Dim c As Range
    Set c = ws.Rows(headerRow).Find(headerText, , xlValues, xlPart)
    If Not c Is Nothing Then FindHeader = c.Column
End Function


Private Function ExtractSecPosToken(ByVal secPosText As String, ByVal eqTypeText As String) As String
    Dim s As String, t As Variant, tokens As Variant
    
    s = Trim$(secPosText)
    If s = "" Or s = "-" Then s = Trim$(eqTypeText)   ' fallback when merged
    
    If s = "" Or s = "-" Then
        ExtractSecPosToken = ""
        Exit Function
    End If
    
    ' Normalize separators a bit
    s = Replace(s, vbTab, " ")
    s = Replace(s, "/", " ")
    s = Replace(s, "\", " ")
    s = Replace(s, ",", " ")
    
    tokens = Split(s, " ")
    
    For Each t In tokens
        t = Trim$(CStr(t))
        If Len(t) >= 3 Then
            ' Accept A-#, B-#, C-#, D-#, Q-#
            ' (Q will be treated as A later)
            If UCase$(Left$(t, 1)) Like "[ABCDQ]" And Mid$(t, 2, 1) = "-" Then
                ExtractSecPosToken = t
                Exit Function
            End If
        End If
    Next t
    
    ExtractSecPosToken = ""
End Function


'--- remove the first Sec-Pos token from a string (so we can recover Equipment Type text)
Private Function RemoveSecPosFromText(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "([A-Z])\s*-\s*(\d+)"
    re.IgnoreCase = True
    re.Global = False
    
    If re.Test(s) Then
        RemoveSecPosFromText = Trim$(re.Replace(s, ""))
    Else
        RemoveSecPosFromText = Trim$(s)
    End If
End Function


Sub FillInTheBlanks_ATT()
Dim r As Long
Dim ws As Worksheet, wb As Workbook
Dim nMembers As String

nMembers = WorksheetFunction.Max(Sheets("Geometry").Range("G:G")) + 7

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False

Set wb = ThisWorkbook
Set ws = wb.Sheets("Discrete Loads")

For r = 4 To 53
    With ws
        If .Cells(r, 14) <> "" Then
            .Cells(r, 12).Value = Range("AlphaMountAzimuth").Value
            .Cells(r, 13).Value = 1
            .Cells(r, 18) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(O" & r & "/(XLOOKUP(IF(ISNUMBER(N" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "N" & r & "," & Chr(34) & "$B" & Chr(34) & "), N" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
            .Cells(r, 19) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",Q" & r & "/(XLOOKUP(IF(ISNUMBER(P" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "P" & r & "," & Chr(34) & "$B" & Chr(34) & "), P" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
        End If
        If .Cells(r, 22) <> "" Then
            .Cells(r, 20).Value = Range("BetaMountAzimuth").Value
            .Cells(r, 21).Value = 1
            .Cells(r, 26) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(W" & r & "/(XLOOKUP(IF(ISNUMBER(V" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "V" & r & "," & Chr(34) & "$B" & Chr(34) & "), V" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
            .Cells(r, 27) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",Y" & r & "/(XLOOKUP(IF(ISNUMBER(X" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "X" & r & "," & Chr(34) & "$B" & Chr(34) & "), X" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
        End If
        If .Cells(r, 30) <> "" Then
            .Cells(r, 28).Value = Range("GammaMountAzimuth").Value
            .Cells(r, 29).Value = 1
            .Cells(r, 34) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(AE" & r & "/(XLOOKUP(IF(ISNUMBER(AD" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AD" & r & "," & Chr(34) & "$B" & Chr(34) & "), AD" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
            .Cells(r, 35) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",AG" & r & "/(XLOOKUP(IF(ISNUMBER(AF" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AF" & r & "," & Chr(34) & "$B" & Chr(34) & "), AF" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
        End If
        If .Cells(r, 38) <> "" Then
            .Cells(r, 36).Value = Range("DeltaMountAzimuth").Value
            .Cells(r, 37).Value = 1
            .Cells(r, 42) = "=IF($B" & r & "=" & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(AM" & r & "/(XLOOKUP(IF(ISNUMBER(AL" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AL" & r & "," & Chr(34) & "$B" & Chr(34) & "), AL" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1))," & Chr(34) & Chr(34) & "))))"
            .Cells(r, 43) = "=IF(OR($B" & r & "=" & Chr(34) & Chr(34) & ", $F" & r & "=1), " & Chr(34) & Chr(34) & ",MIN(100%, MAX(0%, IFERROR(IF($F" & r & "=1," & Chr(34) & Chr(34) & ",AO" & r & "/(XLOOKUP(IF(ISNUMBER(AN" & r & "),CONCATENATE(" & Chr(34) & "MP" & Chr(34) & "," & "AN" & r & "," & Chr(34) & "$B" & Chr(34) & "), AN" & r & "),Geometry!$B$6:$B$" & nMembers & ",Geometry!$O$6:$O$" & nMembers & ",0,0,1)))," & Chr(34) & Chr(34) & "))))"
        End If
    End With
Next r

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
End Sub


' Prompts for up to 4 azimuths (A,B,C,D) like:
'   0, 120, 240
'   0,, 180, 270
'   , 90, 180, 270
'
' Returns True if user provided at least one azimuth or blanks are allowed.
' Returns False if user cancels.
Private Function PromptSectorAzimuths(ByRef az() As Variant) As Boolean
    Dim s As String
    Dim parts() As String
    Dim i As Long
    Dim v As String

    ' default all to blank
    For i = 1 To 4
        az(i) = vbNullString
    Next i

    s = InputBox( _
        "Enter sector azimuths for A, B, C, D, E, F as comma-separated values." & vbCrLf & _
        "Use blanks to skip a sector." & vbCrLf & vbCrLf & _
        "Examples:" & vbCrLf & _
        "  0, 120, 240" & vbCrLf & _
        "  0,, 180, 270" & vbCrLf & _
        "  , 90, 180, 270", _
        "Mount Azimuths (A, B, C, D)" _
    )

    If StrPtr(s) = 0 Then
        ' user hit Cancel
        PromptSectorAzimuths = False
        Exit Function
    End If

    s = Trim$(s)
    If Len(s) = 0 Then
        ' user hit OK with blank -> treat as all blank (allowed)
        PromptSectorAzimuths = True
        Exit Function
    End If

    parts = Split(s, ",")

    For i = 0 To Application.Min(UBound(parts), 3)
        v = Trim$(parts(i))

        If Len(v) = 0 Then
            az(i + 1) = vbNullString
        ElseIf IsNumeric(v) Then
            ' Store as number; normalize 0-359 if you want
            az(i + 1) = CDbl(v)
        Else
            ' Non-number -> treat as blank (or reject; your call)
            az(i + 1) = vbNullString
        End If
    Next i

    PromptSectorAzimuths = True
End Function


