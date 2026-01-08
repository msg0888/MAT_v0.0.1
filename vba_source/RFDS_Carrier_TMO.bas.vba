Attribute VB_Name = "RFDS_Carrier_TMO"
Option Explicit

' ======================================================================
'  T-MOBILE SECTOR LAYOUT ? DISCRETE LOADS (MOUNT ANALYSIS)
'  - Handles sheets named: Sector_1_Proposed, Sector_2_Existing, etc.
'  - One row per unique (Antenna_Model + Condition)
'  - Sector 1?Alpha, 2?Beta, 3?Gamma, 4?Delta
'  - Condition from sheet name (Proposed / Existing / Removed)
'  - Model parsing preserves format, strips brackets
'  - Defaults: Type="TME (Front)", Shape="Flat", shields/offsets as before
' ======================================================================
Public Sub ImportRFDS_TMO()

    Dim fd As FileDialog
    Dim srcPath As String
    Dim wbSrc As Workbook
    Dim wsDest As Worksheet
    Dim recs As Object          ' Dictionary of antenna records
    Dim ws As Worksheet

    '---- Pick source workbook ----
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "Select T-Mobile Converted RFDS"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
    If fd.Show <> -1 Then Exit Sub
    srcPath = fd.SelectedItems(1)

    Application.ScreenUpdating = False

    Set wbSrc = Workbooks.Open(srcPath)
    Set wsDest = ThisWorkbook.Worksheets("Discrete Loads")
    Set recs = CreateObject("Scripting.Dictionary")

    '---- Scan all Sector_*_* sheets ----
    For Each ws In wbSrc.Worksheets
        Process_Sector_Sheet ws, recs
    Next ws

    '---- Write results to Discrete Loads ----
    Write_Records_To_DiscreteLoads wsDest, recs

    wbSrc.Close False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "T-Mobile RF Inventory appurtenances imported into Discrete Loads.", vbInformation
    Call Fill_Dimensions_From_Arc

End Sub

' ======================================================================
'  PROCESS ONE SECTOR SHEET
' ======================================================================
Private Sub Process_Sector_Sheet(ws As Worksheet, recs As Object)

    Dim sectorIdx As Long
    Dim condText As String
    Dim ok As Boolean
    Dim headerRow As Long
    Dim colModel As Long, colAz As Long, colAzBuilt As Long, colAzDes As Long
    Dim colRadio As Long, colTMAs As Long
    Dim lastRow As Long, i As Long

    Dim rawModel As String
    Dim makeRaw As String, modelDisplay As String
    Dim parentMake As String
    Dim normKey As String
    Dim azVal As Double
    Dim hasAz As Boolean

    Dim recKey As String
    Dim rec As Object

    '---- Parse sheet name: Sector_n_Condition ----
    ok = Parse_Sector_Sheet_Name(ws.Name, sectorIdx, condText)
    If Not ok Then Exit Sub          ' not a sector sheet, ignore

    '---- Find header row / columns ----
    headerRow = FindHeaderRow(ws, "Antenna_Model", "Azimuth", "")
    If headerRow = 0 Then Exit Sub

    colModel = FindHeader(ws, "Antenna_Model", headerRow)
    colAzBuilt = FindHeader(ws, "Azimuth (As-Built)", headerRow)
    colAz = FindHeader(ws, "Azimuth", headerRow)
    colAzDes = FindHeader(ws, "Azimuth (Design)", headerRow)
    colRadio = FindHeader(ws, "Radio", headerRow)
    colTMAs = FindHeader(ws, "TMAs", headerRow)

    If colModel = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, colModel).End(xlUp).row
    If lastRow <= headerRow Then Exit Sub

    '---- Loop rows ----
    For i = headerRow + 1 To lastRow

        rawModel = Trim(CStr(ws.Cells(i, colModel).Value))
        If rawModel <> "" Then

            ' Parse Make + Model display
            TMO_Parse_Make_Model rawModel, makeRaw, modelDisplay
            parentMake = TMO_ParentManufacturer(makeRaw)
            normKey = Normalize_Model_Key(modelDisplay)

            ' Get azimuth with priority: As-Built > Azimuth > Design
            azVal = 0
            hasAz = Get_Row_Azimuth(ws, i, colAzBuilt, colAz, colAzDes, azVal)

            ' Build or fetch record (one per Condition + Make + Model)
            recKey = UCase(condText) & "||" & UCase(parentMake) & "||" & normKey

            If Not recs.Exists(recKey) Then
                Set rec = CreateObject("Scripting.Dictionary")
                rec("Condition") = condText
                rec("Manufacturer") = parentMake
                rec("ModelDisplay") = modelDisplay
                rec("AlphaAz") = Empty
                rec("BetaAz") = Empty
                rec("GammaAz") = Empty
                rec("DeltaAz") = Empty
                recs.Add recKey, rec
            Else
                Set rec = recs(recKey)
            End If

            ' Attach azimuth to proper sector (first valid wins)
            If hasAz Then
                Select Case sectorIdx
                    Case 1
                        If IsEmpty(rec("AlphaAz")) Then rec("AlphaAz") = azVal
                    Case 2
                        If IsEmpty(rec("BetaAz")) Then rec("BetaAz") = azVal
                    Case 3
                        If IsEmpty(rec("GammaAz")) Then rec("GammaAz") = azVal
                    Case 4
                        If IsEmpty(rec("DeltaAz")) Then rec("DeltaAz") = azVal
                End Select
            End If

        End If
        Dim v As String
        Dim vClean As String
        
        ' Radio -> TME
        If colRadio > 0 Then
            v = Trim$(CStr(ws.Cells(i, colRadio).Value))
            vClean = Replace$(v, " ", "")     ' ? ADD THIS LINE
        
            If vClean <> "" And vClean <> "[]" Then
                Add_TMO_Record recs, condText, v, "TME", sectorIdx, hasAz, azVal
            End If
        End If
    
        ' TMAs -> TME
        If colTMAs > 0 Then
            v = Trim$(CStr(ws.Cells(i, colTMAs).Value))
            vClean = Replace$(v, " ", "")     ' ? ADD THIS LINE
        
            If vClean <> "" And vClean <> "[]" Then
                Add_TMO_Record recs, condText, v, "TME", sectorIdx, hasAz, azVal
            End If
        End If
            Next i

End Sub

' ======================================================================
'  WRITE DICTIONARY INTO DISCRETE LOADS
' ======================================================================
Private Sub Write_Records_To_DiscreteLoads(wsDest As Worksheet, recs As Object)

    Dim destRow As Long
    Dim key As Variant
    Dim rec As Object
    Dim hasAnyAz As Boolean
    Dim modVal As String
    Dim manVal As String
    Dim manufacturerIdx As Long

    destRow = 4   ' start row
    manufacturerIdx = 0

    For Each key In recs.Keys
        Set rec = recs(key)

        hasAnyAz = Not (IsEmpty(rec("AlphaAz")) And _
                        IsEmpty(rec("BetaAz")) And _
                        IsEmpty(rec("GammaAz")) And _
                        IsEmpty(rec("DeltaAz")))

        If hasAnyAz Then

            If Trim(rec("Manufacturer") & "") <> "" Then
'                manVal = Trim(rec("Manufacturer"))
                manufacturerIdx = manufacturerIdx + 1
                manVal = "MANUFACTURER #" & manufacturerIdx
                modVal = Trim(rec("ModelDisplay"))
            Else
                manufacturerIdx = manufacturerIdx + 1
                manVal = "MANUFACTURER #" & manufacturerIdx
                modVal = rec("ModelDisplay")
            End If

            ' Basic info
            wsDest.Cells(destRow, "A").Value = manVal
            wsDest.Cells(destRow, "B").Value = modVal
            wsDest.Cells(destRow, "C").Value = rec("Condition")
            wsDest.Cells(destRow, "D").Value = "Antenna" '"TME (Front)"
            wsDest.Cells(destRow, "E").Value = "Flat"
            wsDest.Cells(destRow, "F").Value = 2
            wsDest.Cells(destRow, "G").Value = 0
            wsDest.Cells(destRow, "H").Value = 0
            wsDest.Cells(destRow, "I").Value = 8
            wsDest.Cells(destRow, "J").Value = 1
            wsDest.Cells(destRow, "K").Value = 1

            ' Alpha / Beta / Gamma / Delta sectors
            If Not IsEmpty(rec("AlphaAz")) Then FillAlpha wsDest, destRow, rec("AlphaAz")
            If Not IsEmpty(rec("BetaAz")) Then FillBeta wsDest, destRow, rec("BetaAz")
            If Not IsEmpty(rec("GammaAz")) Then FillGamma wsDest, destRow, rec("GammaAz")
            If Not IsEmpty(rec("DeltaAz")) Then FillDelta wsDest, destRow, rec("DeltaAz")

            destRow = destRow + 1
        End If
    Next key

End Sub

' ======================================================================
'  SECTOR NAME PARSER: "Sector_1_Proposed"
' ======================================================================
Private Function Parse_Sector_Sheet_Name( _
    ByVal sheetName As String, _
    ByRef sectorIdx As Long, _
    ByRef condText As String) As Boolean

    Dim s As String
    s = LCase$(Trim$(sheetName))

    ' normalize common separators to spaces
    s = Replace(s, "_", " ")
    s = Replace(s, "-", " ")
    s = Application.WorksheetFunction.Trim(s) ' collapses multiple spaces

    Dim parts() As String
    parts = Split(s, " ")

    ' Need at least: "sector" + number
    If UBound(parts) < 1 Then Exit Function
    If parts(0) <> "sector" Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function

    sectorIdx = CLng(parts(1))

    ' Default condition if missing (e.g., "Sector 1")
    condText = "Proposed"

    If UBound(parts) >= 2 Then
        Select Case parts(2)
            Case "proposed": condText = "Proposed"
            Case "existing": condText = "Existing"
            Case "removed":  condText = "Removed"
            Case Else
                ' keep whatever it is, but nicely cased
                condText = StrConv(parts(2), vbProperCase)
        End Select
    End If

    Parse_Sector_Sheet_Name = True
End Function

' ======================================================================
'  AZIMUTH FROM ROW WITH PRIORITY
' ======================================================================
Private Function Get_Row_Azimuth(ws As Worksheet, rowNum As Long, _
                                 colAzBuilt As Long, colAz As Long, _
                                 colAzDes As Long, _
                                 ByRef azVal As Double) As Boolean
    Dim cols(1 To 3) As Long
    Dim j As Long, c As Long
    Dim txt As String

    cols(1) = colAzBuilt
    cols(2) = colAz
    cols(3) = colAzDes

    For j = 1 To 3
        c = cols(j)
        If c > 0 Then
            txt = Trim(CStr(ws.Cells(rowNum, c).Value))
            If txt <> "" And txt <> "-" Then
                If IsNumeric(txt) Then
                    azVal = CDbl(txt)        ' 0° allowed
                    Get_Row_Azimuth = True
                    Exit Function
                End If
            End If
        End If
    Next j
End Function

' ======================================================================
'  MODEL / MAKE PARSING HELPERS
' ======================================================================
Private Sub TMO_Parse_Make_Model(rawFull As String, _
                                 ByRef outMake As String, _
                                 ByRef outModelDisplay As String)
    Dim base As String
    Dim posDash As Long
    Dim firstUnd As Long
    Dim firstToken As String

    base = Strip_Parentheses(rawFull)

    ' Case 1: "Make - Model"
    posDash = InStr(base, " - ")
    If posDash > 0 Then
        outMake = Trim(Left$(base, posDash - 1))
        outModelDisplay = Trim(Mid$(base, posDash + 3))
        Exit Sub
    End If

    ' Case 2: "Make_Model" where Make is known
    firstUnd = InStr(base, "_")
    If firstUnd > 0 Then
        firstToken = Left$(base, firstUnd - 1)
        If Is_Known_Make(firstToken) Then
            outMake = firstToken
            outModelDisplay = Trim(Mid$(base, firstUnd + 1))
            Exit Sub
        End If
    End If

    ' Case 3: Model only
    outMake = ""
    outModelDisplay = Trim(base)
End Sub

Private Function Strip_Parentheses(s As String) As String
    Dim p As Long
    p = InStr(s, "(")
    If p > 0 Then
        Strip_Parentheses = Trim(Left$(s, p - 1))
    Else
        Strip_Parentheses = Trim(s)
    End If
End Function

Private Function Is_Known_Make(m As String) As Boolean
    Dim u As String
    u = UCase$(Trim(m))
    Select Case u
        Case "COMMSCOPE", "ANDREW", "ERICSSON", "RFS", "AMPHENOL", "GALTRONICS"
            Is_Known_Make = True
        Case Else
            Is_Known_Make = False
    End Select
End Function

Private Function Normalize_Model_Key(modelDisplay As String) As String
    Dim s As String
    Dim i As Long, ch As String

    s = UCase$(modelDisplay)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch Like "[A-Z0-9]") Then
            Normalize_Model_Key = Normalize_Model_Key & ch
        End If
    Next i
End Function

Private Function TMO_ParentManufacturer(makeIn As String) As String
    Dim u As String
    u = UCase$(Trim(makeIn))
    Select Case u
        Case "COMMSCOPE", "ANDREW"
            TMO_ParentManufacturer = "Commscope"
        Case "ERICSSON"
            TMO_ParentManufacturer = "Ericsson"
        Case "RFS"
            TMO_ParentManufacturer = "RFS"
        Case "GALTRONICS"
            TMO_ParentManufacturer = "Galtronics"
        Case Else
            TMO_ParentManufacturer = makeIn
    End Select
End Function

' ======================================================================
'  DEFAULT-FILL HELPERS FOR DISCRETE LOADS
' ======================================================================
Private Sub FillAlpha(ws As Worksheet, rowNum As Long, azVal As Variant)
    ws.Cells(rowNum, "L").Value = azVal   ' Azimuth
    ws.Cells(rowNum, "M").Value = 1       ' Quantity
End Sub

Private Sub FillBeta(ws As Worksheet, rowNum As Long, azVal As Variant)
    ws.Cells(rowNum, "T").Value = azVal
    ws.Cells(rowNum, "U").Value = 1
End Sub

Private Sub FillGamma(ws As Worksheet, rowNum As Long, azVal As Variant)
    ws.Cells(rowNum, "AB").Value = azVal
    ws.Cells(rowNum, "AC").Value = 1
End Sub

Private Sub FillDelta(ws As Worksheet, rowNum As Long, azVal As Variant)
    ws.Cells(rowNum, "AJ").Value = azVal
    ws.Cells(rowNum, "AK").Value = 1
End Sub

' ======================================================================
'  GENERIC HEADER HELPERS (REUSED)
' ======================================================================
Function FindHeaderRow(ws As Worksheet, h1 As String, h2 As String, h3 As String) As Long
    Dim r As Long, f1 As Range, f2 As Range, f3 As Range
    For r = 1 To 20
        Set f1 = ws.Rows(r).Find(h1, , xlValues, xlPart, , , False)
        Set f2 = ws.Rows(r).Find(h2, , xlValues, xlPart, , , False)
        If h3 <> "" Then
            Set f3 = ws.Rows(r).Find(h3, , xlValues, xlPart, , , False)
        Else
            Set f3 = Nothing
        End If
        If Not f1 Is Nothing And Not f2 Is Nothing Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Function FindHeader(ws As Worksheet, headerText As String, headerRow As Long) As Long
    Dim c As Range
    Set c = ws.Rows(headerRow).Find(headerText, , xlValues, xlPart, , , False)
    If Not c Is Nothing Then FindHeader = c.Column
End Function

Private Sub Add_TMO_Record( _
    ByVal recs As Object, _
    ByVal condText As String, _
    ByVal rawItem As String, _
    ByVal equipType As String, _
    ByVal sectorIdx As Long, _
    ByVal hasAz As Boolean, _
    ByVal azVal As Double)

    Dim items As Collection, it As Variant
    Set items = SplitAndCleanItems(rawItem)

    For Each it In items
        Dim modelDisplay As String
        modelDisplay = CStr(it)
        If modelDisplay = "" Then GoTo NextIt

        Dim normKey As String, recKey As String
        normKey = Normalize_Model_Key(modelDisplay)

        ' IMPORTANT: include equipType so Antenna vs TME don't merge
        recKey = UCase$(equipType) & "||" & UCase$(condText) & "||" & normKey

        Dim rec As Object
        If Not recs.Exists(recKey) Then
            Set rec = CreateObject("Scripting.Dictionary")
            rec("Condition") = condText
            rec("Manufacturer") = "-"          ' you can improve later if you want
            rec("ModelDisplay") = modelDisplay
            rec("EquipType") = equipType       ' "Antenna" or "TME"
            rec("AlphaAz") = Empty
            rec("BetaAz") = Empty
            rec("GammaAz") = Empty
            rec("DeltaAz") = Empty
            recs.Add recKey, rec
        Else
            Set rec = recs(recKey)
        End If

        If hasAz Then
            Select Case sectorIdx
                Case 1: If IsEmpty(rec("AlphaAz")) Then rec("AlphaAz") = azVal
                Case 2: If IsEmpty(rec("BetaAz")) Then rec("BetaAz") = azVal
                Case 3: If IsEmpty(rec("GammaAz")) Then rec("GammaAz") = azVal
                Case 4: If IsEmpty(rec("DeltaAz")) Then rec("DeltaAz") = azVal
            End Select
        End If

NextIt:
    Next it
End Sub

Private Function SplitAndCleanItems(ByVal s As String) As Collection
    Dim out As New Collection
    Dim t As String
    t = s

    ' normalize common junk that causes "extra characters"
    t = Replace(t, ChrW(160), " ") ' NBSP
    t = Replace(t, "•", " ")
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)

    ' split on line breaks first
    Dim arr() As String, i As Long, piece As String
    arr = Split(t, vbLf)

    For i = LBound(arr) To UBound(arr)
        piece = Trim$(arr(i))
        If piece <> "" Then
            ' also split comma/semicolon lists on the same line
            piece = Replace(piece, ";", ",")
            Dim arr2() As String, j As Long, p2 As String
            arr2 = Split(piece, ",")
            For j = LBound(arr2) To UBound(arr2)
                p2 = Trim$(arr2(j))
                If p2 <> "" Then out.Add p2
            Next j
        End If
    Next i

    Set SplitAndCleanItems = out
End Function


