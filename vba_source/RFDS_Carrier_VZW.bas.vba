Attribute VB_Name = "RFDS_Carrier_VZW"
Option Explicit

'========================================================
' ENTRY POINT
'========================================================
Public Sub ImportRFDS_VZW()

    Dim fd As FileDialog
    Dim srcPath As String
    Dim wbSrc As Workbook
    Dim wsDest As Worksheet

    Dim wsAnt As Worksheet, wsNon As Worksheet, wsSvc As Worksheet
    Dim svcMap As Object   ' Project|Model -> (idNumber -> azimuth)

    Dim destRow As Long

    ' pick source file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Select VZW RFDS Extract Workbook"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
        If .Show <> -1 Then Exit Sub
        srcPath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wbSrc = Workbooks.Open(srcPath)

    ' source sheets
    Set wsAnt = GetSheetIfExists(wbSrc, "Antenna Summary")
    Set wsNon = GetSheetIfExists(wbSrc, "Non-Antenna Summary")
    Set wsSvc = GetSheetIfExists(wbSrc, "Services")

    If wsAnt Is Nothing And wsNon Is Nothing Then
        MsgBox "Could not find 'Antenna Summary' or 'Non-Antenna Summary' in the source workbook.", vbCritical
        GoTo CleanExit
    End If
    If wsSvc Is Nothing Then
        MsgBox "Could not find 'Services' sheet. Azimuth mapping will fall back to Summary strings only.", vbExclamation
    End If

    ' destination
    Set wsDest = ThisWorkbook.Worksheets("Discrete Loads")

    ' build services map (optional, but preferred)
    Set svcMap = CreateObject("Scripting.Dictionary")
    If Not wsSvc Is Nothing Then
        Set svcMap = BuildServicesAzimuthMap(wsSvc)
    End If

    ' always start at row 4
    destRow = 4

    ' import antenna summary (equip type forced to Antenna)
    If Not wsAnt Is Nothing Then
        destRow = Import_SummarySheet(wsAnt, wsDest, svcMap, destRow, "Antenna", True)
    End If

    ' import non-antenna summary (equip type forced to TME)
    If Not wsNon Is Nothing Then
        destRow = Import_SummarySheet(wsNon, wsDest, svcMap, destRow, "TME", False)
    End If

    Call Fill_Dimensions_From_Arc
    MsgBox "Verizon RF Inventory appurtenances imported into Discrete Loads.", vbInformation

CleanExit:
    On Error Resume Next
    wbSrc.Close False
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

'========================================================
' CORE IMPORT FOR ONE SUMMARY SHEET
'========================================================
Private Function Import_SummarySheet( _
    ByVal wsSrc As Worksheet, _
    ByVal wsDest As Worksheet, _
    ByVal serviceSectorMap As Object, _
    ByVal startDestRow As Long, _
    ByVal equipType As String, _
    ByVal hasCenterline As Boolean) As Long

    Dim headerRow As Long
    Dim MakeCol As Long, modelCol As Long, statusCol As Long
    Dim qtyCol As Long, azCol As Long, centerlineCol As Long
    Dim azStr As String, alphaVal As Variant, betaVal As Variant, gammaVal As Variant, deltaVal As Variant

    Dim lastRow As Long
    Dim i As Long
    Dim destRow As Long

    Dim makeVal As String, modelVal As String, statusVal As String
    Dim qtyVal As Double
    Dim azRaw As String

    Dim locationCol As Long
    Dim equipTypeCol As Long
    Dim locationVal As String
    Dim equipTypeVal As String

    destRow = startDestRow

    '---- Find header row (first 30 rows) ----
    headerRow = FindHeaderRowAny(wsSrc, Array("Make"), Array("Model"), Array("AntennaType"))
    If headerRow = 0 Then
        MsgBox "Cannot find required headers on sheet: " & wsSrc.Name, vbCritical
        Import_SummarySheet = destRow
        Exit Function
    End If

    MakeCol = FindHeaderAny(wsSrc, headerRow, Array("Make"))
    modelCol = FindHeaderAny(wsSrc, headerRow, Array("Model"))
    statusCol = FindHeaderAny(wsSrc, headerRow, Array("AntennaType"))
    qtyCol = FindHeaderAny(wsSrc, headerRow, Array("Quantity"))
    locationCol = FindHeaderAny(wsSrc, headerRow, Array("Location"))
    equipTypeCol = FindHeaderAny(wsSrc, headerRow, Array("Equipment Type", "Equip Type"))

    ' Azimuth header name varies; include a few common options
    azCol = FindHeaderAny(wsSrc, headerRow, Array("Azimuth", "Azimuths", "Azimuth(s)"))

    centerlineCol = 0
    If hasCenterline Then
        centerlineCol = FindHeaderAny(wsSrc, headerRow, Array("Centerline"))
    End If

    If MakeCol = 0 Or modelCol = 0 Or statusCol = 0 Or qtyCol = 0 Then
        MsgBox "Missing one of: Make / Model / AntennaType / Quantity on sheet: " & wsSrc.Name, vbCritical
        Import_SummarySheet = destRow
        Exit Function
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, MakeCol).End(xlUp).row

    For i = headerRow + 1 To lastRow

        makeVal = Trim$(CStr(wsSrc.Cells(i, MakeCol).Value))
        modelVal = Trim$(CStr(wsSrc.Cells(i, modelCol).Value))
        statusVal = Trim$(CStr(wsSrc.Cells(i, statusCol).Value))

        If IsNumeric(wsSrc.Cells(i, qtyCol).Value) Then
            qtyVal = CDbl(wsSrc.Cells(i, qtyCol).Value)
        Else
            qtyVal = 0
        End If
        
        '=========================================
        ' NON-ANTENNA SUMMARY FILTERS
        '=========================================
        If wsSrc.Name = "Non-Antenna Summary" Then
        
            ' --- Location filter (exclude Shelter)
            If locationCol > 0 Then
                locationVal = UCase$(Trim$(CStr(wsSrc.Cells(i, locationCol).Value)))
                If locationVal = "SHELTER" Then GoTo NextI
            End If
        
            ' --- Equipment Type filter (exclude Cable / Card / BBU)
            If equipTypeCol > 0 Then
                equipTypeVal = UCase$(Trim$(CStr(wsSrc.Cells(i, equipTypeCol).Value)))
        
                If equipTypeVal Like "*CABLE*" _
                Or equipTypeVal Like "*CARD*" _
                Or equipTypeVal Like "*BBU*" Then
                    GoTo NextI
                End If
            End If
        
        End If

        ' Omit quantity=0
        If qtyVal <= 0 Then GoTo NextI
        If makeVal = "" And modelVal = "" Then GoTo NextI

        ' === Parse Azimuths ===
        alphaVal = GetAzimuthValue(azStr, "A")
        betaVal = GetAzimuthValue(azStr, "B")
        gammaVal = GetAzimuthValue(azStr, "C")
        deltaVal = GetAzimuthValue(azStr, "D")

        '============================
        ' BASIC DEST COLUMNS
        '============================
        wsDest.Cells(destRow, "A").Value = makeVal
        wsDest.Cells(destRow, "B").Value = modelVal
        wsDest.Cells(destRow, "C").Value = MapStatusToCondition(statusVal)
        wsDest.Cells(destRow, "D").Value = equipType
        wsDest.Cells(destRow, "E").Value = "Flat"
        Select Case wsDest.Cells(destRow, "D").Value
            Case "Antenna": wsDest.Cells(destRow, "F").Value = 2
            Case Else: wsDest.Cells(destRow, "F").Value = 1
        End Select
        
        ' Defaults you added
        wsDest.Cells(destRow, "H").Value = 0
        wsDest.Cells(destRow, "I").Value = 8
        wsDest.Cells(destRow, "J").Value = 1
        wsDest.Cells(destRow, "K").Value = 1

        ' G = Centerline - Code!H16 (Antenna Summary only)
        If hasCenterline And centerlineCol > 0 Then
            Dim clVal As Variant
            clVal = wsSrc.Cells(i, centerlineCol).Value
            If IsNumeric(clVal) Then
                wsDest.Cells(destRow, "G").Value = CDbl(clVal) - ThisWorkbook.Worksheets("Code").Range("H16").Value
            Else
                wsDest.Cells(destRow, "G").Value = 0
            End If
        Else
            wsDest.Cells(destRow, "G").Value = 0
        End If

        '============================
        ' AZIMUTH + QUANTITY BY SECTOR
        '============================
        ' Clear destination az/qty cells first (so no leftovers)
'        ClearSectorAzQty wsDest, destRow

        If azCol > 0 Then
            azRaw = CStr(wsSrc.Cells(i, azCol).Value)

            ' Treat blank or "-" as no azimuths
            If Trim$(azRaw) <> "" And Trim$(azRaw) <> "-" Then
                ' remove ALL spaces to fix "32 0" -> "320"
                azRaw = Replace(azRaw, " ", "")

                ApplyAzimuthsFromString wsDest, destRow, azRaw, serviceSectorMap
            End If
        End If
        
'            ' === Alpha (L–S) ===
'            If Nz(alphaVal, 0) <> 0 Then
'                wsDest.Cells(destRow, "L").Value = alphaVal
'            End If
'
'            ' === Beta (T-AA) ===
'            If Nz(betaVal, 0) <> 0 Then
'                wsDest.Cells(destRow, "T").Value = betaVal
'            End If
'
'            ' === Gamma (AB–AI) ===
'            If Nz(gammaVal, 0) <> 0 Then
'                wsDest.Cells(destRow, "AB").Value = gammaVal
'            End If
'
'            ' === Delta (AJ–AQ) ===
'            If Nz(deltaVal, 0) <> 0 Then
'                wsDest.Cells(destRow, "AJ").Value = deltaVal
'            End If

        destRow = destRow + 1

NextI:
    Next i

    Import_SummarySheet = destRow

End Function

Function GetAzimuthValue(azStr As String, sector As String) As Double
    Dim regex As Object, matches As Object
    Dim cleaned As String, pattern As String
    Dim i As Long

    ' --- Clean and normalize ---
    cleaned = Replace(azStr, ChrW(160), " ")
    cleaned = Replace(cleaned, vbTab, " ")
    cleaned = Trim(cleaned)
    cleaned = Replace(cleaned, ",", ", ")
    cleaned = Replace(cleaned, "  ", " ")

    ' Remove spaces between digits (e.g., 2 50 -> 250)
    For i = 48 To 57
        cleaned = Replace(cleaned, Chr(i) & " ", Chr(i))
    Next i

    ' --- Unified pattern: number followed by (A|1|B|2|C|3|D|4)
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "(\d+(?:\.\d+)?)\s*\(\s*([A-D1-4])\s*\)"

    If regex.Test(cleaned) Then
        Set matches = regex.Execute(cleaned)
        Dim matchObj As Object
        For Each matchObj In matches
            Dim numVal As Double, sectorTag As String
            numVal = CDbl(matchObj.SubMatches(0))
            sectorTag = UCase(matchObj.SubMatches(1))

            Select Case True
                Case (UCase(sector) = "A" And (sectorTag = "A" Or sectorTag = "01"))
                    GetAzimuthValue = numVal
                    Exit Function
                Case (UCase(sector) = "B" And (sectorTag = "B" Or sectorTag = "02"))
                    GetAzimuthValue = numVal
                    Exit Function
                Case (UCase(sector) = "C" And (sectorTag = "C" Or sectorTag = "03"))
                    GetAzimuthValue = numVal
                    Exit Function
                Case (UCase(sector) = "D" And (sectorTag = "D" Or sectorTag = "04"))
                    GetAzimuthValue = numVal
                    Exit Function
            End Select
        Next
    End If

    ' Default 0 if no match
    GetAzimuthValue = 0
End Function

'========================================================
' AZIMUTH PARSING + APPLY
'   Input example:
'     90(149),210(150),320(151)
'========================================================
Private Sub ApplyAzimuthsFromString( _
    ByVal wsDest As Worksheet, _
    ByVal destRow As Long, _
    ByVal azText As String, _
    ByVal serviceSectorMap As Object)

    Dim qtySectorsTotal As Long
    qtySectorsTotal = CLng(ThisWorkbook.Worksheets("Code").Range("H18").Value) ' named range QtySectorsTotal
    If qtySectorsTotal < 1 Then qtySectorsTotal = 3
    If qtySectorsTotal > 4 Then qtySectorsTotal = 4 ' cap to 4 sectors in destination layout

    ' We'll collect:
    '   sectorIndexNormalized -> count
    '   sectorIndexNormalized -> firstAzimuth
    Dim cnt(1 To 4) As Long
    Dim firstAz(1 To 4) As Variant
    Dim hasFirst(1 To 4) As Boolean

    Dim parts As Variant, p As Variant
    parts = Split(azText, ",")

    For Each p In parts
        Dim token As String
        token = Trim$(CStr(p))
        If token = "" Then GoTo NextP

        ' Expect something like 90(149)
        Dim ang As Double, svcId As String
        If Not TryParseAzToken(token, ang, svcId) Then GoTo NextP

        ' Look up sector number from Services sheet map
        If Not serviceSectorMap.Exists(svcId) Then GoTo NextP

        Dim sectorNum As Long
        sectorNum = CLng(serviceSectorMap(svcId))

        ' Normalize sector number into 1..qtySectorsTotal (wrap 4..6 to 1..3, etc.)
        Dim norm As Long
        norm = ((sectorNum - 1) Mod qtySectorsTotal) + 1

        If norm >= 1 And norm <= 4 Then
            cnt(norm) = cnt(norm) + 1
            If Not hasFirst(norm) Then
                firstAz(norm) = ang
                hasFirst(norm) = True
            End If
        End If

NextP:
    Next p

    ' Write to destination: Azimuth + Quantity
    ' Alpha
    If cnt(1) > 0 Then
        wsDest.Cells(destRow, "L").Value = firstAz(1)
        wsDest.Cells(destRow, "M").Value = cnt(1)
    End If

    ' Beta
    If cnt(2) > 0 Then
        wsDest.Cells(destRow, "T").Value = firstAz(2)
        wsDest.Cells(destRow, "U").Value = cnt(2)
    End If

    ' Gamma
    If cnt(3) > 0 Then
        wsDest.Cells(destRow, "AB").Value = firstAz(3)
        wsDest.Cells(destRow, "AC").Value = cnt(3)
    End If

    ' Delta (only used if QtySectorsTotal=4 and we actually got values)
    If cnt(4) > 0 Then
        wsDest.Cells(destRow, "AJ").Value = firstAz(4)
        wsDest.Cells(destRow, "AK").Value = cnt(4)
    End If

End Sub

' Parse a token like "90(149)" into angle=90 and svcId="149"
Private Function TryParseAzToken(ByVal token As String, ByRef angleOut As Double, ByRef svcIdOut As String) As Boolean
    On Error GoTo Fail

    Dim p1 As Long, p2 As Long
    p1 = InStr(1, token, "(", vbBinaryCompare)
    p2 = InStr(1, token, ")", vbBinaryCompare)

    If p1 <= 1 Or p2 <= p1 + 1 Then GoTo Fail

    Dim angStr As String, idStr As String
    angStr = Left$(token, p1 - 1)
    idStr = Mid$(token, p1 + 1, p2 - p1 - 1)

    If Not IsNumeric(angStr) Then GoTo Fail
    angleOut = CDbl(angStr)
    svcIdOut = idStr

    TryParseAzToken = True
    Exit Function

Fail:
    TryParseAzToken = False
End Function
'
'Private Sub ClearSectorAzQty(ByVal ws As Worksheet, ByVal r As Long)
'    ws.Cells(r, "L").ClearContents:  ws.Cells(r, "M").ClearContents   ' Alpha
'    ws.Cells(r, "T").ClearContents:  ws.Cells(r, "U").ClearContents   ' Beta
'    ws.Cells(r, "AB").ClearContents: ws.Cells(r, "AC").ClearContents  ' Gamma
'    ws.Cells(r, "AJ").ClearContents: ws.Cells(r, "AK").ClearContents  ' Delta
'End Sub

'========================================================
' SERVICES MAP BUILDER
'========================================================
Private Function BuildServiceSectorMap(ByVal wsSvc As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim headerRow As Long
    headerRow = FindHeaderRowAny(wsSvc, Array("Sector"), Array("Sector"), Array("Sector"))
    If headerRow = 0 Then headerRow = 1

    ' Service ID header varies a lot; try a few
    Dim idCol As Long, secCol As Long
    idCol = FindHeaderAny(wsSvc, headerRow, Array("Service", "Service ID", "ServiceID", "Svc", "ID"))
    secCol = FindHeaderAny(wsSvc, headerRow, Array("Sector", "Sector #", "Sector Number"))

    If idCol = 0 Or secCol = 0 Then
        ' Return empty dict; caller will just skip if not found
        Set BuildServiceSectorMap = dict
        Exit Function
    End If

    Dim lastRow As Long, i As Long
    lastRow = wsSvc.Cells(wsSvc.Rows.Count, idCol).End(xlUp).row

    For i = headerRow + 1 To lastRow
        Dim k As String, s As Variant
        k = Trim$(CStr(wsSvc.Cells(i, idCol).Value))
        s = wsSvc.Cells(i, secCol).Value

        If Len(k) > 0 And IsNumeric(s) Then
            If Not dict.Exists(k) Then dict.Add k, CLng(s)
        End If
    Next i

    Set BuildServiceSectorMap = dict
End Function

'========================================================
' SERVICES MAP (Project|Model -> idNum -> azimuth)
'========================================================
Private Function BuildServicesAzimuthMap(ByVal wsSvc As Worksheet) As Object

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim hdrs As Variant
    hdrs = Array("Fuze Project ID", "Sector", "Azimuth", "Antenna Model")

    Dim headerRow As Long
    headerRow = FindHeaderRow_VZW(wsSvc, hdrs)
    If headerRow = 0 Then
        Set BuildServicesAzimuthMap = dict
        Exit Function
    End If

    Dim projCol As Long, sectorCol As Long, azCol As Long, modelCol As Long
    projCol = FindHeader(wsSvc, headerRow, "Fuze Project ID")
    sectorCol = FindHeader(wsSvc, headerRow, "Sector")
    azCol = FindHeader(wsSvc, headerRow, "Azimuth")
    modelCol = FindHeader(wsSvc, headerRow, "Antenna Model")
    If projCol = 0 Or sectorCol = 0 Or azCol = 0 Or modelCol = 0 Then
        Set BuildServicesAzimuthMap = dict
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = wsSvc.Cells(wsSvc.Rows.Count, projCol).End(xlUp).row

    Dim i As Long
    For i = headerRow + 1 To lastRow

        Dim projID As String, modelVal As String
        projID = Trim$(CStr(wsSvc.Cells(i, projCol).Value))
        modelVal = Trim$(CStr(wsSvc.Cells(i, modelCol).Value))
        If Len(projID) = 0 Or Len(modelVal) = 0 Then GoTo NextI

        Dim idNum As Long
        idNum = ExtractDigitsAsLong(CStr(wsSvc.Cells(i, sectorCol).Value))
        If idNum <= 0 Then GoTo NextI

        Dim az As Double
        If Not IsNumeric(wsSvc.Cells(i, azCol).Value) Then GoTo NextI
        az = CDbl(wsSvc.Cells(i, azCol).Value)

        Dim key As String
        key = UCase$(projID) & "|" & UCase$(modelVal)

        Dim subDict As Object
        If Not dict.Exists(key) Then
            Set subDict = CreateObject("Scripting.Dictionary")
            dict.Add key, subDict
        Else
            Set subDict = dict(key)
        End If

        If Not subDict.Exists(idNum) Then subDict.Add idNum, az

NextI:
    Next i

    Set BuildServicesAzimuthMap = dict
End Function

Private Function ExtractDigitsAsLong(ByVal s As String) As Long
    ' turns "0151" -> 151, "01" -> 1, "D3" -> 3, etc.
    Dim i As Long, ch As String, out As String
    s = Trim$(s)
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then out = out & ch
    Next i
    If Len(out) = 0 Then
        ExtractDigitsAsLong = 0
    Else
        ExtractDigitsAsLong = CLng(val(out))
    End If
End Function

'========================================================
' PARSE AZIMUTH STRING
' Supports:
'   90(149),210(150 ),320(151)
'   90(1),210(2),32 0(3),...
' Steps:
'   - remove spaces
'   - split by comma
'   - each token: az + optional (idNum)
'========================================================
Private Function ParseAzimuthTokens(ByVal azText As String) As Collection

    Dim col As New Collection
    azText = Replace(Trim$(azText), " ", "")  ' remove ALL spaces (32 0 -> 320)

    If Len(azText) = 0 Or azText = "-" Then
        Set ParseAzimuthTokens = col
        Exit Function
    End If

    Dim parts As Variant, p As Variant
    parts = Split(azText, ",")

    For Each p In parts
        Dim token As String
        token = Trim$(CStr(p))
        If token = "" Or token = "-" Then GoTo NextP

        Dim azStr As String, idStr As String
        Dim az As Double, idNum As Long

        If InStr(1, token, "(", vbTextCompare) > 0 And InStr(1, token, ")", vbTextCompare) > 0 Then
            azStr = Left$(token, InStr(1, token, "(") - 1)
            idStr = Mid$(token, InStr(1, token, "(") + 1, InStr(1, token, ")") - InStr(1, token, "(") - 1)
        Else
            azStr = token
            idStr = ""
        End If

        If IsNumeric(azStr) Then
            az = CDbl(azStr)
        Else
            GoTo NextP
        End If

        idNum = 0
        If Len(idStr) > 0 Then idNum = ExtractDigitsAsLong(idStr)

        col.Add Array(az, idNum)

NextP:
    Next p

    Set ParseAzimuthTokens = col
End Function

''========================================================
'' SEC-POS WRITE (your column mapping)
''========================================================
'Private Sub WriteSecPosToSector(ByVal wsDest As Worksheet, ByVal destRow As Long, ByVal secLetter As String, ByVal secPos As String)
'    Select Case UCase$(secLetter)
'        Case "A"
'            wsDest.Cells(destRow, "N").Value = secPos
'            wsDest.Cells(destRow, "P").Value = secPos
'        Case "B"
'            wsDest.Cells(destRow, "V").Value = secPos
'            wsDest.Cells(destRow, "X").Value = secPos
'        Case "C"
'            wsDest.Cells(destRow, "AD").Value = secPos
'            wsDest.Cells(destRow, "AF").Value = secPos
'        Case "D"
'            wsDest.Cells(destRow, "AL").Value = secPos
'            wsDest.Cells(destRow, "AN").Value = secPos
'    End Select
'End Sub

'========================================================
' HELPERS
'========================================================
Private Function MapStatusToCondition(ByVal statusVal As String) As String
    Dim s As String
    s = UCase$(Trim$(statusVal))

    Select Case s
        Case "ADDED ANTENNA", "ADDED NON ANTENNA"
            MapStatusToCondition = "Proposed"
        Case "RETAINED ANTENNA", "RETAINED NON ANTENNA"
            MapStatusToCondition = "Existing"
        Case "REMOVED ANTENNA", "REMOVED NON ANTENNA"
            MapStatusToCondition = "Removed"
        Case Else
            MapStatusToCondition = statusVal
    End Select
End Function

Private Function GetSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetIfExists = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

' Find a header row where ALL three “families” exist
Private Function FindHeaderRow_VZW(ByVal ws As Worksheet, ByVal headers As Variant, _
                                     Optional ByVal maxRowsToScan As Long = 50) As Long
    Dim r As Long, i As Long
    Dim f As Range
    Dim foundAll As Boolean

    For r = 1 To maxRowsToScan
        foundAll = True

        For i = LBound(headers) To UBound(headers)
            Set f = ws.Rows(r).Find(What:=CStr(headers(i)), LookIn:=xlValues, LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
            If f Is Nothing Then
                foundAll = False
                Exit For
            End If
        Next i

        If foundAll Then
            FindHeaderRow_VZW = r
            Exit Function
        End If
    Next r

    FindHeaderRow_VZW = 0
End Function

Private Function FindHeaderRowAny(ws As Worksheet, h1List As Variant, h2List As Variant, h3List As Variant) As Long
    Dim r As Long
    For r = 1 To 30
        If (FindHeaderAny(ws, r, h1List) > 0) And _
           (FindHeaderAny(ws, r, h2List) > 0) And _
           (FindHeaderAny(ws, r, h3List) > 0) Then
            FindHeaderRowAny = r
            Exit Function
        End If
    Next r
End Function

Private Function FindHeaderAny(ws As Worksheet, headerRow As Long, headers As Variant) As Long
    Dim h As Variant, c As Range
    For Each h In headers
        Set c = ws.Rows(headerRow).Find(What:=CStr(h), LookIn:=xlValues, LookAt:=xlPart)
        If Not c Is Nothing Then
            FindHeaderAny = c.Column
            Exit Function
        End If
    Next h
End Function

Private Function FindHeader(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim c As Range
    Set c = ws.Rows(headerRow).Find(headerText, LookIn:=xlValues, LookAt:=xlPart)
    If Not c Is Nothing Then FindHeader = c.Column Else FindHeader = 0
End Function


Sub FillInTheBlanks_VZW()
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

