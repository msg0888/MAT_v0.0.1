Attribute VB_Name = "RFDS_tnxProcess"
Option Explicit

'========================================================
' Helpers (existing)
'========================================================

Private Function IsCciDatabaseName(ByVal fileBaseName As String) As Boolean
    IsCciDatabaseName = (LCase$(Right$(Trim$(fileBaseName), 6)) = " (cci)")
End Function

Private Function EndsWithTIA(ByVal model As String) As Boolean
    EndsWithTIA = (LCase$(Right$(Trim$(model), 4)) = "_tia")
End Function

' Reads one .arc file into a dictionary keyed by USName (Model)
' dict(key) = Variant() array of numeric strings from the Values= line
Private Sub LoadArcFileToDict( _
    ByVal folderPath As String, _
    ByVal fileBaseName As String, _
    ByRef dict As Object)

    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim nums As Variant
    Dim currentUSName As String
    Dim filePath As String

    ' Safety: do not ever use "(cci)" databases
    If IsCciDatabaseName(fileBaseName) Then
        Set dict = CreateObject("Scripting.Dictionary")
        Exit Sub
    End If

    filePath = folderPath & "\" & fileBaseName & ".arc"

    If Dir(filePath) = "" Then
        Set dict = CreateObject("Scripting.Dictionary")
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)   ' ForReading

    Do While Not ts.AtEndOfStream
        line = ts.ReadLine

        If LCase$(Left$(line, 7)) = "usname=" Then
            currentUSName = Mid$(line, 8)
        End If

        If LCase$(Left$(line, 7)) = "values=" Then
            nums = Split(Mid$(line, 8), " ")
            If UBound(nums) >= 14 Then
                If Len(currentUSName) > 0 Then
                    dict(currentUSName) = nums
                End If
            End If
        End If
    Loop

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub


Private Sub LoadArcFileToDict_ByFullPath(ByVal fullPath As String, ByRef dict As Object)

    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim nums As Variant
    Dim currentUSName As String

    Set dict = CreateObject("Scripting.Dictionary")

    If Len(Dir(fullPath)) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(fullPath, 1)   ' ForReading

    Do While Not ts.AtEndOfStream
        line = ts.ReadLine

        If LCase$(Left$(line, 7)) = "usname=" Then
            currentUSName = Mid$(line, 8)
        End If

        If LCase$(Left$(line, 7)) = "values=" Then
            nums = Split(Mid$(line, 8), " ")
            If UBound(nums) >= 14 Then
                If Len(currentUSName) > 0 Then
                    dict(currentUSName) = nums
                End If
            End If
        End If
    Loop

    ts.Close
End Sub


Private Function PromptManualDims(ByVal title As String, _
                                  ByVal promptKey As String, _
                                  ByRef h As Double, ByRef W As Double, _
                                  ByRef d As Double, ByRef WT As Double) As Boolean
    Dim s As String, parts As Variant

    s = InputBox( _
        "Enter: Height, Width, Depth, Weight" & vbCrLf & _
        "(in, in, in, lbf) as comma-separated values." & vbCrLf & vbCrLf & _
        "Example: 78.5, 12.3, 6.0, 45.2" & vbCrLf & vbCrLf & _
        promptKey, _
        title)

    If Len(Trim$(s)) = 0 Then
        PromptManualDims = False
        Exit Function
    End If

    parts = Split(s, ",")
    If UBound(parts) < 3 Then
        MsgBox "Invalid entry. Please enter 4 numbers separated by commas.", vbExclamation, title
        PromptManualDims = False
        Exit Function
    End If

    h = CDbl(val(parts(0)))
    W = CDbl(val(parts(1)))
    d = CDbl(val(parts(2)))
    WT = CDbl(val(parts(3)))

    PromptManualDims = True
End Function

Private Sub WriteDimsToRow(ByVal lr As ListRow, _
                           ByVal hCol As Long, ByVal wCol As Long, _
                           ByVal dCol As Long, ByVal wtCol As Long, _
                           ByVal h As Double, ByVal W As Double, _
                           ByVal d As Double, ByVal WT As Double)
    lr.Range.Cells(1, hCol).Value = h
    lr.Range.Cells(1, wCol).Value = W
    lr.Range.Cells(1, dCol).Value = d
    lr.Range.Cells(1, wtCol).Value = WT
End Sub

'Build a Values= line that matches your current parsing expectations:
' - nums(2) = Weight 0"
' - last 7 values = txtValue1..7 (store H W D Generic 0 0 0)
Private Function BuildArcValuesLine(ByVal h As Double, ByVal W As Double, ByVal d As Double, ByVal WT As Double) As String
    Dim i As Long
    Dim s As String

    s = "Values="

    ' First 15 values; populate 3rd token (nums(2)) with WT
    For i = 0 To 14
        If i = 2 Then
            s = s & CStr(WT) & " "
        Else
            s = s & "0 "
        End If
    Next i

    ' Last 7 values: H W D Generic 0 0 0  (Generic stored as 0)
    s = s & CStr(h) & " " & CStr(W) & " " & CStr(d) & " 0 0 0 0"

    BuildArcValuesLine = Trim$(s)
End Function

'Append-only: always adds a new entry at the end (never replaces existing).
'IMPORTANT: Caller should skip appending for *_TIA models (per your request).
Private Sub AppendArcEntry(ByVal folderPath As String, _
                           ByVal manufacturer As String, _
                           ByVal model As String, _
                           ByVal h As Double, ByVal W As Double, ByVal d As Double, ByVal WT As Double)

    Dim fso As Object, ts As Object
    Dim filePath As String

    If Len(Trim$(manufacturer)) = 0 Or Len(Trim$(model)) = 0 Then Exit Sub
    If IsCciDatabaseName(manufacturer) Then Exit Sub

    filePath = folderPath & "\" & manufacturer & ".arc"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 8, True) ' ForAppending, create if missing

    ts.WriteLine ""
    ts.WriteLine "USName=" & model
    ts.WriteLine "SIName=" & model
    ts.WriteLine BuildArcValuesLine(h, W, d, WT)

    ts.Close
End Sub

'========================================================
' NEW: Similar-name matching helpers
'========================================================

Private Function NormalizeModel(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = UCase$(Trim$(s))
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch Like "[A-Z]") Or (ch Like "[0-9]") Then
            out = out & ch
        End If
    Next i
    NormalizeModel = out
End Function

Private Function BuildNormMap(ByVal fileDict As Object) As Object
    Dim m As Object, k As Variant, nk As String
    Set m = CreateObject("Scripting.Dictionary")
    If fileDict Is Nothing Then
        Set BuildNormMap = m
        Exit Function
    End If

    For Each k In fileDict.Keys
        nk = NormalizeModel(CStr(k))
        If Len(nk) > 0 Then
            If Not m.Exists(nk) Then m.Add nk, CStr(k)
        End If
    Next k

    Set BuildNormMap = m
End Function

Private Function Levenshtein(ByVal a As String, ByVal b As String) As Long
    Dim i As Long, j As Long
    Dim la As Long, lb As Long
    Dim cost As Long
    Dim v0() As Long, v1() As Long

    la = Len(a): lb = Len(b)
    If la = 0 Then Levenshtein = lb: Exit Function
    If lb = 0 Then Levenshtein = la: Exit Function

    ReDim v0(0 To lb)
    ReDim v1(0 To lb)

    For j = 0 To lb
        v0(j) = j
    Next j

    For i = 1 To la
        v1(0) = i
        For j = 1 To lb
            cost = IIf(Mid$(a, i, 1) = Mid$(b, j, 1), 0, 1)
            v1(j) = Application.Min(v1(j - 1) + 1, v0(j) + 1, v0(j - 1) + cost)
        Next j
        For j = 0 To lb
            v0(j) = v1(j)
        Next j
    Next i

    Levenshtein = v0(lb)
End Function

Private Function SuggestClosestModelKey( _
    ByVal fileDict As Object, _
    ByVal normMap As Object, _
    ByVal model As String, _
    ByRef suggestedKey As String) As Boolean

    Dim nm As String, k As Variant
    Dim bestKey As String, bestScore As Long, score As Long
    Dim nk As String

    suggestedKey = ""
    If fileDict Is Nothing Then Exit Function

    nm = NormalizeModel(model)
    If Len(nm) = 0 Then Exit Function

    ' 1) Fast path: normalized exact hit
    If Not normMap Is Nothing Then
        If normMap.Exists(nm) Then
            suggestedKey = CStr(normMap(nm))
            SuggestClosestModelKey = True
            Exit Function
        End If
    End If

    ' 2) "Contains" heuristic
    bestScore = 10 ^ 9
    For Each k In fileDict.Keys
        nk = NormalizeModel(CStr(k))
        If Len(nk) > 0 Then
            If (InStr(1, nk, nm, vbTextCompare) > 0) Or (InStr(1, nm, nk, vbTextCompare) > 0) Then
                score = Abs(Len(nk) - Len(nm))
                If score < bestScore Then
                    bestScore = score
                    bestKey = CStr(k)
                End If
            End If
        End If
    Next k

    If Len(bestKey) > 0 Then
        suggestedKey = bestKey
        SuggestClosestModelKey = True
        Exit Function
    End If

    ' 3) Edit distance fallback (pruned by length)
    bestScore = 10 ^ 9
    bestKey = ""
    For Each k In fileDict.Keys
        nk = NormalizeModel(CStr(k))
        If Len(nk) > 0 Then
            If Abs(Len(nk) - Len(nm)) <= 6 Then
                score = Levenshtein(nm, nk)
                If score < bestScore Then
                    bestScore = score
                    bestKey = CStr(k)
                End If
            End If
        End If
    Next k

    If Len(bestKey) > 0 And bestScore <= 6 Then
        suggestedKey = bestKey
        SuggestClosestModelKey = True
    End If
End Function

'========================================================
' Main
'========================================================

Public Sub Fill_Dimensions_From_Arc()

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False

    Dim wsCalc As Worksheet, wsDim As Worksheet
    Dim loCalc As ListObject, loDim As ListObject
    Dim lr As ListRow
    Dim rowIdx As Long

    Dim folderPath As String
    Dim manColCalc As Long, modelColCalc As Long
    Dim hColDim As Long, wColDim As Long, dColDim As Long, wtColDim As Long

    Dim manufacturer As String, model As String, altModel As String
    Dim fileDict As Object
    Dim lastFile As String
    Dim nums As Variant

    Dim heightVal As Double, widthVal As Double, depthVal As Double, weightVal As Double
    Dim baseIndex As Long

    Dim missing As Object
    Dim manualCache As Object
    Dim key As Variant
    Dim askKey As String
    Dim msg As String

    Dim normMap As Object
    Dim foundKey As String
    Dim suggestedKey As String

    Dim suggestChoiceCache As Object   ' Manufacturer|Model -> suggestedKey or "" (rejected)

    Set suggestChoiceCache = CreateObject("Scripting.Dictionary")
    Set missing = CreateObject("Scripting.Dictionary")
    Set manualCache = CreateObject("Scripting.Dictionary") ' Manufacturer|Model -> Array(H,W,D,WT)
    
    Dim dbFile As String                 ' actual .arc file chosen (full path)
    Dim dbChoiceCache As Object          ' rawManufacturer -> dbFile (full path)
    Dim lastDbFile As String

    '------------------------------------------------------
    ' 1) Folder path
    '------------------------------------------------------
    folderPath = ThisWorkbook.Worksheets("Meta").Range("Filepath_APPURT").Value
    If Len(folderPath) = 0 Or Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Appurtenance CFD folder path is not set or not valid. " & vbCrLf & _
               "Please check Meta!Filepath_APPURT.", vbExclamation, "Folder Not Found"
        GoTo CleanExit
    End If

    '------------------------------------------------------
    ' 2) Tables + columns
    '------------------------------------------------------
    Set wsCalc = ThisWorkbook.Worksheets("Discrete Loads")
    Set loCalc = wsCalc.ListObjects("Calculations_APPURT")

    Set wsDim = ThisWorkbook.Worksheets("Discrete Loads")
    Set loDim = wsDim.ListObjects("Dimensions_APPURT")

    Set dbChoiceCache = CreateObject("Scripting.Dictionary")
    lastDbFile = ""

    ' your existing clear
    wsDim.Range("AR4:AU53").ClearContents

    On Error GoTo ColError
    manColCalc = loCalc.ListColumns("Manufacturer").Index
    modelColCalc = loCalc.ListColumns("Model").Index

    hColDim = loDim.ListColumns("Height (in)").Index
    wColDim = loDim.ListColumns("Width (in)").Index
    dColDim = loDim.ListColumns("Depth (in)").Index
    wtColDim = loDim.ListColumns("Weight (lbf)").Index
    On Error GoTo CleanFail

    If loCalc.ListRows.Count <> loDim.ListRows.Count Then
        If MsgBox("Row counts for Calculations_APPURT (" & loCalc.ListRows.Count & _
                  ") and Dimensions_APPURT (" & loDim.ListRows.Count & _
                  ") do not match." & vbCrLf & vbCrLf & _
                  "Continue using row index alignment?", _
                  vbExclamation + vbOKCancel, "Row Count Mismatch") = vbCancel Then
            GoTo CleanExit
        End If
    End If

    '------------------------------------------------------
    ' 3) Loop rows
    '------------------------------------------------------
    lastFile = ""
    Set fileDict = Nothing
    Set normMap = Nothing

    For Each lr In loDim.ListRows

        rowIdx = lr.Index

        manufacturer = Trim$(loCalc.DataBodyRange.Cells(rowIdx, manColCalc).Value)
        model = Trim$(loCalc.DataBodyRange.Cells(rowIdx, modelColCalc).Value)

        If Len(manufacturer) = 0 Or Len(model) = 0 Then GoTo NextRow
        If IsCciDatabaseName(manufacturer) Then GoTo NextRow

        ' Load dict when manufacturer changes (but now we resolve the best .arc database first)
        If UCase$(manufacturer) <> UCase$(lastFile) Then
        
            Dim resolvedManUpper As String
            dbFile = ResolveArcDatabase(folderPath, manufacturer, dbChoiceCache, resolvedManUpper)
            
            ' If user picked a DB (or we resolved to a different name), overwrite Column A (Manufacturer) with ALL CAPS
            If Len(resolvedManUpper) > 0 Then
                If UCase$(manufacturer) <> resolvedManUpper Then
                    loCalc.DataBodyRange.Cells(rowIdx, manColCalc).Value = resolvedManUpper
                    manufacturer = resolvedManUpper
                End If
            End If

        
            If Len(dbFile) > 0 Then
                ' Only reload if actual DB file changed
                If StrComp(dbFile, lastDbFile, vbTextCompare) <> 0 Then
                    LoadArcFileToDict_ByFullPath dbFile, fileDict
                    lastDbFile = dbFile
                    Set normMap = BuildNormMap(fileDict)
                End If
            Else
                ' Could not resolve any db at all
                Set fileDict = Nothing
                Set normMap = Nothing
            End If
        
            lastFile = manufacturer
        End If

        ' If manually entered earlier this run, reuse (and do NOT re-append)
        askKey = UCase$(manufacturer) & "|" & UCase$(model)
        If manualCache.Exists(askKey) Then
            Dim arrCached As Variant
            arrCached = manualCache(askKey)
            WriteDimsToRow lr, hColDim, wColDim, dColDim, wtColDim, arrCached(0), arrCached(1), arrCached(2), arrCached(3)
            GoTo NextRow
        End If

        heightVal = 0: widthVal = 0: depthVal = 0: weightVal = 0
        
        If Not fileDict Is Nothing Then

            '==================================================
            ' NEW: exact match -> else suggest similar -> ask user
            '==================================================
            foundKey = ""
            
            ' First try exact match
            If fileDict.Exists(model) Then
                foundKey = model
            
            Else
                ' Apply the SAME Yes/No decision for duplicates of this missing model
                ' Keyed by manufacturer+model (case-insensitive)
                Dim decKey As String
                decKey = UCase$(manufacturer) & "|" & UCase$(model)
            
                If suggestChoiceCache.Exists(decKey) Then
                    ' If previously accepted, use the stored suggestion
                    ' If previously rejected, skip suggesting again
                    If Len(suggestChoiceCache(decKey)) > 0 Then
                        foundKey = CStr(suggestChoiceCache(decKey))
                    Else
                        foundKey = ""  ' rejected earlier
                    End If
            
                Else
                    ' No prior decision yet -> ask once
                    If SuggestClosestModelKey(fileDict, normMap, model, suggestedKey) Then
                        If MsgBox("Model not found:" & vbCrLf & vbCrLf & _
                                  manufacturer & " | " & model & vbCrLf & vbCrLf & _
                                  "Did you mean:" & vbCrLf & _
                                  "  " & suggestedKey & vbCrLf & vbCrLf & _
                                  "Use this match for ALL instances of this model?", _
                                  vbYesNo + vbQuestion, "Use Similar Model?") = vbYes Then
            
                            suggestChoiceCache.Add decKey, suggestedKey
                            foundKey = suggestedKey
                        Else
                            suggestChoiceCache.Add decKey, ""   ' remember rejection
                            foundKey = ""
                        End If
                    Else
                        ' no suggestion available -> also remember so we don't keep trying
                        suggestChoiceCache.Add decKey, ""
                        foundKey = ""
                    End If
                End If
            End If

            If Len(foundKey) > 0 Then
                ' NEW: write the matched USName back to Column B (Model) in ALL CAPS
                If StrComp(foundKey, model, vbTextCompare) <> 0 Then
                    loCalc.DataBodyRange.Cells(rowIdx, modelColCalc).Value = UCase$(foundKey)
                End If

                '-----------------------
                ' Base model lookup (using foundKey)
                '-----------------------
                nums = fileDict(foundKey)
                baseIndex = UBound(nums) - 7 + 1

                If baseIndex >= 0 Then
                    heightVal = CDbl(val(nums(baseIndex)))
                    widthVal = CDbl(val(nums(baseIndex + 1)))
                    depthVal = CDbl(val(nums(baseIndex + 2)))
                    weightVal = CDbl(val(nums(2)))
                End If

                '-----------------------
                ' Fallback to Model_TIA if dims are all zero
                ' (still based on ORIGINAL model name, not foundKey)
                '-----------------------
                If heightVal = 0 And widthVal = 0 And depthVal = 0 Then
                    altModel = model & "_TIA"
                    If fileDict.Exists(altModel) Then
                        nums = fileDict(altModel)
                        baseIndex = UBound(nums) - 7 + 1
                        If baseIndex >= 0 Then
                            heightVal = CDbl(val(nums(baseIndex)))
                            widthVal = CDbl(val(nums(baseIndex + 1)))
                            depthVal = CDbl(val(nums(baseIndex + 2)))
                            weightVal = CDbl(val(nums(2)))
                        End If
                    End If
                End If

                '-----------------------
                ' If still zero dims -> prompt manual entry
                ' (and append to .arc unless model ends with _TIA)
                '-----------------------
                If heightVal = 0 And widthVal = 0 And depthVal = 0 Then
                    If MsgBox("Missing/zero dimensions for:" & vbCrLf & vbCrLf & _
                              manufacturer & " | " & model & vbCrLf & vbCrLf & _
                              "Would you like to manually add Height/Width/Depth/Weight?", _
                              vbYesNo + vbQuestion, "Missing Dimensions") = vbYes Then

                        Dim Hm As Double, Wm As Double, Dm As Double, WTm As Double
                        If PromptManualDims("Manual Dimensions", manufacturer & " | " & model, Hm, Wm, Dm, WTm) Then

                            ' Write to sheet
                            WriteDimsToRow lr, hColDim, wColDim, dColDim, wtColDim, Hm, Wm, Dm, WTm

                            ' Cache (avoid duplicate prompts this run)
                            manualCache.Add askKey, Array(Hm, Wm, Dm, WTm)

                            ' APPEND ONLY (NO REPLACE), but OMIT _TIA models
                            If Not EndsWithTIA(model) Then
                                AppendArcEntry folderPath, manufacturer, model, Hm, Wm, Dm, WTm
                            End If

                            GoTo NextRow
                        Else
                            key = manufacturer & " | " & model & " (0 dims)"
                            If Not missing.Exists(key) Then missing.Add key, True
                            GoTo NextRow
                        End If
                    Else
                        key = manufacturer & " | " & model & " (0 dims)"
                        If Not missing.Exists(key) Then missing.Add key, True
                        GoTo NextRow
                    End If
                End If

                ' If we got valid values from base / suggested / _TIA fallback, write them
                WriteDimsToRow lr, hColDim, wColDim, dColDim, wtColDim, heightVal, widthVal, depthVal, weightVal

            Else
                '==================================================
                ' STILL NOT FOUND (no suggestion accepted)
                ' -> your existing "Model not found" manual prompt / missing log
                '==================================================
                If MsgBox("Missing dimensions for:" & vbCrLf & vbCrLf & _
                          manufacturer & " | " & model & vbCrLf & vbCrLf & _
                          "Would you like to manually add Height/Width/Depth/Weight?", _
                          vbYesNo + vbQuestion, "Missing Dimensions") = vbYes Then

                    Dim Hm2 As Double, Wm2 As Double, Dm2 As Double, WTm2 As Double
                    If PromptManualDims("Manual Dimensions", manufacturer & " | " & model, Hm2, Wm2, Dm2, WTm2) Then

                        WriteDimsToRow lr, hColDim, wColDim, dColDim, wtColDim, Hm2, Wm2, Dm2, WTm2
                        manualCache.Add askKey, Array(Hm2, Wm2, Dm2, WTm2)

                        If Not EndsWithTIA(model) Then
                            AppendArcEntry folderPath, manufacturer, model, Hm2, Wm2, Dm2, WTm2
                        End If

                    Else
                        key = manufacturer & " | " & model
                        If Not missing.Exists(key) Then missing.Add key, True
                    End If
                Else
                    key = manufacturer & " | " & model
                    If Not missing.Exists(key) Then missing.Add key, True
                End If
            End If

        Else
            key = manufacturer & " | " & model & " (database missing/empty)"
            If Not missing.Exists(key) Then missing.Add key, True
        End If

NextRow:
    Next lr

    '------------------------------------------------------
    ' 4) Summary popup
    '------------------------------------------------------
    If missing.Count > 0 Then
        msg = "Missing dimensions for these appurtenances:" & vbCrLf & vbCrLf
        For Each key In missing.Keys
            msg = msg & " • " & CStr(key) & vbCrLf
        Next key
        MsgBox msg, vbExclamation, "Missing Dimensions"
    Else
        MsgBox "Dimensions_APPURT updated from .arc CFD database.", vbInformation
    End If

    GoTo CleanExit

ColError:
    MsgBox "Could not find one or more required columns:" & vbCrLf & _
           "  Calculations_APPURT: Manufacturer, Model" & vbCrLf & _
           "  Dimensions_APPURT: Height (in), Width (in), Depth (in), Weight (lbf)", _
           vbCritical, "Column Error"
    GoTo CleanExit

CleanFail:
    MsgBox "An unexpected error occurred:" & vbCrLf & Err.Description, vbCritical, "Error"

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True

End Sub


Private Function ResolveArcDatabase(ByVal folderPath As String, _
                                    ByVal rawManufacturer As String, _
                                    ByVal cache As Object, _
                                    ByRef resolvedManufacturerUpper As String) As String
    resolvedManufacturerUpper = UCase$(Trim$(rawManufacturer))
    
    Dim key As String
    key = UCase$(Trim$(rawManufacturer))

    If cache.Exists(key) Then
        ResolveArcDatabase = CStr(cache(key))
        Exit Function
    End If

    Dim exactPath As String
    exactPath = BuildArcPath(folderPath, rawManufacturer)

    ' 1) Exact manufacturer.arc
    If FileExists(exactPath) Then
        cache(key) = exactPath
        ResolveArcDatabase = exactPath
        Exit Function
    End If

    ' 2) Token-based (split on common separators)
    Dim tokens As Collection
    Set tokens = TokenizeManufacturer(rawManufacturer)

    Dim t As Variant, candPath As String
    If tokens.Count > 1 Then
        Dim tokenHits As Collection
        Set tokenHits = New Collection

        For Each t In tokens
            candPath = BuildArcPath(folderPath, CStr(t))
            If FileExists(candPath) Then tokenHits.Add candPath
        Next t

        If tokenHits.Count = 1 Then
            ResolveArcDatabase = CStr(tokenHits(1))
            cache(key) = ResolveArcDatabase
            Exit Function

        ElseIf tokenHits.Count > 1 Then
            ResolveArcDatabase = PromptPickDb( _
                "Multiple manufacturer tokens matched for:" & vbCrLf & _
                rawManufacturer & vbCrLf & vbCrLf & _
                "Pick the database to use:", tokenHits)

            If Len(ResolveArcDatabase) > 0 Then
                cache(key) = ResolveArcDatabase
            End If
            Exit Function
        End If
    End If

    ' 3) Contains search in filenames (e.g., "RFS" -> "rfs celwave.arc")
    Dim containsHits As Collection
    Set containsHits = FindArcFilesContaining(folderPath, tokens)

    If containsHits.Count = 1 Then
        ResolveArcDatabase = CStr(containsHits(1))
        cache(key) = ResolveArcDatabase
        Exit Function

    ElseIf containsHits.Count > 1 Then
        ResolveArcDatabase = PromptPickDb( _
            "No exact database found for:" & vbCrLf & _
            rawManufacturer & vbCrLf & vbCrLf & _
            "But these databases contain the manufacturer text." & vbCrLf & _
            "Pick one:", containsHits)

        If Len(ResolveArcDatabase) > 0 Then
            cache(key) = ResolveArcDatabase
        End If
        Exit Function
    End If

'    ' 4) Nothing found -> use your picker form (db_Manufacturers)
'    ResolveArcDatabase = PickArcDbUsingPickerForm(folderPath, rawManufacturer)
'
'    If Len(ResolveArcDatabase) > 0 Then
'        cache(key) = ResolveArcDatabase
'    End If
    ' 4) Nothing found -> use your picker form (db_Manufacturers)
    ResolveArcDatabase = PickArcDbUsingPickerForm(folderPath, rawManufacturer)
    
    If Len(ResolveArcDatabase) > 0 Then
        resolvedManufacturerUpper = ArcBaseNameUpper(ResolveArcDatabase) ' <-- THIS is the key line
        cache(key) = ResolveArcDatabase
    End If
End Function


Private Function TokenizeManufacturer(ByVal rawManufacturer As String) As Collection
    Dim s As String
    s = UCase$(Trim$(rawManufacturer))

    ' Normalize separators to spaces
    s = Replace(s, ",", " ")
    s = Replace(s, "&", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, "\", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, "_", " ")

    ' Collapse multiple spaces
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    Dim parts() As String
    parts = Split(s, " ")

    Dim col As New Collection
    Dim i As Long, tok As String

    On Error Resume Next
    For i = LBound(parts) To UBound(parts)
        tok = Trim$(parts(i))
        If Len(tok) > 0 Then col.Add tok, tok ' keyed to avoid duplicates
    Next i
    On Error GoTo 0

    Set TokenizeManufacturer = col
End Function


Private Function FindArcFilesContaining(ByVal folderPath As String, ByVal tokens As Collection) As Collection
    Dim hits As New Collection
    Dim f As String, full As String
    Dim t As Variant

    f = Dir(EnsureTrailingSlash(folderPath) & "*.arc", vbNormal)
    Do While Len(f) > 0
        For Each t In tokens
            If InStr(1, f, CStr(t), vbTextCompare) > 0 Then
                full = EnsureTrailingSlash(folderPath) & f
                hits.Add full
                Exit For
            End If
        Next t
        f = Dir()
    Loop

    Set FindArcFilesContaining = hits
End Function


Private Function BuildArcPath(ByVal folderPath As String, ByVal manufacturerName As String) As String
    Dim nameOnly As String
    nameOnly = LCase$(Trim$(manufacturerName))

    ' if you already sanitize manufacturer names elsewhere, keep consistent:
    nameOnly = Replace(nameOnly, "/", " ")
    nameOnly = Replace(nameOnly, "\", " ")
    nameOnly = Replace(nameOnly, ":", " ")
    nameOnly = Replace(nameOnly, "*", " ")
    nameOnly = Replace(nameOnly, "?", " ")
    nameOnly = Replace(nameOnly, """", " ")
    nameOnly = Replace(nameOnly, "<", " ")
    nameOnly = Replace(nameOnly, ">", " ")
    nameOnly = Replace(nameOnly, "|", " ")

    Do While InStr(nameOnly, "  ") > 0
        nameOnly = Replace(nameOnly, "  ", " ")
    Loop

    BuildArcPath = EnsureTrailingSlash(folderPath) & nameOnly & ".arc"
End Function

Private Function EnsureTrailingSlash(ByVal p As String) As String
    If Len(p) = 0 Then
        EnsureTrailingSlash = ""
    ElseIf Right$(p, 1) = "\" Or Right$(p, 1) = "/" Then
        EnsureTrailingSlash = p
    Else
        EnsureTrailingSlash = p & "\"
    End If
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Len(fullPath) > 0 And Len(Dir(fullPath, vbNormal)) > 0)
End Function

Private Function ListAllArcFiles(ByVal folderPath As String) As Collection
    Dim col As New Collection
    Dim f As String
    f = Dir(EnsureTrailingSlash(folderPath) & "*.arc", vbNormal)
    Do While Len(f) > 0
        col.Add EnsureTrailingSlash(folderPath) & f
        f = Dir()
    Loop
    Set ListAllArcFiles = col
End Function


Private Function PromptPickDb(ByVal titleText As String, ByVal options As Collection) As String
    Dim msg As String, i As Long, opt As String
    msg = titleText & vbCrLf & vbCrLf

    ' Keep message readable
    Dim maxShow As Long: maxShow = 30
    For i = 1 To options.Count
        If i > maxShow Then
            msg = msg & "  ... (" & (options.Count - maxShow) & " more not shown)" & vbCrLf
            Exit For
        End If
        opt = CStr(options(i))
        msg = msg & "  " & i & ") " & Mid$(opt, InStrRev(opt, "\") + 1) & vbCrLf
    Next i

    msg = msg & vbCrLf & "Enter a number (1-" & options.Count & "), or leave blank to skip:"

    Dim s As String
    s = InputBox(msg, "Select Database")

    s = Trim$(s)

    ' Cancel or blank -> skip
    If Len(s) = 0 Then
        PromptPickDb = ""
        Exit Function
    End If

    ' Validate numeric
    If Not IsNumeric(s) Then
        MsgBox "Please enter a number from 1 to " & options.Count & ".", vbExclamation, "Invalid Selection"
        PromptPickDb = ""
        Exit Function
    End If

    Dim idx As Long
    idx = CLng(val(s))

    If idx < 1 Or idx > options.Count Then
        MsgBox "Selection out of range. Enter 1 to " & options.Count & ".", vbExclamation, "Invalid Selection"
        PromptPickDb = ""
        Exit Function
    End If

    PromptPickDb = CStr(options(idx))
End Function

Private Function PickArcDbUsingPickerForm(ByVal folderPath As String, ByVal rawManufacturer As String) As String
    Dim f As db_Manufacturers
    Set f = New db_Manufacturers

    f.BeginPick folderPath, _
        "No database found for: " & rawManufacturer & vbCrLf & vbCrLf & _
        "Select an alternate .arc database:"

    f.Show vbModal

    If f.PickerCancelled Then
        PickArcDbUsingPickerForm = ""
    Else
        PickArcDbUsingPickerForm = f.PickedArcFullPath
    End If

    Unload f
    Set f = Nothing
End Function


Sub FillInTheBlanks()
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


Private Function ArcBaseNameUpper(ByVal arcFullPath As String) As String
    ' "C:\...\rfs celwave.arc" -> "RFS CELWAVE"
    Dim fn As String
    fn = Mid$(arcFullPath, InStrRev(arcFullPath, "\") + 1)
    If LCase$(Right$(fn, 4)) = ".arc" Then fn = Left$(fn, Len(fn) - 4)
    ArcBaseNameUpper = UCase$(Trim$(fn))
End Function

