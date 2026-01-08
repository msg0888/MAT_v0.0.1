Attribute VB_Name = "Loads_Point"
Sub LoadsPoint()

Sheets("Point Load Tables").Range("A4:GB4").ClearContents
    Sheets("Point Load Tables").Range("4:1048576").ClearContents

    ' Alpha
    Wind_LoadsPoint 14, 16, 187, 18, 19
    IceWind_LoadsPoint 14, 16, 251, 18, 19
    Gravity_LoadsPoint 14, 16, 18, 19
    SeismicEhZ_LoadsPoint 14, 16, 18, 19
    SeismicEhX_LoadsPoint 14, 16, 18, 19
    ' Beta
    Wind_LoadsPoint 22, 24, 203, 26, 27
    IceWind_LoadsPoint 22, 24, 267, 26, 27
    Gravity_LoadsPoint 22, 24, 26, 27
    SeismicEhZ_LoadsPoint 22, 24, 26, 27
    SeismicEhX_LoadsPoint 22, 24, 26, 27
    ' Gamma
    Wind_LoadsPoint 30, 32, 219, 34, 35
    IceWind_LoadsPoint 30, 32, 283, 34, 35
    Gravity_LoadsPoint 30, 32, 34, 35
    SeismicEhZ_LoadsPoint 30, 32, 34, 35
    SeismicEhX_LoadsPoint 30, 32, 34, 35
    ' Delta
    Wind_LoadsPoint 38, 40, 235, 42, 43
    IceWind_LoadsPoint 38, 40, 299, 42, 43
    Gravity_LoadsPoint 38, 40, 42, 43
    SeismicEhZ_LoadsPoint 38, 40, 42, 43
    SeismicEhX_LoadsPoint 38, 40, 42, 43
    
    Call Maintenance_Loads
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
End Sub


Sub Wind_LoadsPoint(posTopCol As Long, posBtmCol As Long, baseForceCol As Long, locTopCol As Long, locBtmCol As Long)
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim i As Long, j As Long
    Dim posTop As String, posBtm As String
    Dim locTop As Variant, locBtm As Variant
    Dim rawForce As Variant, adjustedForce As Double
    Dim direction As Variant
    Dim Directions As Variant
    Dim Phi() As Variant, CosPhi(15) As Double, SinPhi(15) As Double
    Dim Pi As Double: Pi = WorksheetFunction.Pi()
    Dim forceStartCol As Long
    Dim srcData As Variant
    Dim outputData() As Variant
    Dim outIdx As Long
    Dim cv As String
    Dim AType As Variant
    Dim coeff As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False

    Directions = Array("Z", "X", "Mz", "Mx")
    Phi = Array(0, 30, 45, 60, 90, 120, 135, 150, 180, 210, 225, 240, 270, 300, 315, 330)
       
    ' Precompute Sin and Cos values
    For j = 0 To 15
        CosPhi(j) = Cos(Pi / 180 * Phi(j))
        SinPhi(j) = Sin(Pi / 180 * Phi(j))
    Next j

    Set wsSource = ThisWorkbook.Sheets("Discrete Loads")
    Set wsTarget = ThisWorkbook.Sheets("Point Load Tables")

    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).row
    outputRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).row + 1

    ' Read entire source data into memory
    srcData = wsSource.Range(wsSource.Cells(4, 1), wsSource.Cells(lastRow, baseForceCol + 330)).Value

    ' Estimate output rows (max 2 * directions * 16 columns per input row)
    Const MaxCols As Long = 80
    ReDim outputData(1 To (lastRow - 4) * 8, 1 To MaxCols)

    outIdx = 1

For i = 1 To UBound(srcData, 1)
    posTop = srcData(i, posTopCol)
    posBtm = srcData(i, posBtmCol)
    locTop = srcData(i, locTopCol)
    locBtm = srcData(i, locBtmCol)
    AType = Trim$(srcData(i, 4))             ' <-- Column D = Type

    For Each direction In Directions
        Select Case direction
            Case "Z", "X":  forceStartCol = baseForceCol
            Case "Mx":      forceStartCol = baseForceCol + (315 - 187)
            Case "Mz":      forceStartCol = baseForceCol + (443 - 187)
        End Select

        '------------------ TOP ------------------
        If posTop <> "" And IsNumeric(locTop) Then
            For j = 0 To 15
                ' If Type = Dish, pull CA from TIA-222-H; else use existing table value

                coeff = CoeffForDirection(CStr(direction))
                If IsSupportedType(CStr(AType)) And Len(coeff) > 0 Then
                    rawForce = GetTIAFactor(CStr(AType), coeff, CDbl(Phi(j)))
                Else
                    rawForce = srcData(i, forceStartCol + j)
                End If

                Select Case direction
                    Case "Z", "Mz": adjustedForce = -CDbl(rawForce) * CosPhi(j)
                    Case "X", "Mx": adjustedForce = CDbl(rawForce) * SinPhi(j)
                End Select

                outputData(outIdx, 1 + j * 5) = posTop
                outputData(outIdx, 2 + j * 5) = direction
                outputData(outIdx, 3 + j * 5) = adjustedForce
                outputData(outIdx, 4 + j * 5) = locTop
            Next j
            outIdx = outIdx + 1
        End If

        '---------------- BOTTOM -----------------
        If posBtm <> "" And IsNumeric(locBtm) Then
            For j = 0 To 15

                coeff = CoeffForDirection(CStr(direction))
               If IsSupportedType(CStr(AType)) And Len(coeff) > 0 Then
                    rawForce = GetTIAFactor(CStr(AType), coeff, CDbl(Phi(j)))
                Else
                    rawForce = srcData(i, forceStartCol + j)
                End If

                Select Case direction
                    Case "Z", "Mz": adjustedForce = -CDbl(rawForce) * CosPhi(j)
                    Case "X", "Mx": adjustedForce = CDbl(rawForce) * SinPhi(j)
                End Select

                outputData(outIdx, 1 + j * 5) = posBtm
                outputData(outIdx, 2 + j * 5) = direction
                outputData(outIdx, 3 + j * 5) = adjustedForce
                outputData(outIdx, 4 + j * 5) = locBtm
            Next j
            outIdx = outIdx + 1
        End If
    
    Next direction
Next i

    ' Write output in one go
    If outIdx > 1 Then
        Dim resultArray As Variant
    
        On Error GoTo Cleanup ' trap if Index fails
    
        resultArray = Application.Index(outputData, Evaluate("ROW(1:" & outIdx - 1 & ")"), Evaluate("COLUMN(1:" & MaxCols & ")"))
    
        wsTarget.Range(wsTarget.Cells(outputRow, 1), _
                       wsTarget.Cells(outputRow + UBound(resultArray, 1) - 1, 80)).Value = resultArray
    End If
    
    GoTo Done
    
Cleanup:
        MsgBox "No valid rows to output or output array failed.", vbExclamation
    
Done:
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'    Application.DisplayStatusBar = True
End Sub


Sub IceWind_LoadsPoint(posTopCol As Long, posBtmCol As Long, baseForceCol As Long, locTopCol As Long, locBtmCol As Long)
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim i As Long, j As Long
    Dim posTop As String, posBtm As String
    Dim locTop As Variant, locBtm As Variant
    Dim rawForce As Variant, adjustedForce As Double
    Dim direction As Variant
    Dim Directions As Variant
    Dim Phi() As Variant, CosPhi(15) As Double, SinPhi(15) As Double
    Dim Pi As Double: Pi = WorksheetFunction.Pi()
    Dim forceStartCol As Long
    Dim srcData As Variant
    Dim outputData() As Variant
    Dim outIdx As Long
    Const MaxCols As Long = 80
    Const OutputStartCol As Long = 81 ' Column CC
    Dim AType As Variant
    Dim coeff As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False

    Directions = Array("Z", "X", "Mz", "Mx")
    Phi = Array(0, 30, 45, 60, 90, 120, 135, 150, 180, 210, 225, 240, 270, 300, 315, 330)
    
    ' Precompute Sin and Cos values
    For j = 0 To 15
        CosPhi(j) = Cos(Pi / 180 * Phi(j))
        SinPhi(j) = Sin(Pi / 180 * Phi(j))
    Next j

    Set wsSource = ThisWorkbook.Sheets("Discrete Loads")
    Set wsTarget = ThisWorkbook.Sheets("Point Load Tables")

    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).row
    outputRow = wsTarget.Cells(wsTarget.Rows.Count, OutputStartCol).End(xlUp).row + 1
    If outputRow < 4 Then outputRow = 4

    ' Read entire source data into memory
    srcData = wsSource.Range(wsSource.Cells(4, 1), wsSource.Cells(lastRow, baseForceCol + 330)).Value

    ' Estimate output rows (max 2 * directions * 16 columns per input row)
    ReDim outputData(1 To (lastRow - 4) * 8, 1 To MaxCols)

    outIdx = 1

    For i = 1 To UBound(srcData, 1)
        posTop = srcData(i, posTopCol)
        posBtm = srcData(i, posBtmCol)
        locTop = srcData(i, locTopCol)
        locBtm = srcData(i, locBtmCol)

        For Each direction In Directions
            Select Case direction
                Case "Z", "X": forceStartCol = baseForceCol
                Case "Mx": forceStartCol = baseForceCol + (379 - 251)
                Case "Mz": forceStartCol = baseForceCol + (507 - 251)
            End Select

        '------------------ TOP ------------------
        If posTop <> "" And IsNumeric(locTop) Then
            For j = 0 To 15
                ' If Type = Dish, pull CA from TIA-222-H; else use existing table value

                coeff = CoeffForDirection(CStr(direction))
                If IsSupportedType(CStr(AType)) And Len(coeff) > 0 Then
                    rawForce = GetTIAFactor(CStr(AType), coeff, CDbl(Phi(j)))
                Else
                    rawForce = srcData(i, forceStartCol + j)
                End If

                Select Case direction
                    Case "Z", "Mz": adjustedForce = -CDbl(rawForce) * CosPhi(j)
                    Case "X", "Mx": adjustedForce = CDbl(rawForce) * SinPhi(j)
                End Select

                outputData(outIdx, 1 + j * 5) = posTop
                outputData(outIdx, 2 + j * 5) = direction
                outputData(outIdx, 3 + j * 5) = adjustedForce
                outputData(outIdx, 4 + j * 5) = locTop
            Next j
            outIdx = outIdx + 1
        End If

        '---------------- BOTTOM -----------------
        If posBtm <> "" And IsNumeric(locBtm) Then
            For j = 0 To 15

                coeff = CoeffForDirection(CStr(direction))
               If IsSupportedType(CStr(AType)) And Len(coeff) > 0 Then
                    rawForce = GetTIAFactor(CStr(AType), coeff, CDbl(Phi(j)))
                Else
                    rawForce = srcData(i, forceStartCol + j)
                End If

                Select Case direction
                    Case "Z", "Mz": adjustedForce = -CDbl(rawForce) * CosPhi(j)
                    Case "X", "Mx": adjustedForce = CDbl(rawForce) * SinPhi(j)
                End Select

                outputData(outIdx, 1 + j * 5) = posBtm
                outputData(outIdx, 2 + j * 5) = direction
                outputData(outIdx, 3 + j * 5) = adjustedForce
                outputData(outIdx, 4 + j * 5) = locBtm
            Next j
            outIdx = outIdx + 1
        End If
        
        Next direction
    Next i

    ' Write output starting at column CC (81)
    If outIdx > 1 Then
        wsTarget.Range(wsTarget.Cells(outputRow, OutputStartCol), _
                       wsTarget.Cells(outputRow + outIdx - 2, OutputStartCol + MaxCols - 1)).Value = _
            Application.Index(outputData, Evaluate("ROW(1:" & outIdx - 1 & ")"), Evaluate("COLUMN(81:160)"))
    End If

'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'    Application.DisplayStatusBar = True
End Sub


Sub Gravity_LoadsPoint(posTopCol As Long, posBtmCol As Long, locTopCol As Long, locBtmCol As Long)
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim i As Long, j As Long
    Dim posTop As String, posBtm As String
    Dim locTop As Variant, locBtm As Variant
    Dim forceVal As Variant, rawForce As Variant
    Dim divisor As Variant, offset As Variant
    Dim wtCol47 As Variant, iwtCol52 As Variant
    Dim direction As Variant, Directions As Variant
    Dim cv_F As Double, cv_M As Double
    Dim Ev As Double, Eh As Double

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Directions = Array("Y", "My")

    Select Case shCode.Range("RISA3D.Unit.Force")
        Case "kip": cv_F = 1 / 1000
        Case "lbf": cv_F = 1
        Case "mt": cv_F = 1 / 2204.623
        Case "kN": cv_F = 1 / 224.80894
        Case "N": cv_F = 1 / 0.224809
        Case "kg": cv_F = 1 / 2.204623
    End Select

    Select Case shCode.Range("RISA3D.Unit.Moment")
        Case "kip-ft": cv_M = 1 / 1000
        Case "kip-in": cv_M = 12 / 1000
        Case "lbf-ft": cv_M = 1
        Case "lbf-in": cv_M = 12
    End Select

    Set wsSource = ThisWorkbook.Sheets("Discrete Loads")
    Set wsTarget = ThisWorkbook.Sheets("Point Load Tables")
    Ev = Evaluate("Ev")
    Eh = Evaluate("Eh")

    lastRow = wsSource.Cells(wsSource.Rows.Count, 160).End(xlUp).row
    outputRow = wsTarget.Cells(wsTarget.Rows.Count, 161).End(xlUp).row + 1
    If outputRow < 4 Then outputRow = 4

    For i = 4 To lastRow
        posTop = wsSource.Cells(i, posTopCol).Value
        posBtm = wsSource.Cells(i, posBtmCol).Value
        locTop = wsSource.Cells(i, locTopCol).Value
        locBtm = wsSource.Cells(i, locBtmCol).Value
        divisor = wsSource.Cells(i, 6).Value   'Brackets
        offset = wsSource.Cells(i, 9).Value / 12 'Horizontal Offset
        wtCol47 = wsSource.Cells(i, 47).Value 'Weight
        iwtCol52 = wsSource.Cells(i, 52).Value 'Ice Weight

        If IsNumeric(divisor) And divisor <> 0 Then
            ' Top
            If posTop <> "" And IsNumeric(locTop) Then
                For Each direction In Directions
                    For j = 0 To 2
                        Select Case j
                            Case 0: rawForce = wtCol47 / divisor
                            Case 1: rawForce = iwtCol52 / divisor
                            Case 2: rawForce = wtCol47 * Ev / divisor
                        End Select
                        wsTarget.Cells(outputRow, 161 + j * 5).Value = posTop
                        wsTarget.Cells(outputRow, 162 + j * 5).Value = direction
                        Select Case direction
                            Case "Y": wsTarget.Cells(outputRow, 163 + j * 5).Value = -rawForce * cv_F
                            Case "My": wsTarget.Cells(outputRow, 163 + j * 5).Value = -rawForce * offset * cv_M
                        End Select
                        wsTarget.Cells(outputRow, 164 + j * 5).Value = locTop
                    Next j
                    outputRow = outputRow + 1
                Next direction
            End If

            ' Bottom
            If posBtm <> "" And IsNumeric(locBtm) Then
                For Each direction In Directions
                    For j = 0 To 2
                        Select Case j
                            Case 0: rawForce = wtCol47 / divisor
                            Case 1: rawForce = iwtCol52 / divisor
                            Case 2: rawForce = wtCol47 * Ev / divisor
                        End Select
                        wsTarget.Cells(outputRow, 161 + j * 5).Value = posBtm
                        wsTarget.Cells(outputRow, 162 + j * 5).Value = direction
                        Select Case direction
                            Case "Y": wsTarget.Cells(outputRow, 163 + j * 5).Value = -rawForce * cv_F
                            Case "My": wsTarget.Cells(outputRow, 163 + j * 5).Value = -rawForce * offset * cv_M
                        End Select
                        wsTarget.Cells(outputRow, 164 + j * 5).Value = locBtm
                    Next j
                    outputRow = outputRow + 1
                Next direction
            End If
        End If
    Next i

'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'    Application.DisplayStatusBar = True
End Sub


Sub SeismicEhZ_LoadsPoint(posTopCol As Long, posBtmCol As Long, locTopCol As Long, locBtmCol As Long)
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim i As Long, j As Long
    Dim posTop As String, posBtm As String
    Dim locTop As Variant, locBtm As Variant
    Dim forceVal As Variant, rawForce As Variant
    Dim divisor As Variant, offset As Variant
    Dim wtCol47 As Variant, iwtCol52 As Variant
    Dim direction As Variant, Directions As Variant
    Dim cv_F As Double, cv_M As Double
    Dim Ev As Double, Eh As Double

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Directions = Array("Z", "Mz")

    Select Case shCode.Range("RISA3D.Unit.Force")
        Case "kip": cv_F = 1 / 1000
        Case "lbf": cv_F = 1
        Case "mt": cv_F = 1 / 2204.623
        Case "kN": cv_F = 1 / 224.80894
        Case "N": cv_F = 1 / 0.224809
        Case "kg": cv_F = 1 / 2.204623
    End Select

    Select Case shCode.Range("RISA3D.Unit.Moment")
        Case "kip-ft": cv_M = 1 / 1000
        Case "kip-in": cv_M = 12 / 1000
        Case "lbf-ft": cv_M = 1
        Case "lbf-in": cv_M = 12
    End Select

    Set wsSource = ThisWorkbook.Sheets("Discrete Loads")
    Set wsTarget = ThisWorkbook.Sheets("Point Load Tables")
    Ev = Evaluate("Ev")
    Eh = Evaluate("Eh")

    lastRow = wsSource.Cells(wsSource.Rows.Count, 160).End(xlUp).row
    outputRow = wsTarget.Cells(wsTarget.Rows.Count, 176).End(xlUp).row + 1
    If outputRow < 4 Then outputRow = 4

    For i = 4 To lastRow
        posTop = wsSource.Cells(i, posTopCol).Value
        posBtm = wsSource.Cells(i, posBtmCol).Value
        locTop = wsSource.Cells(i, locTopCol).Value
        locBtm = wsSource.Cells(i, locBtmCol).Value
        divisor = wsSource.Cells(i, 6).Value   'Brackets
        offset = wsSource.Cells(i, 9).Value / 12 'Horizontal Offset
        wtCol47 = wsSource.Cells(i, 47).Value 'Weight
        iwtCol52 = wsSource.Cells(i, 52).Value 'Ice Weight

        If IsNumeric(divisor) And divisor <> 0 Then
            ' Top
            If posTop <> "" And IsNumeric(locTop) Then
                For Each direction In Directions
                    For j = 0 To 1
                        Select Case j
                            Case 0: rawForce = wtCol47 * Eh / divisor
                            Case 1: rawForce = wtCol47 * Eh / divisor
                        End Select
                        wsTarget.Cells(outputRow, 176).Value = posTop
                        wsTarget.Cells(outputRow, 177).Value = direction
                        Select Case direction
                            Case "Z": wsTarget.Cells(outputRow, 178).Value = -rawForce * cv_F
                            Case "Mz": wsTarget.Cells(outputRow, 178).Value = -rawForce * offset * cv_M
                        End Select
                        wsTarget.Cells(outputRow, 179).Value = locTop
                    Next j
                    outputRow = outputRow + 1
                Next direction
            End If

            ' Bottom
            If posBtm <> "" And IsNumeric(locBtm) Then
                For Each direction In Directions
                    For j = 0 To 1
                        Select Case j
                            Case 0: rawForce = wtCol47 * Eh / divisor
                            Case 1: rawForce = wtCol47 * Eh / divisor
                        End Select
                        wsTarget.Cells(outputRow, 176).Value = posTop
                        wsTarget.Cells(outputRow, 177).Value = direction
                        Select Case direction
                            Case "Z": wsTarget.Cells(outputRow, 178).Value = -rawForce * cv_F
                            Case "Mz": wsTarget.Cells(outputRow, 178).Value = -rawForce * offset * cv_M
                        End Select
                        wsTarget.Cells(outputRow, 179).Value = locTop
                    Next j
                    outputRow = outputRow + 1
                Next direction
            End If
        End If
    Next i
    
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'    Application.DisplayStatusBar = True
End Sub


Sub SeismicEhX_LoadsPoint(posTopCol As Long, posBtmCol As Long, locTopCol As Long, locBtmCol As Long)
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim i As Long, j As Long
    Dim posTop As String, posBtm As String
    Dim locTop As Variant, locBtm As Variant
    Dim forceVal As Variant, rawForce As Variant
    Dim divisor As Variant, offset As Variant
    Dim wtCol47 As Variant, iwtCol52 As Variant
    Dim direction As Variant, Directions As Variant
    Dim cv_F As Double, cv_M As Double
    Dim Ev As Double, Eh As Double

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Directions = Array("X", "Mx")

    Select Case shCode.Range("RISA3D.Unit.Force")
        Case "kip": cv_F = 1 / 1000
        Case "lbf": cv_F = 1
        Case "mt": cv_F = 1 / 2204.623
        Case "kN": cv_F = 1 / 224.80894
        Case "N": cv_F = 1 / 0.224809
        Case "kg": cv_F = 1 / 2.204623
    End Select

    Select Case shCode.Range("RISA3D.Unit.Moment")
        Case "kip-ft": cv_M = 1 / 1000
        Case "kip-in": cv_M = 12 / 1000
        Case "lbf-ft": cv_M = 1
        Case "lbf-in": cv_M = 12
    End Select

    Set wsSource = ThisWorkbook.Sheets("Discrete Loads")
    Set wsTarget = ThisWorkbook.Sheets("Point Load Tables")
    Ev = Evaluate("Ev")
    Eh = Evaluate("Eh")

    lastRow = wsSource.Cells(wsSource.Rows.Count, 160).End(xlUp).row
    outputRow = wsTarget.Cells(wsTarget.Rows.Count, 181).End(xlUp).row + 1
    If outputRow < 4 Then outputRow = 4

    For i = 4 To lastRow
        posTop = wsSource.Cells(i, posTopCol).Value
        posBtm = wsSource.Cells(i, posBtmCol).Value
        locTop = wsSource.Cells(i, locTopCol).Value
        locBtm = wsSource.Cells(i, locBtmCol).Value
        divisor = wsSource.Cells(i, 6).Value   'Brackets
        offset = wsSource.Cells(i, 9).Value / 12 'Horizontal Offset
        wtCol47 = wsSource.Cells(i, 47).Value 'Weight
        iwtCol52 = wsSource.Cells(i, 52).Value 'Ice Weight

        If IsNumeric(divisor) And divisor <> 0 Then
            ' Top
            If posTop <> "" And IsNumeric(locTop) Then
                For Each direction In Directions
                    For j = 0 To 1
                        Select Case j
                            Case 0: rawForce = wtCol47 * Eh / divisor
                            Case 1: rawForce = wtCol47 * Eh / divisor
                        End Select
                        wsTarget.Cells(outputRow, 181).Value = posTop
                        wsTarget.Cells(outputRow, 182).Value = direction
                        Select Case direction
                            Case "X": wsTarget.Cells(outputRow, 183).Value = rawForce * cv_F
                            Case "Mx": wsTarget.Cells(outputRow, 183).Value = rawForce * offset * cv_M
                        End Select
                        wsTarget.Cells(outputRow, 184).Value = locTop
                    Next j
                    outputRow = outputRow + 1
                Next direction
            End If

            ' Bottom
            If posBtm <> "" And IsNumeric(locBtm) Then
                For Each direction In Directions
                    For j = 0 To 1
                        Select Case j
                            Case 0: rawForce = wtCol47 * Eh / divisor
                            Case 1: rawForce = wtCol47 * Eh / divisor
                        End Select
                        wsTarget.Cells(outputRow, 181).Value = posTop
                        wsTarget.Cells(outputRow, 182).Value = direction
                        Select Case direction
                            Case "X": wsTarget.Cells(outputRow, 183).Value = rawForce * cv_F
                            Case "Mx": wsTarget.Cells(outputRow, 183).Value = rawForce * offset * cv_M
                        End Select
                        wsTarget.Cells(outputRow, 184).Value = locTop
                    Next j
                    outputRow = outputRow + 1
                Next direction
            End If
        End If
    Next i
    
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'    Application.DisplayStatusBar = True
End Sub


Private Function GetTIAFactor(typeName As String, coeffName As String, phiDeg As Double) As Variant
    'Replicates:
    ' =INDEX('TIA-222-H'!$U$5:$AF$45,
    '        MATCH(phi, 'TIA-222-H'!$T$5:$T$45, 0),
    '        MATCH(1, ('TIA-222-H'!$U$3:$AF$3=typeName)*('TIA-222-H'!$U$4:$AF$4=coeffName), 0))

    Dim ws As Worksheet
    Dim rowIdx As Variant, colIdx As Variant
    Dim topHdr As Range, subHdr As Range, dataRng As Range
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("TIA-222-H")
    Set topHdr = ws.Range("U3:AF3")   ' Dish / Radome / Shroud / Grid
    Set subHdr = ws.Range("U4:AF4")   ' CA / CS / CM
    Set dataRng = ws.Range("U5:AF45") ' values
    ' Row header for phi (deg)
    rowIdx = Application.Match(phiDeg, ws.Range("T5:T45"), 0)
    If IsError(rowIdx) Then
        GetTIAFactor = CVErr(xlErrNA)
        Exit Function
    End If

    colIdx = CVErr(xlErrNA)
    For i = 1 To topHdr.Columns.Count
        If StrComp(CStr(topHdr.Cells(1, i).Value), typeName, vbTextCompare) = 0 _
           And StrComp(CStr(subHdr.Cells(1, i).Value), coeffName, vbTextCompare) = 0 Then
            colIdx = i
            Exit For
        End If
    Next i

    If IsError(colIdx) Then
        GetTIAFactor = CVErr(xlErrNA)
    Else
        GetTIAFactor = dataRng.Cells(rowIdx, colIdx).Value
    End If
End Function


Private Function CoeffForDirection(ByVal dirText As String) As String
    Select Case UCase$(dirText)
        Case "Z":    CoeffForDirection = "CA"
        Case "X":    CoeffForDirection = "CS"
        Case "MX", "MZ": CoeffForDirection = "CM"
        Case Else:   CoeffForDirection = ""
    End Select
End Function


Private Function IsSupportedType(ByVal typeText As String) As Boolean
    Select Case UCase$(typeText)
        Case "DISH", "RADOME", "SHROUD", "GRID"
            IsSupportedType = True
        Case Else
            IsSupportedType = False
    End Select
End Function
