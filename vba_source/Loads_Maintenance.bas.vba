Attribute VB_Name = "Loads_Maintenance"
Option Explicit

Sub Maintenance_Loads()

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    
    Dim LmFirstRow As Long, LvFirstRow As Long
    Dim LmLabelCol As Long, LvLabelCol As Long
    Dim LmDirCol As Long, LvDirCol As Long
    Dim LmMagCol As Long, LvMagCol As Long
    Dim LmLocCol As Long, LvLocCol As Long
    
    Dim r As Long, lastRow As Long
    Dim lmCount As Long, lvCount As Long
    Dim baseCol As Long, i As Long
    
    '--- sheet references
    Set wsSrc = shMaintenance_Loads
    Set wsDst = shPointLoadTables
    
    '--- source sheet layout
    LmFirstRow = 7       ' Lm begins at B7
    LvFirstRow = 7       ' Lv begins at G7 (your correction)

    LmLabelCol = 2       ' Column B
    LvLabelCol = 7       ' Column G

    LmDirCol = 3         ' Direction Y column
    LvDirCol = 8         ' Direction Y column
    
    LmMagCol = 4         ' Magnitude
    LvMagCol = 9         ' Magnitude
    
    LmLocCol = 5         ' Location %
    LvLocCol = 10        ' Location %

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .StatusBar = "Building maintenance point load tables..."
    End With

' ========================================================
' 1) Count Lm rows
' ========================================================
    r = LmFirstRow
    Do While wsSrc.Cells(r, LmLabelCol).Value <> ""
        lmCount = lmCount + 1
        r = r + 1
    Loop

' ========================================================
' 2) Count Lv rows
' ========================================================
    r = LvFirstRow
    Do While wsSrc.Cells(r, LvLabelCol).Value <> ""
        lvCount = lvCount + 1
        r = r + 1
    Loop

    If lmCount = 0 And lvCount = 0 Then
        MsgBox "No maintenance loads found.", vbInformation
        GoTo CleanExit
    End If


' ========================================================
' 3) Clear destination area
' ========================================================
    wsDst.Range("GD4:HP200").ClearContents


' ========================================================
' 4) Write Lm tables first (grouped)
' ========================================================
    baseCol = wsDst.Range("GD4").Column

    For i = 1 To lmCount

        Dim srcRow As Long
        Dim dstCol As Long

        srcRow = LmFirstRow + (i - 1)
        dstCol = baseCol + (i - 1) * 5     ' 4 columns + 1 separator

        ' Write only the title (no column headers)
        wsDst.Cells(2, dstCol).Value = "Maintenance " & " Lm_" & i

        ' Write first load row (Member Label, Y, Magnitude, Location)
        wsDst.Cells(4, dstCol).Value = wsSrc.Cells(srcRow, LmLabelCol).Value
        wsDst.Cells(4, dstCol + 1).Value = wsSrc.Cells(srcRow, LmDirCol).Value
        wsDst.Cells(4, dstCol + 2).Value = wsSrc.Cells(srcRow, LmMagCol).Value
        wsDst.Cells(4, dstCol + 3).Value = wsSrc.Cells(srcRow, LmLocCol).Value

    Next i


' ========================================================
' 5) Write Lv tables (after all Lm tables)
' ========================================================
    Dim lvBaseCol As Long
    lvBaseCol = baseCol + lmCount * 5

    For i = 1 To lvCount

        Dim srcRowLv As Long
        Dim dstColLv As Long

        srcRowLv = LvFirstRow + (i - 1)
        dstColLv = lvBaseCol + (i - 1) * 5

        ' Write title only
        wsDst.Cells(2, dstColLv).Value = "Maintenance " & " Lv_" & i

        ' Write the load row
        wsDst.Cells(4, dstColLv).Value = wsSrc.Cells(srcRowLv, LvLabelCol).Value
        wsDst.Cells(4, dstColLv + 1).Value = wsSrc.Cells(srcRowLv, LvDirCol).Value
        wsDst.Cells(4, dstColLv + 2).Value = wsSrc.Cells(srcRowLv, LvMagCol).Value
        wsDst.Cells(4, dstColLv + 3).Value = wsSrc.Cells(srcRowLv, LvLocCol).Value

    Next i


CleanExit:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = "Ready"
    End With

End Sub
