Attribute VB_Name = "Buttons"
Sub MAT_Support()
    frm_Support.Show
End Sub

Sub ButtonID_Click()
    ButtonID.Show
End Sub

Sub APPURT_Click()
    db_DiscreteLoads.Show
End Sub

Sub DISH_Click()
    db_Dishes.Show
End Sub

Sub RFDS_Click()
    db_RFDS.Show
End Sub

Sub windDirections_Click()
    windDirections.Show
End Sub

Sub ClearEnvelope()

    With shResults
        .Unprotect
        .Range("A4:R99999").ClearContents
    End With

End Sub

Sub PasteEnvelope()

Call ClearEnvelope

With shResults
    .Unprotect
    .Range("A4").Select
    ActiveSheet.Paste
End With

End Sub

Sub PrintPDF()
Dim Answer As String
Dim PDFName As String

    PDFName = Range("Carrier_SiteNumber") & "_" & Range("Carrier") & "_" & Format(Date, "YYYY-MM-DD") & "_" & Format(Now, "hh-mm-ss") & ".pdf"

        shCode.ExportAsFixedFormat xlTypePDF, ThisWorkbook.path & "\0a. Code_" & PDFName, , , False, 1, 4, True
        shCode.Select
        
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
'        "C:\Users\mgirgis\OneDrive - Network Building & Consulting, LLC\Documents\~Development\Tools_Mounts\Mount Analysis\Versions\MAT_v3.X\MAT_v3.1.0\NB+C MAT v3.1.0.3.pdf" _
'        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'        :=False, OpenAfterPublish:=True
End Sub


Sub GoToAlpha()
With shDiscrete_Loads
Application.ScreenUpdating = False
    ActiveWindow.SmallScroll ToLeft:=48
    ActiveWindow.SmallScroll ToRight:=5 + 8 * 0
    .Range("L4").Select
Application.ScreenUpdating = True
End With
End Sub


Sub GoToBeta()
With shDiscrete_Loads
Application.ScreenUpdating = False
    ActiveWindow.SmallScroll ToLeft:=48
    ActiveWindow.SmallScroll ToRight:=5 + 8 * 1
    .Range("T4").Select
Application.ScreenUpdating = True
End With
End Sub


Sub GoToGamma()
With shDiscrete_Loads
Application.ScreenUpdating = False
    ActiveWindow.SmallScroll ToLeft:=48
    ActiveWindow.SmallScroll ToRight:=5 + 8 * 2
    .Range("AB4").Select
Application.ScreenUpdating = True
End With
End Sub


Sub GoToDelta()
With shDiscrete_Loads
Application.ScreenUpdating = False
    ActiveWindow.SmallScroll ToLeft:=48
    ActiveWindow.SmallScroll ToRight:=5 + 8 * 3
    .Range("AJ4").Select
Application.ScreenUpdating = True
End With
End Sub


Sub GoToDimensions()
With shDiscrete_Loads
Application.ScreenUpdating = False
    ActiveWindow.SmallScroll ToLeft:=48
    ActiveWindow.SmallScroll ToRight:=5 + 8 * 4
    .Range("AR4").Select
Application.ScreenUpdating = True
End With
End Sub


'======================================================================'
'============================ COPY BUTTONS ============================'
'======================================================================'

Sub Copy_BLC()
lastRow = shRISA_3D.Range("A5").End(xlDown).row
shRISA_3D.Range("B5", "F" & lastRow).Copy
End Sub

Sub Copy_LC()
lastRow = shRISA_3D.Range("M5").End(xlDown).row
shRISA_3D.Range("N5", "AA" & lastRow).Copy
End Sub

Sub Copy_BLC1_Point()
lastRow = shPointLoadTables.Range("A3").End(xlDown).row
shPointLoadTables.Range("A4", "D" & lastRow).Copy
End Sub

Sub Copy_BLC2_Point()
lastRow = shPointLoadTables.Range("F3").End(xlDown).row
shPointLoadTables.Range("F4", "I" & lastRow).Copy
End Sub

Sub Copy_BLC3_Point()
lastRow = shPointLoadTables.Range("K3").End(xlDown).row
shPointLoadTables.Range("K4", "N" & lastRow).Copy
End Sub

Sub Copy_BLC4_Point()
lastRow = shPointLoadTables.Range("P3").End(xlDown).row
shPointLoadTables.Range("P4", "S" & lastRow).Copy
End Sub

Sub Copy_BLC5_Point()
lastRow = shPointLoadTables.Range("U3").End(xlDown).row
shPointLoadTables.Range("U4", "X" & lastRow).Copy
End Sub

Sub Copy_BLC6_Point()
lastRow = shPointLoadTables.Range("Z3").End(xlDown).row
shPointLoadTables.Range("Z4", "AC" & lastRow).Copy
End Sub

Sub Copy_BLC7_Point()
lastRow = shPointLoadTables.Range("AE3").End(xlDown).row
shPointLoadTables.Range("AE4", "AH" & lastRow).Copy
End Sub

Sub Copy_BLC8_Point()
lastRow = shPointLoadTables.Range("AJ3").End(xlDown).row
shPointLoadTables.Range("AJ4", "AM" & lastRow).Copy
End Sub

Sub Copy_BLC9_Point()
lastRow = shPointLoadTables.Range("AO3").End(xlDown).row
shPointLoadTables.Range("AO4", "AR" & lastRow).Copy
End Sub

Sub Copy_BLC10_Point()
lastRow = shPointLoadTables.Range("AT3").End(xlDown).row
shPointLoadTables.Range("AT4", "AW" & lastRow).Copy
End Sub

Sub Copy_BLC11_Point()
lastRow = shPointLoadTables.Range("AY3").End(xlDown).row
shPointLoadTables.Range("AY4", "BB" & lastRow).Copy
End Sub

Sub Copy_BLC12_Point()
lastRow = shPointLoadTables.Range("BD3").End(xlDown).row
shPointLoadTables.Range("BD4", "BG" & lastRow).Copy
End Sub

Sub Copy_BLC13_Point()
lastRow = shPointLoadTables.Range("BI3").End(xlDown).row
shPointLoadTables.Range("BI4", "BL" & lastRow).Copy
End Sub

Sub Copy_BLC14_Point()
lastRow = shPointLoadTables.Range("BN3").End(xlDown).row
shPointLoadTables.Range("BN4", "BQ" & lastRow).Copy
End Sub

Sub Copy_BLC15_Point()
lastRow = shPointLoadTables.Range("BS3").End(xlDown).row
shPointLoadTables.Range("BS4", "BV" & lastRow).Copy
End Sub

Sub Copy_BLC16_Point()
lastRow = shPointLoadTables.Range("BX3").End(xlDown).row
shPointLoadTables.Range("BX4", "CA" & lastRow).Copy
End Sub

Sub Copy_BLC33_Point()
lastRow = shPointLoadTables.Range("CC3").End(xlDown).row
shPointLoadTables.Range("CC4", "CF" & lastRow).Copy
End Sub

Sub Copy_BLC34_Point()
lastRow = shPointLoadTables.Range("CH3").End(xlDown).row
shPointLoadTables.Range("CH4", "CK" & lastRow).Copy
End Sub

Sub Copy_BLC35_Point()
lastRow = shPointLoadTables.Range("CM3").End(xlDown).row
shPointLoadTables.Range("CM4", "CP" & lastRow).Copy
End Sub

Sub Copy_BLC36_Point()
lastRow = shPointLoadTables.Range("CR3").End(xlDown).row
shPointLoadTables.Range("CR4", "CU" & lastRow).Copy
End Sub

Sub Copy_BLC37_Point()
lastRow = shPointLoadTables.Range("CW3").End(xlDown).row
shPointLoadTables.Range("CW4", "CZ" & lastRow).Copy
End Sub

Sub Copy_BLC38_Point()
lastRow = shPointLoadTables.Range("DB3").End(xlDown).row
shPointLoadTables.Range("DB4", "DE" & lastRow).Copy
End Sub

Sub Copy_BLC39_Point()
lastRow = shPointLoadTables.Range("DG3").End(xlDown).row
shPointLoadTables.Range("DG4", "DJ" & lastRow).Copy
End Sub

Sub Copy_BLC40_Point()
lastRow = shPointLoadTables.Range("DL3").End(xlDown).row
shPointLoadTables.Range("DL4", "DO" & lastRow).Copy
End Sub

Sub Copy_BLC41_Point()
lastRow = shPointLoadTables.Range("DQ3").End(xlDown).row
shPointLoadTables.Range("DQ4", "DT" & lastRow).Copy
End Sub

Sub Copy_BLC42_Point()
lastRow = shPointLoadTables.Range("DV3").End(xlDown).row
shPointLoadTables.Range("DV4", "DY" & lastRow).Copy
End Sub

Sub Copy_BLC43_Point()
lastRow = shPointLoadTables.Range("EA3").End(xlDown).row
shPointLoadTables.Range("EA4", "ED" & lastRow).Copy
End Sub

Sub Copy_BLC44_Point()
lastRow = shPointLoadTables.Range("EF3").End(xlDown).row
shPointLoadTables.Range("EF4", "EI" & lastRow).Copy
End Sub

Sub Copy_BLC45_Point()
lastRow = shPointLoadTables.Range("EK3").End(xlDown).row
shPointLoadTables.Range("EK4", "EN" & lastRow).Copy
End Sub

Sub Copy_BLC46_Point()
lastRow = shPointLoadTables.Range("EP3").End(xlDown).row
shPointLoadTables.Range("EP4", "ES" & lastRow).Copy
End Sub

Sub Copy_BLC47_Point()
lastRow = shPointLoadTables.Range("EU3").End(xlDown).row
shPointLoadTables.Range("EU4", "EX" & lastRow).Copy
End Sub

Sub Copy_BLC48_Point()
lastRow = shPointLoadTables.Range("EZ3").End(xlDown).row
shPointLoadTables.Range("EZ4", "FC" & lastRow).Copy
End Sub

Sub Copy_BLC65_Point()
lastRow = shPointLoadTables.Range("FE3").End(xlDown).row
shPointLoadTables.Range("FE4", "FH" & lastRow).Copy
End Sub

Sub Copy_BLC66_Point()
lastRow = shPointLoadTables.Range("FJ3").End(xlDown).row
shPointLoadTables.Range("FJ4", "FM" & lastRow).Copy
End Sub

Sub Copy_BLC67_Point()
lastRow = shPointLoadTables.Range("FO3").End(xlDown).row
shPointLoadTables.Range("FO4", "FR" & lastRow).Copy
End Sub

Sub Copy_BLC68_Point()
lastRow = shPointLoadTables.Range("FT3").End(xlDown).row
shPointLoadTables.Range("FT4", "FW" & lastRow).Copy
End Sub

Sub Copy_BLC69_Point()
lastRow = shPointLoadTables.Range("FY3").End(xlDown).row
shPointLoadTables.Range("FY4", "GB" & lastRow).Copy
End Sub

Sub Copy_BLC86_Point()
lastRow = shPointLoadTables.Range("GD3").End(xlDown).row
shPointLoadTables.Range("GD4", "GG" & lastRow).Copy
End Sub

Sub Copy_BLC87_Point()
lastRow = shPointLoadTables.Range("GI3").End(xlDown).row
shPointLoadTables.Range("GI4", "GL" & lastRow).Copy
End Sub

Sub Copy_BLC88_Point()
lastRow = shPointLoadTables.Range("GN3").End(xlDown).row
shPointLoadTables.Range("GN4", "GQ" & lastRow).Copy
End Sub

Sub Copy_BLC89_Point()
lastRow = shPointLoadTables.Range("GS3").End(xlDown).row
shPointLoadTables.Range("GS4", "GV" & lastRow).Copy
End Sub

Sub Copy_BLC90_Point()
lastRow = shPointLoadTables.Range("GX3").End(xlDown).row
shPointLoadTables.Range("GX4", "HA" & lastRow).Copy
End Sub

Sub Copy_BLC91_Point()
lastRow = shPointLoadTables.Range("HC3").End(xlDown).row
shPointLoadTables.Range("HC4", "HF" & lastRow).Copy
End Sub

Sub Copy_BLC92_Point()
lastRow = shPointLoadTables.Range("HH3").End(xlDown).row
shPointLoadTables.Range("HH4", "HK" & lastRow).Copy
End Sub

Sub Copy_BLC93_Point()
lastRow = shPointLoadTables.Range("HM3").End(xlDown).row
shPointLoadTables.Range("HM4", "HP" & lastRow).Copy
End Sub


Sub Copy_BLC17_Distributed()
lastRow = shDistributedLoadTables.Range("A3").End(xlDown).row
shDistributedLoadTables.Range("A4", "F" & lastRow).Copy
End Sub



Sub Copy_BLC18_Distributed()
lastRow = shDistributedLoadTables.Range("H3").End(xlDown).row
shDistributedLoadTables.Range("H4", "M" & lastRow).Copy
End Sub



Sub Copy_BLC19_Distributed()
lastRow = shDistributedLoadTables.Range("O3").End(xlDown).row
shDistributedLoadTables.Range("O4", "T" & lastRow).Copy
End Sub



Sub Copy_BLC20_Distributed()
lastRow = shDistributedLoadTables.Range("V3").End(xlDown).row
shDistributedLoadTables.Range("V4", "AA" & lastRow).Copy
End Sub



Sub Copy_BLC21_Distributed()
lastRow = shDistributedLoadTables.Range("AC3").End(xlDown).row
shDistributedLoadTables.Range("AC4", "AH" & lastRow).Copy
End Sub



Sub Copy_BLC22_Distributed()
lastRow = shDistributedLoadTables.Range("AJ3").End(xlDown).row
shDistributedLoadTables.Range("AJ4", "AO" & lastRow).Copy
End Sub



Sub Copy_BLC23_Distributed()
lastRow = shDistributedLoadTables.Range("AQ3").End(xlDown).row
shDistributedLoadTables.Range("AQ4", "AV" & lastRow).Copy
End Sub



Sub Copy_BLC24_Distributed()
lastRow = shDistributedLoadTables.Range("AX3").End(xlDown).row
shDistributedLoadTables.Range("AX4", "BC" & lastRow).Copy
End Sub



Sub Copy_BLC25_Distributed()
lastRow = shDistributedLoadTables.Range("BE3").End(xlDown).row
shDistributedLoadTables.Range("BE4", "BJ" & lastRow).Copy
End Sub



Sub Copy_BLC26_Distributed()
lastRow = shDistributedLoadTables.Range("BL3").End(xlDown).row
shDistributedLoadTables.Range("BL4", "BQ" & lastRow).Copy
End Sub



Sub Copy_BLC27_Distributed()
lastRow = shDistributedLoadTables.Range("BS3").End(xlDown).row
shDistributedLoadTables.Range("BS4", "BX" & lastRow).Copy
End Sub



Sub Copy_BLC28_Distributed()
lastRow = shDistributedLoadTables.Range("BZ3").End(xlDown).row
shDistributedLoadTables.Range("BZ4", "CE" & lastRow).Copy
End Sub



Sub Copy_BLC29_Distributed()
lastRow = shDistributedLoadTables.Range("CG3").End(xlDown).row
shDistributedLoadTables.Range("CG4", "CL" & lastRow).Copy
End Sub



Sub Copy_BLC30_Distributed()
lastRow = shDistributedLoadTables.Range("CN3").End(xlDown).row
shDistributedLoadTables.Range("CN4", "CS" & lastRow).Copy
End Sub



Sub Copy_BLC31_Distributed()
lastRow = shDistributedLoadTables.Range("CU3").End(xlDown).row
shDistributedLoadTables.Range("CU4", "CZ" & lastRow).Copy
End Sub



Sub Copy_BLC32_Distributed()
lastRow = shDistributedLoadTables.Range("DB3").End(xlDown).row
shDistributedLoadTables.Range("DB4", "DG" & lastRow).Copy
End Sub


Sub Copy_BLC49_Distributed()
lastRow = shDistributedLoadTables.Range("DI3").End(xlDown).row
shDistributedLoadTables.Range("DI4", "DN" & lastRow).Copy
End Sub



Sub Copy_BLC50_Distributed()
lastRow = shDistributedLoadTables.Range("DP3").End(xlDown).row
shDistributedLoadTables.Range("DP4", "DU" & lastRow).Copy
End Sub



Sub Copy_BLC51_Distributed()
lastRow = shDistributedLoadTables.Range("DW3").End(xlDown).row
shDistributedLoadTables.Range("DW4", "EB" & lastRow).Copy
End Sub



Sub Copy_BLC52_Distributed()
lastRow = shDistributedLoadTables.Range("ED3").End(xlDown).row
shDistributedLoadTables.Range("ED4", "EI" & lastRow).Copy
End Sub



Sub Copy_BLC53_Distributed()
lastRow = shDistributedLoadTables.Range("EK3").End(xlDown).row
shDistributedLoadTables.Range("EK4", "EP" & lastRow).Copy
End Sub



Sub Copy_BLC54_Distributed()
lastRow = shDistributedLoadTables.Range("ER3").End(xlDown).row
shDistributedLoadTables.Range("ER4", "EW" & lastRow).Copy
End Sub



Sub Copy_BLC55_Distributed()
lastRow = shDistributedLoadTables.Range("EY3").End(xlDown).row
shDistributedLoadTables.Range("EY4", "FD" & lastRow).Copy
End Sub



Sub Copy_BLC56_Distributed()
lastRow = shDistributedLoadTables.Range("FF3").End(xlDown).row
shDistributedLoadTables.Range("FF4", "FK" & lastRow).Copy
End Sub



Sub Copy_BLC57_Distributed()
lastRow = shDistributedLoadTables.Range("FM3").End(xlDown).row
shDistributedLoadTables.Range("FM4", "FR" & lastRow).Copy
End Sub



Sub Copy_BLC58_Distributed()
lastRow = shDistributedLoadTables.Range("FT3").End(xlDown).row
shDistributedLoadTables.Range("FT4", "FY" & lastRow).Copy
End Sub



Sub Copy_BLC59_Distributed()
lastRow = shDistributedLoadTables.Range("GA3").End(xlDown).row
shDistributedLoadTables.Range("GA4", "GF" & lastRow).Copy
End Sub



Sub Copy_BLC60_Distributed()
lastRow = shDistributedLoadTables.Range("GH3").End(xlDown).row
shDistributedLoadTables.Range("GH4", "GM" & lastRow).Copy
End Sub



Sub Copy_BLC61_Distributed()
lastRow = shDistributedLoadTables.Range("GO3").End(xlDown).row
shDistributedLoadTables.Range("GO4", "GT" & lastRow).Copy
End Sub



Sub Copy_BLC62_Distributed()
lastRow = shDistributedLoadTables.Range("GV3").End(xlDown).row
shDistributedLoadTables.Range("GV4", "HA" & lastRow).Copy
End Sub



Sub Copy_BLC63_Distributed()
lastRow = shDistributedLoadTables.Range("HC3").End(xlDown).row
shDistributedLoadTables.Range("HC4", "HH" & lastRow).Copy
End Sub



Sub Copy_BLC64_Distributed()
lastRow = shDistributedLoadTables.Range("HJ3").End(xlDown).row
shDistributedLoadTables.Range("HJ4", "HO" & lastRow).Copy
End Sub



Sub Copy_BLC66_Distributed()
lastRow = shDistributedLoadTables.Range("HQ3").End(xlDown).row
shDistributedLoadTables.Range("HQ4", "HV" & lastRow).Copy
End Sub





Sub Copy_BLC65_Area()
lastRow = shAreaLoadTables.Range("A3").End(xlDown).row
shAreaLoadTables.Range("A4", "G" & lastRow).Copy
End Sub





Sub Copy_BLC66_Area()
lastRow = shAreaLoadTables.Range("I3").End(xlDown).row
shAreaLoadTables.Range("I4", "O" & lastRow).Copy
End Sub



