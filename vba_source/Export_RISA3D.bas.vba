Attribute VB_Name = "Export_RISA3D"
Option Explicit

Sub All_Loads()
    Call Distributed_Loads
    Call LoadsPoint
    'Call Maintenance_Loads
    Call Area_Loads
    Call DeleteProjectData
End Sub

Sub Generate()
'ThisWorkbook.Save
Dim Answer As String

'frm_Company.Show

If Sheets("Distributed Load Tables").Range("A4") <> "" Or Sheets("Point Load Tables").Range("A4") <> "" Then
    Answer = MsgBox("Previous loads exist. Re-Calculate loading when generating new RISA-3D model?", vbYesNoCancel, "RISA-3D Model and Appurtenance Loading")
    If Answer = vbYes Then
        Call All_Loads
    Else
    End If
Else
    Answer = MsgBox("No loads have been created. Calculate loading when generating new RISA-3D model?", vbYesNoCancel, "RISA-3D Model and Appurtenance Loading")
    If Answer = vbYes Then
        Call All_Loads
    Else
    End If
End If


With shCode
If .Range("ProjectNumber") = "" Or .Range("Client") = "" Or .Range("Carrier") = "" Or _
.Range("Carrier_SiteName") = "" Or .Range("Carrier_SiteNumber") = "" Or .Range("Carrier_Project") = "" Then
    Answer = MsgBox("The Project Information has not properly been input. Please verify inputs and update as needed.", vbCritical + vbOKOnly, "Missing Project Information!")
    shCode.Activate
    .Range("C2:D7").Select
    Exit Sub
End If
End With
        
        
With WorksheetFunction
    If .CountA(shMaintenance_Loads.Range("B7:B10")) <> .CountA(shMaintenance_Loads.Range("E7:E10")) Or _
       .CountA(shMaintenance_Loads.Range("G7:G10")) <> .CountA(shMaintenance_Loads.Range("J7:J10")) Then
        Answer = MsgBox("The Maintenance Load Table has not properly been input. Please verify inputs and update as needed.", vbCritical + vbOKOnly, "Missing Maintenance Load Information!")
        Exit Sub
    End If
End With
        

If Range("UserInputCheck") = True Or Range("FalseCheck") = True Then
    Answer = MsgBox("The RISA-3D Section Sets table on the Geometry tab is missing information. Please update any values listed as User Input.", vbCritical + vbOKOnly, "Missing Section Set Information!")
    Exit Sub
End If


If WorksheetFunction.CountA(shMaintenance_Loads.Range("L7:L10")) = 0 Then
    Call Generate_NoPlatforms
Else
    Call Generate_WithPlatforms
End If

'Call ModelEPA

End Sub

Sub Generate_NoPlatforms()
    '=============================
    ' RISA export / LOADS variables
    '=============================
 
    Dim PointLoads_Header As Range, DistributedLoads_Header As Range               'Point Load Header, Distributed Load Header, Area Load Header'
    Dim BLC_Header As Range, LC_Header As Range                                'BLC Header, LC Header'
    Dim BLCPadding As String                                                   'To pad BLC description'
    Dim LCLongPadding As String, LCShortPadding As String                      'To pad strings required in Load Combinations'
    Dim Members As Long, EndMembers As Long                                    'Members Header, End of Members Section'
    Dim Nodes As Long, EndNodes As Long                                        'Nodes Header, End of Nodes Section'
    Dim Tables_PointLoads As Worksheet, Tables_DistributedLoads As Worksheet, Tables_AreaLoads As Worksheet          '"Table" = Loads tables tab, "Txt" = Txtorary RISA input tab'
    Dim Temp As Worksheet
    Dim PointLoads_List() As Variant, DistributedLoads_List() As Variant           'Array for Point Loads, Array for Distributed Loads, Array for Area Loads'
    Dim LC_List() As Variant                                                   'Array for Load Combinations to optimise padding & decimal updates'
    Dim LC_Decimals() As Variant, LC_Padding() As Variant                      'to hold col nums of Load Combinations which require 6 decimals, which required padding'
    Dim LC_Range As Range                                                      'create variable for dynamic range referenced multiple times'
    Dim r As Integer, i As Long, c As Long, t As Long, a As Long, cnt As Long  'r=detemrine size of point load array'
    Dim ManCount As Integer
    Dim n1 As String, n2 As String, n3 As String, n4 As String
    Dim x1 As Range, x2 As Range, x3 As Range, x4 As Range
    Dim h As Byte
    Dim m As String, n As String
    Dim m1 As String, m2 As String, m3 As String, m4 As String
    Dim X As Range
    Dim firstRow As Long, lastRow As Long, lastCol As Long
    Dim Width As Integer, Depth As Long
    Dim headers() As String                                                    'to hold names of 4 headers - for sorting approach'
    Dim anchors() As String                                                    'to hold anchor headers for any headers missing'
    Dim inserts(0 To 3) As Long                                             'to hold number of new items in each LOADS section'
    Dim findHdr As Range, insertRow As Long
    Const StartRow = 3
    Dim Start As Double, Answer As String
    
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevSU As Boolean
    Dim prevSB As Variant

    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevSU = Application.ScreenUpdating
    prevSB = Application.StatusBar

    On Error GoTo CleanFail

    If MsgBox("Export loads into RISA-3D model? This may take several minutes.", _
              vbOKCancel, "Export Loads") = vbCancel Then
        GoTo CleanExit
    End If

    'TURN OFF (for OK path)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Busy..."
    
    Start = Timer
 
    Set Temp = Sheets(Sheets.Count)
    With Temp
        If Application.WorksheetFunction.CountA(.Cells) = 0 Then
            lastRow = 1: lastCol = 1
        Else
            lastRow = .Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious).row
            lastCol = .Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
                                  SearchDirection:=xlPrevious).Column
        End If
    End With
    
    Depth = lastRow
    Width = lastCol
    
    Set Tables_PointLoads = shPointLoadTables
    Set Tables_DistributedLoads = shDistributedLoadTables
    
    headers = Strings.Split("[BASIC_LOAD_CASES],[POINT_LOADS],[DIRECT_DISTRIBUTED_LOADS],[LOAD_COMBINATIONS]", ",")
    
    With WorksheetFunction
        'new BLC entries'
        inserts(0) = shRISA_3D.Range("A5").End(xlDown).row - 4
        
        'to determine size of Point Loads array'
        inserts(1) = .CountA(Range(Tables_PointLoads.Range("A4").End(xlDown), Tables_PointLoads.Cells(4, 181 + 3 + (Range("nLm").Value + Range("nLv").Value) * 5))) / 4
        
        'to determine size of Distributed Loads array'
        inserts(2) = .CountA(Range(Tables_DistributedLoads.Range("A4").End(xlDown), Tables_DistributedLoads.Cells(4, 230))) / 6
        
        'new LOAD COMBINATION entries'
        inserts(3) = Range("nLC").Value
    End With
    
    anchors = Strings.Split("[END_MEMBERS],[END_BASIC_LOAD_CASES],[END_POINT_LOADS],[END_DIRECT_DISTRIBUTED_LOADS],[END_SPECTRA_SCALING_FACTOR],[END_NODES]", ",")
    
    
    With Temp
       'Populate sort key column'
        .Cells(1, Width + 1).Value = 1
        Range(.Cells(1, Width + 1), .Cells(Depth, Width + 1)).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, Step:=1, Trend:=False
        
         'Populate LOADS headers if any are missing'
        insertRow = Depth + 1
        For i = 0 To UBound(headers)
            If .Range("A:A").Find(headers(i)) Is Nothing Then
                Set findHdr = .Range("A:A").Find(anchors(i)).offset(2, 0)     'offset of 2 so that header will be inserted to leave 2 blank rows after anchor header'
                .Cells(insertRow, 1).Value = headers(i)                                                         'header'
                .Cells(insertRow, Width + 1).Value = findHdr.row + 0.2
                .Cells(insertRow + 1, 1).Value = "[END_" & Strings.Right(headers(i), Len(headers(i)) - 1)       'footer'
                .Cells(insertRow + 1, Width + 1).Value = findHdr.row + 0.4
                .Cells(insertRow + 2, Width + 1).Value = findHdr.row + 0.6       'first blank row before next header (after sort)'
                .Cells(insertRow + 3, Width + 1).Value = findHdr.row + 0.8       'second blank row before next header (after sort)'
                insertRow = insertRow + 4
                'Sort data to put missing LOADS sections in correct positions'
                Range(.Cells(1, 1), .Cells(insertRow - 1, Width + 1)).Sort key1:=.Cells(1, Width + 1), order1:=xlAscending, Header:=xlNo
                'Recreate sort key to eliminate decimals'
                Range(.Cells(1, Width + 1), .Cells(insertRow - 1, Width + 1)).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, Step:=1, Trend:=False
            End If
        Next i
        
        'iterate through 3 LOAD TYPE sections'
        For i = 0 To UBound(headers)
            Set findHdr = .Range("A:A").Find(headers(i))
                'Clear any existing entries'
                If InStr(findHdr.offset(1, 0), Strings.Right(headers(i), Len(headers(i)) - 1)) < 1 Then
                    Range(findHdr.offset(1, 0), findHdr.offset(1, 0).End(xlDown).offset(-1, 0)).EntireRow.ClearContents
                End If
                'Flag inserts required'
                insertRow = .Cells(Rows.Count, Width + 1).End(xlUp).row + 1
                Range(.Cells(insertRow, Width + 1), .Cells(insertRow + inserts(i) - 1, Width + 1)).Value = findHdr.row + 0.5
                'Write number of inserts as content'
                findHdr.offset(0, 1).Value = "<" & inserts(i) & ">"
        Next i
        'Sort data'
        Range(.Cells(1, 1), .Cells(insertRow + inserts(3) - 1, Width + 1)).Sort key1:=.Cells(1, Width + 1), order1:=xlAscending, Header:=xlNo
        'Clear sort key'
        .Columns(Width + 1).ClearContents
    End With
    
    BLCPadding = WorksheetFunction.Rept(" ", 32)
    LCLongPadding = WorksheetFunction.Rept(" ", 79)
    LCShortPadding = WorksheetFunction.Rept(" ", 32)
    LC_Decimals = Array(7, 8, 9, 26, 28, 30, 32, 34, 36, 38, 40, 42)
    LC_Padding = Array(25, 27, 29, 31, 33, 35, 37, 39, 41, 43)
 
    Members = Temp.Range("A:A").Find("[.MEMBERS_MAIN_DATA]").row
    EndMembers = Temp.Range("A:A").Find("[.END_MEMBERS_MAIN_DATA]").row
    
    Nodes = Temp.Range("A:A").Find("[NODES]").row
    EndNodes = Temp.Range("A:A").Find("[END_NODES]").row
 
    '>>>>>>POINT LOADS<<<<<<'
    Set PointLoads_Header = Temp.Range("A:A").Find("[POINT_LOADS]") 'Finds "POINT LOADS" section of Text file'
        ReDim PointLoads_List(1 To inserts(1), 1 To 10)
        ManCount = 181 + (Range("nLm").Value + Range("nLv").Value) * 5
        
        a = 1
        'loops to capture each loading configuration. Loop move horizontally by setting "t" as a "column" variable'
        For t = 1 To ManCount Step 5
            i = 1
            'populate array until there is no member label'
            Do Until Tables_PointLoads.Cells(i + StartRow, t) = ""
               Set X = Temp.Range("A" & Members, "A" & EndMembers).Find(What:=Tables_PointLoads.Cells(StartRow + i, t).Value, LookIn:=xlValues, LookAt:=xlPart)
               'match members label to table index'
                m = Application.WorksheetFunction.Match(X, Temp.Range("A" & Members, "A" & EndMembers), 0) - 1
                PointLoads_List(a, 1) = m
                    'converts load direction to RISA file format'
                    Select Case Tables_PointLoads.Cells(i + StartRow, t + 1)
                        Case "X": PointLoads_List(a, 2) = 88
                        Case "Y": PointLoads_List(a, 2) = 89
                        Case "Z": PointLoads_List(a, 2) = 90
                        Case "Mx": PointLoads_List(a, 2) = 125
                        Case "My": PointLoads_List(a, 2) = 126
                        Case "Mz": PointLoads_List(a, 2) = 127
                    End Select
                'Load magnitude'
                PointLoads_List(a, 3) = Strings.Format(Tables_PointLoads.Cells(i + StartRow, t + 2).Value, "0.000!")
                'Load location'
                PointLoads_List(a, 4) = Strings.Format(Tables_PointLoads.Cells(i + StartRow, t + 3) * -100, "0.0!")
                PointLoads_List(a, 5) = 0
                PointLoads_List(a, 6) = -1
                PointLoads_List(a, 7) = "0.000000!"
                PointLoads_List(a, 8) = 2
                PointLoads_List(a, 9) = "0.000000!"
                PointLoads_List(a, 10) = "0.000000;"
                i = i + 1
                a = a + 1
            Loop
        Next t
        'Write fully populated array to POINTS_LOADS section of sheet'
        Range(PointLoads_Header.offset(1, 0), PointLoads_Header.offset(inserts(1), 9)).Value = PointLoads_List
    
    
    '>>>>>>DISTRIBUTED LOADS<<<<<<'
    'Finds "DISTRIBUTED LOADS" section of Text file'
    Set DistributedLoads_Header = Temp.Range("A:A").Find("[DIRECT_DISTRIBUTED_LOADS]")
        ReDim DistributedLoads_List(1 To inserts(2), 1 To 14)

        a = 1
        For t = 1 To 225 Step 7  'loops to capture each loading configuration. Loop move horizontally by setting "t" as a "column" variable'
            i = 1
            'populate array until there is no member label'
            Do Until Tables_DistributedLoads.Cells(i + StartRow, t) = ""
               Set X = Temp.Range("A" & Members, "A" & EndMembers).Find(What:=Tables_DistributedLoads.Cells(StartRow + i, t).Value, LookIn:=xlValues, LookAt:=xlPart)
               'match members label to table index'
                m = Application.WorksheetFunction.Match(X, Temp.Range("A" & Members, "A" & EndMembers), 0) - 1
                DistributedLoads_List(a, 1) = m
                'converts load direction to RISA file format'
                Select Case Tables_DistributedLoads.Cells(i + StartRow, t + 1)
                    Case "X": DistributedLoads_List(a, 2) = 88
                    Case "Y": DistributedLoads_List(a, 2) = 89
                    Case "Z": DistributedLoads_List(a, 2) = 90
                End Select
                'Load magnitude'
                DistributedLoads_List(a, 3) = Strings.Format(Tables_DistributedLoads.Cells(i + StartRow, t + 2).Value, "0.000!")
                DistributedLoads_List(a, 4) = Strings.Format(Tables_DistributedLoads.Cells(i + StartRow, t + 3).Value, "0.000!")
                'Load location'
                DistributedLoads_List(a, 5) = Strings.Format(Tables_DistributedLoads.Cells(i + StartRow, t + 4).Value * -100, "0.0!")
                DistributedLoads_List(a, 6) = Strings.Format(Tables_DistributedLoads.Cells(i + StartRow, t + 5).Value * -100, "0.0!")
                DistributedLoads_List(a, 7) = 0
                DistributedLoads_List(a, 8) = -1
                DistributedLoads_List(a, 9) = "0.000000!"
                DistributedLoads_List(a, 10) = "0.000000!"
                DistributedLoads_List(a, 11) = 2
                DistributedLoads_List(a, 12) = "0.000000!"
                DistributedLoads_List(a, 13) = "0.000000!"
                DistributedLoads_List(a, 14) = "0.000000;"
                i = i + 1
                a = a + 1
            Loop
        Next t
        'Write fully populated array to DIRECT_DISTRIBUTED_LOADS section of sheet'
        Range(DistributedLoads_Header.offset(1, 0), DistributedLoads_Header.offset(inserts(2), 13)).Value = DistributedLoads_List

    
        '>>>>>>POPULATES BASIC LOAD CASES TABLE<<<<<<'
    Set BLC_Header = Temp.Range("A:A").Find("[BASIC_LOAD_CASES]") 'Finds "BLC" section of Text file'
    For i = 1 To inserts(0)
        BLC_Header.offset(i) = i
        BLC_Header.offset(i, 1) = "^" & Strings.Left(shRISA_3D.Range("B" & 4 + i).Value & BLCPadding, 42) & "^" 'Populates whole table with "0". Populated respective values below'
        Range(BLC_Header.offset(i, 2), BLC_Header.offset(i, 8)).Value = 0
        Range(BLC_Header.offset(i, 9), BLC_Header.offset(i, 11)).Value = "0.000000!"
        BLC_Header.offset(i, 12) = "-1;"
    Next i
 

   With BLC_Header
         '>>>>>>WIND AND ICE WIND DISTRIBUTED LOADS<<<<<<'
        .offset(17, 5) = shDistributedLoadTables.Range("A3").End(xlDown).row - StartRow
        .offset(18, 5) = shDistributedLoadTables.Range("H3").End(xlDown).row - StartRow
        .offset(19, 5) = shDistributedLoadTables.Range("O3").End(xlDown).row - StartRow
        .offset(20, 5) = shDistributedLoadTables.Range("V3").End(xlDown).row - StartRow
        .offset(21, 5) = shDistributedLoadTables.Range("AC3").End(xlDown).row - StartRow
        .offset(22, 5) = shDistributedLoadTables.Range("AJ3").End(xlDown).row - StartRow
        .offset(23, 5) = shDistributedLoadTables.Range("AQ3").End(xlDown).row - StartRow
        .offset(24, 5) = shDistributedLoadTables.Range("AX3").End(xlDown).row - StartRow
        .offset(25, 5) = shDistributedLoadTables.Range("BE3").End(xlDown).row - StartRow
        .offset(26, 5) = shDistributedLoadTables.Range("BL3").End(xlDown).row - StartRow
        .offset(27, 5) = shDistributedLoadTables.Range("BS3").End(xlDown).row - StartRow
        .offset(28, 5) = shDistributedLoadTables.Range("BZ3").End(xlDown).row - StartRow
        .offset(29, 5) = shDistributedLoadTables.Range("CG3").End(xlDown).row - StartRow
        .offset(30, 5) = shDistributedLoadTables.Range("CN3").End(xlDown).row - StartRow
        .offset(31, 5) = shDistributedLoadTables.Range("CU3").End(xlDown).row - StartRow
        .offset(32, 5) = shDistributedLoadTables.Range("DB3").End(xlDown).row - StartRow

        .offset(49, 5) = shDistributedLoadTables.Range("DI3").End(xlDown).row - StartRow
        .offset(50, 5) = shDistributedLoadTables.Range("DP3").End(xlDown).row - StartRow
        .offset(51, 5) = shDistributedLoadTables.Range("DW3").End(xlDown).row - StartRow
        .offset(52, 5) = shDistributedLoadTables.Range("ED3").End(xlDown).row - StartRow
        .offset(53, 5) = shDistributedLoadTables.Range("EK3").End(xlDown).row - StartRow
        .offset(54, 5) = shDistributedLoadTables.Range("ER3").End(xlDown).row - StartRow
        .offset(55, 5) = shDistributedLoadTables.Range("EY3").End(xlDown).row - StartRow
        .offset(56, 5) = shDistributedLoadTables.Range("FF3").End(xlDown).row - StartRow
        .offset(57, 5) = shDistributedLoadTables.Range("FM3").End(xlDown).row - StartRow
        .offset(58, 5) = shDistributedLoadTables.Range("FT3").End(xlDown).row - StartRow
        .offset(59, 5) = shDistributedLoadTables.Range("GA3").End(xlDown).row - StartRow
        .offset(60, 5) = shDistributedLoadTables.Range("GH3").End(xlDown).row - StartRow
        .offset(61, 5) = shDistributedLoadTables.Range("GO3").End(xlDown).row - StartRow
        .offset(62, 5) = shDistributedLoadTables.Range("GV3").End(xlDown).row - StartRow
        .offset(63, 5) = shDistributedLoadTables.Range("HC3").End(xlDown).row - StartRow
        .offset(64, 5) = shDistributedLoadTables.Range("HJ3").End(xlDown).row - StartRow
    
          '>>>>>>ICE DEAD DISTRIBUTED LOADS<<<<<<'
        .offset(66, 5) = shDistributedLoadTables.Range("HQ3").End(xlDown).row - StartRow
 
   End With


    With BLC_Header
         '>>>>>>DEAD AND ICE DEAD POINT LOADS<<<<<<'
        .offset(65, 3) = shPointLoadTables.Range("FE3").End(xlDown).row - StartRow
        .offset(66, 3) = shPointLoadTables.Range("FJ3").End(xlDown).row - StartRow
        
                
        '>>>>>>WIND AND ICE WIND POINT LOADS<<<<<<'
        .offset(1, 3) = shPointLoadTables.Range("A3").End(xlDown).row - StartRow
        .offset(2, 3) = shPointLoadTables.Range("F3").End(xlDown).row - StartRow
        .offset(3, 3) = shPointLoadTables.Range("K3").End(xlDown).row - StartRow
        .offset(4, 3) = shPointLoadTables.Range("P3").End(xlDown).row - StartRow
        .offset(5, 3) = shPointLoadTables.Range("U3").End(xlDown).row - StartRow
        .offset(6, 3) = shPointLoadTables.Range("Z3").End(xlDown).row - StartRow
        .offset(7, 3) = shPointLoadTables.Range("AE3").End(xlDown).row - StartRow
        .offset(8, 3) = shPointLoadTables.Range("AJ3").End(xlDown).row - StartRow
        .offset(9, 3) = shPointLoadTables.Range("AO3").End(xlDown).row - StartRow
        .offset(10, 3) = shPointLoadTables.Range("AT3").End(xlDown).row - StartRow
        .offset(11, 3) = shPointLoadTables.Range("AY3").End(xlDown).row - StartRow
        .offset(12, 3) = shPointLoadTables.Range("BD3").End(xlDown).row - StartRow
        .offset(13, 3) = shPointLoadTables.Range("BI3").End(xlDown).row - StartRow
        .offset(14, 3) = shPointLoadTables.Range("BN3").End(xlDown).row - StartRow
        .offset(15, 3) = shPointLoadTables.Range("BS3").End(xlDown).row - StartRow
        .offset(16, 3) = shPointLoadTables.Range("BX3").End(xlDown).row - StartRow
        
        .offset(33, 3) = shPointLoadTables.Range("CC3").End(xlDown).row - StartRow
        .offset(34, 3) = shPointLoadTables.Range("CH3").End(xlDown).row - StartRow
        .offset(35, 3) = shPointLoadTables.Range("CM3").End(xlDown).row - StartRow
        .offset(36, 3) = shPointLoadTables.Range("CR3").End(xlDown).row - StartRow
        .offset(37, 3) = shPointLoadTables.Range("CW3").End(xlDown).row - StartRow
        .offset(38, 3) = shPointLoadTables.Range("DB3").End(xlDown).row - StartRow
        .offset(39, 3) = shPointLoadTables.Range("DG3").End(xlDown).row - StartRow
        .offset(40, 3) = shPointLoadTables.Range("DL3").End(xlDown).row - StartRow
        .offset(41, 3) = shPointLoadTables.Range("DQ3").End(xlDown).row - StartRow
        .offset(42, 3) = shPointLoadTables.Range("DV3").End(xlDown).row - StartRow
        .offset(43, 3) = shPointLoadTables.Range("EA3").End(xlDown).row - StartRow
        .offset(44, 3) = shPointLoadTables.Range("EF3").End(xlDown).row - StartRow
        .offset(45, 3) = shPointLoadTables.Range("EK3").End(xlDown).row - StartRow
        .offset(46, 3) = shPointLoadTables.Range("EP3").End(xlDown).row - StartRow
        .offset(47, 3) = shPointLoadTables.Range("EU3").End(xlDown).row - StartRow
        .offset(48, 3) = shPointLoadTables.Range("EZ3").End(xlDown).row - StartRow


        '>>>>>>SEISMIC LOADS<<<<<<'
'        If Sheets("Loads and Equipment").Range("Code") = "H" Then
        .offset(67, 3) = shRISA_3D.Range("H71")
        .offset(68, 3) = shRISA_3D.Range("H72")
        .offset(69, 3) = shRISA_3D.Range("H73")
'        End If
        
        '>>>>>>MAINTENANCE LOADS<<<<<<'
        .offset(86, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("GD4:GD19"))
        .offset(87, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("GI4:GI19"))
        .offset(88, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("GN4:GN19"))
        .offset(89, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("GS4:GS19"))
        .offset(90, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("GX4:GX19"))
        .offset(91, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("HC4:HC19"))
        .offset(92, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("HH4:HH19"))
        .offset(93, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("HM4:HM19"))

        If shCode.Range("QtySectors") > 3 Then
        .offset(45, 3) = shPointLoadTables.Range("DL3").End(xlDown).row - StartRow
        .offset(49, 3) = shPointLoadTables.Range("DQ3").End(xlDown).row - StartRow
        End If
        
        .offset(65, 10) = "-1.000000!"
        .offset(65, 8) = "1"
        .offset(66, 8) = "5"
    
        .offset(70, 9) = shRISA_3D.Range("D74")
        .offset(71, 9) = shRISA_3D.Range("D75")
        .offset(72, 9) = shRISA_3D.Range("D76")
        .offset(73, 9) = shRISA_3D.Range("D77")
        .offset(74, 9) = shRISA_3D.Range("D78")
        .offset(75, 9) = shRISA_3D.Range("D79")
        .offset(76, 9) = shRISA_3D.Range("D80")
        .offset(77, 9) = shRISA_3D.Range("D81")
        .offset(78, 9) = shRISA_3D.Range("D82")
        .offset(79, 9) = shRISA_3D.Range("D83")
        .offset(80, 9) = shRISA_3D.Range("D84")
        .offset(81, 9) = shRISA_3D.Range("D85")
        .offset(82, 9) = shRISA_3D.Range("D86")
        .offset(83, 9) = shRISA_3D.Range("D87")
        .offset(84, 9) = shRISA_3D.Range("D88")
        .offset(85, 9) = shRISA_3D.Range("D89")
                
        .offset(70, 10) = shRISA_3D.Range("E74")
        .offset(71, 10) = shRISA_3D.Range("E75")
        .offset(72, 10) = shRISA_3D.Range("E76")
        .offset(73, 10) = shRISA_3D.Range("E77")
        .offset(74, 10) = shRISA_3D.Range("E78")
        .offset(75, 10) = shRISA_3D.Range("E79")
        .offset(76, 10) = shRISA_3D.Range("E80")
        .offset(77, 10) = shRISA_3D.Range("E81")
        .offset(78, 10) = shRISA_3D.Range("E82")
        .offset(79, 10) = shRISA_3D.Range("E83")
        .offset(80, 10) = shRISA_3D.Range("E84")
        .offset(81, 10) = shRISA_3D.Range("E85")
        .offset(82, 10) = shRISA_3D.Range("E86")
        .offset(83, 10) = shRISA_3D.Range("E87")
        .offset(84, 10) = shRISA_3D.Range("E88")
        .offset(85, 10) = shRISA_3D.Range("E89")
                
        .offset(70, 11) = shRISA_3D.Range("F74")
        .offset(71, 11) = shRISA_3D.Range("F75")
        .offset(72, 11) = shRISA_3D.Range("F76")
        .offset(73, 11) = shRISA_3D.Range("F77")
        .offset(74, 11) = shRISA_3D.Range("F78")
        .offset(75, 11) = shRISA_3D.Range("F79")
        .offset(76, 11) = shRISA_3D.Range("F80")
        .offset(77, 11) = shRISA_3D.Range("F81")
        .offset(78, 11) = shRISA_3D.Range("F82")
        .offset(79, 11) = shRISA_3D.Range("F83")
        .offset(80, 11) = shRISA_3D.Range("F84")
        .offset(81, 11) = shRISA_3D.Range("F85")
        .offset(82, 11) = shRISA_3D.Range("F86")
        .offset(83, 11) = shRISA_3D.Range("F87")
        .offset(84, 11) = shRISA_3D.Range("F88")
        .offset(85, 11) = shRISA_3D.Range("F89")

    End With

    'Range(BLC_Header.Offset(1, 5), BLC_Header.Offset(4, 5)).Value = Lastrow + 1 - FirstRow
 
    '>>>>>>POPULATES LOAD COMBINATIONS TABLE<<<<<<'
    Set LC_Header = Temp.Range("A:A").Find("[LOAD_COMBINATIONS]") 'Finds "LC" section of Text file'
    cnt = inserts(3)
    
    Range(LC_Header.offset(1), LC_Header.offset(cnt)).Value = shRISA_3D.Range("N5", "N" & cnt + 4).Value
    LC_Header.offset(1, 1).Resize(cnt) = 0
    For r = 5 To cnt + 5
        If shRISA_3D.Range("O" & r) = "True" Then
            LC_Header.offset(1 + r - 5, 3) = 1
        Else
            LC_Header.offset(1 + r - 5, 3) = 0
        End If
        If shRISA_3D.Range("P" & r) = "Y" Then
            LC_Header.offset(1 + r - 5, 2) = 1
        Else
            LC_Header.offset(1 + r - 5, 2) = 0
        End If
    Next r
    LC_Header.offset(1, 4).Resize(cnt, 5) = 0
    LC_Header.offset(1, 6).Resize(cnt, 3) = 1
    LC_Header.offset(1, 9).Resize(cnt, 2) = 0
    LC_Header.offset(1, 11).Resize(cnt, 7) = 1
    LC_Header.offset(1, 18).Resize(cnt) = -1
    LC_Header.offset(1, 19).Resize(cnt, 5) = 1
 
    For i = 1 To cnt
        Range(LC_Header.offset(i, 24), LC_Header.offset(i, 33)).Value = Range(shRISA_3D.Range("R" & 4 + i), shRISA_3D.Range("AA" & 4 + i)).Value
    Next i
    
    'Copy LOAD_COMBINATIONS to array for more efficient updating of decimals and padding'
    Set LC_Range = Range(LC_Header.offset(1, 0), LC_Header.offset(cnt, 43))
    LC_List = LC_Range.Value
    For i = 1 To cnt
        LC_List(i, 1) = "^" & Strings.Left(LC_List(i, 1) & LCLongPadding, 79) & "^"
        For a = 0 To UBound(LC_Padding)
            LC_List(i, LC_Padding(a)) = "^" & Strings.Left(LC_List(i, LC_Padding(a)) & LCShortPadding, 32) & "^"
        Next a
        For a = 0 To UBound(LC_Decimals)
            LC_List(i, LC_Decimals(a)) = Strings.Format(LC_List(i, LC_Decimals(a)), "0.000000!")
        Next a
    Next i
    
    'Write updated array back to sheet'
    LC_Range.Value = LC_List
    
    'Populate any residual blank cells with decimal 0'
    On Error Resume Next
    LC_Range.SpecialCells(xlCellTypeBlanks).Value = "0.000000!"
    On Error GoTo 0
  
    Dim f As Range, firstAddr As String, tgt As Range

    With Temp.Columns("A")
        Set f = .Find(What:="[WOOD_SCHEDULE", LookIn:=xlValues, LookAt:=xlPart, _
                      SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        If Not f Is Nothing Then
            firstAddr = f.Address
            Do
                ' Require ":" in adjacent Column B (allowing stray spaces)
                If Trim$(CStr(f.offset(0, 1).Value)) = ":" Then
                    Set tgt = f.offset(1, 4)   ' one row down, 4 cols right (col E relative to A)
                    If Not IsError(tgt.Value) And Len(CStr(tgt.Value)) > 0 Then
                        tgt.NumberFormat = "@"                    ' store as Text
                        tgt.Value = Format$(CDbl(tgt.Value), "0.000")  ' write as text "0.000"
                    End If
                End If
                Set f = .FindNext(f)
            Loop While Not f Is Nothing And f.Address <> firstAddr
        End If
    End With
    
    Call text_to_RISA
    
    Start = Timer - Start
    Debug.Print Start
    
Handle:
    Select Case Err.Number
        Case 1004
        AppActivate Application.Caption
        MsgBox "It appears that member label """ & Tables_DistributedLoads.Cells(StartRow + i, t) & """ does not exist in the model. Please fix Appurtenance/Dish Table and Re-calcuate Point Loads.", vbInformation, "Error!"
        Exit Sub
    End Select
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Ready"
    
CleanExit:
    'RESTORE
    Application.StatusBar = prevSB
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevSU
    Exit Sub

CleanFail:
    'optional: MsgBox Err.Description
    Resume CleanExit

End Sub


Sub Generate_WithPlatforms()
    '=============================
    ' RISA export / LOADS variables
    '=============================
    
    Dim PointLoads_Header As Range, DistributedLoads_Header As Range, AreaLoads_Header As Range                'Point Load Header, Distributed Load Header, Area Load Header'
    Dim BLC_Header As Range, LC_Header As Range                                'BLC Header, LC Header'
    Dim BLCPadding As String                                                   'To pad BLC description'
    Dim LCLongPadding As String, LCShortPadding As String                      'To pad strings required in Load Combinations'
    Dim Members As Long, EndMembers As Long                                    'Members Header, End of Members Section'
    Dim Nodes As Long, EndNodes As Long                                        'Nodes Header, End of Nodes Section'
    Dim Tables_PointLoads As Worksheet, Tables_DistributedLoads As Worksheet, Tables_AreaLoads As Worksheet          '"Table" = Loads tables tab, "Txt" = Txtorary RISA input tab'
    Dim Temp As Worksheet
    Dim PointLoads_List() As Variant, DistributedLoads_List() As Variant, AreaLoads_List As Variant            'Array for Point Loads, Array for Distributed Loads, Array for Area Loads'
    Dim LC_List() As Variant                                                   'Array for Load Combinations to optimise padding & decimal updates'
    Dim LC_Decimals() As Variant, LC_Padding() As Variant                      'to hold col nums of Load Combinations which require 6 decimals, which required padding'
    Dim LC_Range As Range                                                      'create variable for dynamic range referenced multiple times'
    Dim r As Integer, i As Long, c As Long, t As Long, a As Long, cnt As Long  'r=detemrine size of point load array'
    Dim ManCount As Integer
    Dim n1 As String, n2 As String, n3 As String, n4 As String
    Dim x1 As Range, x2 As Range, x3 As Range, x4 As Range
    Dim h As Byte
    Dim m As String, n As String
    Dim m1 As String, m2 As String, m3 As String, m4 As String
    Dim X As Range
    Dim firstRow As Integer, lastRow As Long, lastCol As Long
    Dim Width As Integer, Depth As Long
    Dim headers() As String                                                    'to hold names of 4 headers - for sorting approach'
    Dim anchors() As String                                                    'to hold anchor headers for any headers missing'
    Dim inserts(0 To 4) As Integer                                             'to hold number of new items in each LOADS section'
    Dim findHdr As Range, insertRow As Long
    Const StartRow = 3
    Dim Start As Double, Answer As String
    
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevSU As Boolean
    Dim prevSB As Variant

    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevSU = Application.ScreenUpdating
    prevSB = Application.StatusBar

    On Error GoTo CleanFail

    If MsgBox("Export loads into RISA-3D model? This may take several minutes.", _
              vbOKCancel, "Export Loads") = vbCancel Then
        GoTo CleanExit
    End If

    'TURN OFF (for OK path)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Busy..."
    
    Start = Timer
 
    Set Temp = Sheets(Sheets.Count)
    With Temp
        If Application.WorksheetFunction.CountA(.Cells) = 0 Then
            lastRow = 1: lastCol = 1
        Else
            lastRow = .Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious).row
            lastCol = .Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
                                  SearchDirection:=xlPrevious).Column
        End If
    End With
    
    Depth = lastRow
    Width = lastCol
 
    Set Tables_PointLoads = shPointLoadTables
    Set Tables_DistributedLoads = shDistributedLoadTables
    Set Tables_AreaLoads = shAreaLoadTables
    
    If WorksheetFunction.CountA(shMaintenance_Loads.Range("L7:L10")) = 0 Then
        headers = Strings.Split("[BASIC_LOAD_CASES],[POINT_LOADS],[DIRECT_DISTRIBUTED_LOADS],[LOAD_COMBINATIONS]", ",")
    Else
        headers = Strings.Split("[BASIC_LOAD_CASES],[POINT_LOADS],[DIRECT_DISTRIBUTED_LOADS],[AREA_LOADS],[LOAD_COMBINATIONS]", ",")
    End If
    
    With WorksheetFunction
        'new BLC entries'
        inserts(0) = shRISA_3D.Range("A5").End(xlDown).row - 4
        
        'to determine size of Point Loads array'
        inserts(1) = .CountA(Range(Tables_PointLoads.Range("A4").End(xlDown), Tables_PointLoads.Cells(4, 181 + 3 + (Range("nLm").Value + Range("nLv").Value) * 5))) / 4
        
        'to determine size of Distributed Loads array'
        inserts(2) = .CountA(Range(Tables_DistributedLoads.Range("A4").End(xlDown), Tables_DistributedLoads.Cells(4, 230))) / 6
        
        'to determine size of Area Loads array'
        If .CountA(shMaintenance_Loads.Range("L7:L10")) = 0 Then
        Else
            inserts(3) = .CountA(Range(Tables_AreaLoads.Range("A4").End(xlDown), Tables_AreaLoads.Cells(4, 15))) / 7
        End If
        
        'new LOAD COMBINATION entries'
        inserts(4) = shRISA_3D.Range("nLC").Value 'shRISA_3D.Range("M5").End(xlDown).Row - 4
    End With
    
    If WorksheetFunction.CountA(shMaintenance_Loads.Range("L7:L10")) = 0 Then
        anchors = Strings.Split("[END_MEMBERS],[END_BASIC_LOAD_CASES],[END_POINT_LOADS],[END_DIRECT_DISTRIBUTED_LOADS],[END_SPECTRA_SCALING_FACTOR],[END_NODES]", ",")
    Else
        anchors = Strings.Split("[END_MEMBERS],[END_BASIC_LOAD_CASES],[END_POINT_LOADS],[END_DIRECT_DISTRIBUTED_LOADS],[END_AREA_LOADS],[END_SPECTRA_SCALING_FACTOR],[END_NODES]", ",")
    End If
    
    
    With Temp
       'Populate sort key column'
        .Cells(1, Width + 1).Value = 1
        Range(.Cells(1, Width + 1), .Cells(Depth, Width + 1)).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, Step:=1, Trend:=False
        
         'Populate LOADS headers if any are missing'
        insertRow = Depth + 1
        For i = 0 To UBound(headers)
            If .Range("A:A").Find(headers(i)) Is Nothing Then
                Set findHdr = .Range("A:A").Find(anchors(i)).offset(2, 0)     'offset of 2 so that header will be inserted to leave 2 blank rows after anchor header'
                .Cells(insertRow, 1).Value = headers(i)                                                         'header'
                .Cells(insertRow, Width + 1).Value = findHdr.row + 0.2
                .Cells(insertRow + 1, 1).Value = "[END_" & Strings.Right(headers(i), Len(headers(i)) - 1)       'footer'
                .Cells(insertRow + 1, Width + 1).Value = findHdr.row + 0.4
                .Cells(insertRow + 2, Width + 1).Value = findHdr.row + 0.6       'first blank row before next header (after sort)'
                .Cells(insertRow + 3, Width + 1).Value = findHdr.row + 0.8       'second blank row before next header (after sort)'
                insertRow = insertRow + 4
                'Sort data to put missing LOADS sections in correct positions'
                Range(.Cells(1, 1), .Cells(insertRow - 1, Width + 1)).Sort key1:=.Cells(1, Width + 1), order1:=xlAscending, Header:=xlNo
                'Recreate sort key to eliminate decimals'
                Range(.Cells(1, Width + 1), .Cells(insertRow - 1, Width + 1)).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, Step:=1, Trend:=False
            End If
        Next i
        
        'iterate through 4 LOAD TYPE sections'
        For i = 0 To UBound(headers)
            Set findHdr = .Range("A:A").Find(headers(i))
                'Clear any existing entries'
                If InStr(findHdr.offset(1, 0), Strings.Right(headers(i), Len(headers(i)) - 1)) < 1 Then
                    Range(findHdr.offset(1, 0), findHdr.offset(1, 0).End(xlDown).offset(-1, 0)).EntireRow.ClearContents
                End If
                'Flag inserts required'
                insertRow = .Cells(Rows.Count, Width + 1).End(xlUp).row + 1
                Range(.Cells(insertRow, Width + 1), .Cells(insertRow + inserts(i) - 1, Width + 1)).Value = findHdr.row + 0.5
                'Write number of inserts as content'
                findHdr.offset(0, 1).Value = "<" & inserts(i) & ">"
        Next i
        'Sort data'
        Range(.Cells(1, 1), .Cells(insertRow + inserts(4) - 1, Width + 1)).Sort key1:=.Cells(1, Width + 1), order1:=xlAscending, Header:=xlNo
        'Clear sort key'
        .Columns(Width + 1).ClearContents
    End With
    
    BLCPadding = WorksheetFunction.Rept(" ", 32)
    LCLongPadding = WorksheetFunction.Rept(" ", 79)
    LCShortPadding = WorksheetFunction.Rept(" ", 32)
    LC_Decimals = Array(7, 8, 9, 26, 28, 30, 32, 34, 36, 38, 40, 42)
    LC_Padding = Array(25, 27, 29, 31, 33, 35, 37, 39, 41, 43)
 
    Members = Temp.Range("A:A").Find("[.MEMBERS_MAIN_DATA]").row
    EndMembers = Temp.Range("A:A").Find("[.END_MEMBERS_MAIN_DATA]").row
    
    Nodes = Temp.Range("A:A").Find("[NODES]").row
    EndNodes = Temp.Range("A:A").Find("[END_NODES]").row
 
    
    If shDiscrete_Loads.Range("A4") = "" Then
    Else
    '>>>>>>POINT LOADS<<<<<<'
    Set PointLoads_Header = Temp.Range("A:A").Find("[POINT_LOADS]") 'Finds "POINT LOADS" section of Text file'
        ReDim PointLoads_List(1 To inserts(1), 1 To 10)
        ManCount = 181 + (Range("nLm").Value + Range("nLv").Value) * 5
        
        a = 1
        For t = 1 To ManCount Step 5  'loops to capture each loading configuration. Loop move horizontally by setting "t" as a "column" variable'
            i = 1
            'populate array until there is no member label'
            Do Until Tables_PointLoads.Cells(i + StartRow, t) = ""
               Set X = Temp.Range("A" & Members, "A" & EndMembers).Find(What:=Tables_PointLoads.Cells(StartRow + i, t).Value, LookIn:=xlValues, LookAt:=xlPart)
               'match members label to table index'
                m = Application.WorksheetFunction.Match(X, Temp.Range("A" & Members, "A" & EndMembers), 0) - 1
                PointLoads_List(a, 1) = m
                    'converts load direction to RISA file format'
                    Select Case Tables_PointLoads.Cells(i + StartRow, t + 1)
                        Case "X": PointLoads_List(a, 2) = 88
                        Case "Y": PointLoads_List(a, 2) = 89
                        Case "Z": PointLoads_List(a, 2) = 90
                        Case "Mx": PointLoads_List(a, 2) = 125
                        Case "My": PointLoads_List(a, 2) = 126
                        Case "Mz": PointLoads_List(a, 2) = 127
                    End Select
                'Load magnitude'
                PointLoads_List(a, 3) = Strings.Format(Tables_PointLoads.Cells(i + StartRow, t + 2).Value, "0.000!")
                'Load location'
                PointLoads_List(a, 4) = Strings.Format(Tables_PointLoads.Cells(i + StartRow, t + 3) * -100, "0.0!")
                PointLoads_List(a, 5) = 0
                PointLoads_List(a, 6) = -1
                PointLoads_List(a, 7) = "0.000000!"
                PointLoads_List(a, 8) = 2
                PointLoads_List(a, 9) = "0.000000!"
                PointLoads_List(a, 10) = "0.000000;"
                i = i + 1
                a = a + 1
            Loop
        Next t
        'Write fully populated array to POINTS_LOADS section of sheet'
        Range(PointLoads_Header.offset(1, 0), PointLoads_Header.offset(inserts(1), 9)).Value = PointLoads_List
    End If
    
    '>>>>>>DISTRIBUTED LOADS<<<<<<'
    'Finds "DISTRIBUTED LOADS" section of Text file'
    Set DistributedLoads_Header = Temp.Range("A:A").Find("[DIRECT_DISTRIBUTED_LOADS]")
        ReDim DistributedLoads_List(1 To inserts(2), 1 To 14)

        a = 1
        'loops to capture each loading configuration. Loop move horizontally by setting "t" as a "column" variable'
        For t = 1 To 225 Step 7
            i = 1
            'populate array until there is no member label'
            Do Until Tables_DistributedLoads.Cells(i + StartRow, t) = ""
               Set X = Temp.Range("A" & Members, "A" & EndMembers).Find(What:=Tables_DistributedLoads.Cells(StartRow + i, t).Value, LookIn:=xlValues, LookAt:=xlPart)
               'match members label to table index'
                m = Application.WorksheetFunction.Match(X, Temp.Range("A" & Members, "A" & EndMembers), 0) - 1
                DistributedLoads_List(a, 1) = m
                'converts load direction to RISA file format'
                Select Case Tables_DistributedLoads.Cells(i + StartRow, t + 1)
                    Case "X": DistributedLoads_List(a, 2) = 88
                    Case "Y": DistributedLoads_List(a, 2) = 89
                    Case "Z": DistributedLoads_List(a, 2) = 90
                End Select
                'Load magnitude'
                DistributedLoads_List(a, 3) = Strings.Format(Tables_DistributedLoads.Cells(i + StartRow, t + 2).Value, "0.000!")
                DistributedLoads_List(a, 4) = Strings.Format(Tables_DistributedLoads.Cells(i + StartRow, t + 3).Value, "0.000!")
                'Load location'
                DistributedLoads_List(a, 5) = Strings.Format(Tables_DistributedLoads.Cells(i + StartRow, t + 4).Value * -100, "0.0!")
                DistributedLoads_List(a, 6) = Strings.Format(Tables_DistributedLoads.Cells(i + StartRow, t + 5).Value * -100, "0.0!")
                DistributedLoads_List(a, 7) = 0
                DistributedLoads_List(a, 8) = -1
                DistributedLoads_List(a, 9) = "0.000000!"
                DistributedLoads_List(a, 10) = "0.000000!"
                DistributedLoads_List(a, 11) = 2
                DistributedLoads_List(a, 12) = "0.000000!"
                DistributedLoads_List(a, 13) = "0.000000!"
                DistributedLoads_List(a, 14) = "0.000000;"
                i = i + 1
                a = a + 1
            Loop
        Next t
        'Write fully populated array to DIRECT_DISTRIBUTED_LOADS section of sheet'
        Range(DistributedLoads_Header.offset(1, 0), DistributedLoads_Header.offset(inserts(2), 13)).Value = DistributedLoads_List

     '>>>>>>AREA LOADS<<<<<<'
     If WorksheetFunction.CountA(Range("nArea").Value) = 0 Then
     Else
        'Finds "AREA LOADS" section of Text file'
        Set AreaLoads_Header = Temp.Range("A:A").Find("[AREA_LOADS]")
        ReDim AreaLoads_List(1 To inserts(3), 1 To 10)

        a = 1
        'loops to capture each loading configuration. Loop move horizontally by setting "t" as a "column" variable'
        For t = 1 To 15 Step 8
            i = 1
            Do Until Tables_AreaLoads.Cells(i + StartRow, t) = ""     'populate array until there is no member label'
               Set x1 = Temp.Range("A" & Nodes, "A" & EndNodes).Find(What:=Tables_AreaLoads.Cells(StartRow + i, t).Value, LookIn:=xlValues, LookAt:=xlPart)
                n1 = Application.WorksheetFunction.Match(x1, Temp.Range("A" & Nodes, "A" & EndNodes), 0) - 1 'match nodes label to table index'
                AreaLoads_List(a, 1) = n1
               Set x2 = Temp.Range("A" & Nodes, "A" & EndNodes).Find(What:=Tables_AreaLoads.Cells(StartRow + i, t + 1).Value, LookIn:=xlValues, LookAt:=xlPart)
                n2 = Application.WorksheetFunction.Match(x2, Temp.Range("A" & Nodes, "A" & EndNodes), 0) - 1 'match nodes label to table index'
                AreaLoads_List(a, 2) = n2
               Set x3 = Temp.Range("A" & Nodes, "A" & EndNodes).Find(What:=Tables_AreaLoads.Cells(StartRow + i, t + 2).Value, LookIn:=xlValues, LookAt:=xlPart)
                n3 = Application.WorksheetFunction.Match(x3, Temp.Range("A" & Nodes, "A" & EndNodes), 0) - 1 'match nodes label to table index'
                AreaLoads_List(a, 3) = n3
               Set x4 = Temp.Range("A" & Nodes, "A" & EndNodes).Find(What:=Tables_AreaLoads.Cells(StartRow + i, t + 3).Value, LookIn:=xlValues, LookAt:=xlPart)
                n4 = Application.WorksheetFunction.Match(x4, Temp.Range("A" & Nodes, "A" & EndNodes), 0) - 1 'match nodes label to table index'
                AreaLoads_List(a, 4) = n4
                AreaLoads_List(a, 5) = 89
                AreaLoads_List(a, 6) = 0
                AreaLoads_List(a, 7) = Tables_AreaLoads.Cells(i + StartRow, t + 6).Value
                AreaLoads_List(a, 8) = -1
                AreaLoads_List(a, 9) = 0
                AreaLoads_List(a, 10) = "1;"
                i = i + 1
                a = a + 1
            Loop
        Next t
        
        'Write fully populated array to AREA_LOADS section of sheet'
        Range(AreaLoads_Header.offset(1, 0), AreaLoads_Header.offset(inserts(3), 9)).Value = AreaLoads_List
        End If
    
        '>>>>>>POPULATES BASIC LOAD CASES TABLE<<<<<<'
    Set BLC_Header = Temp.Range("A:A").Find("[BASIC_LOAD_CASES]") 'Finds "BLC" section of Text file'
    For i = 1 To inserts(0)
        BLC_Header.offset(i) = i
        BLC_Header.offset(i, 1) = "^" & Strings.Left(shRISA_3D.Range("B" & 4 + i).Value & BLCPadding, 42) & "^" 'Populates whole table with "0". Populated respective values below'
        Range(BLC_Header.offset(i, 2), BLC_Header.offset(i, 8)).Value = 0
        Range(BLC_Header.offset(i, 9), BLC_Header.offset(i, 11)).Value = "0.000000!"
        BLC_Header.offset(i, 12) = "-1;"
    Next i
 

   With BLC_Header
         '>>>>>>WIND AND ICE WIND DISTRIBUTED LOADS<<<<<<'
        .offset(17, 5) = shDistributedLoadTables.Range("A3").End(xlDown).row - StartRow
        .offset(18, 5) = shDistributedLoadTables.Range("H3").End(xlDown).row - StartRow
        .offset(19, 5) = shDistributedLoadTables.Range("O3").End(xlDown).row - StartRow
        .offset(20, 5) = shDistributedLoadTables.Range("V3").End(xlDown).row - StartRow
        .offset(21, 5) = shDistributedLoadTables.Range("AC3").End(xlDown).row - StartRow
        .offset(22, 5) = shDistributedLoadTables.Range("AJ3").End(xlDown).row - StartRow
        .offset(23, 5) = shDistributedLoadTables.Range("AQ3").End(xlDown).row - StartRow
        .offset(24, 5) = shDistributedLoadTables.Range("AX3").End(xlDown).row - StartRow
        .offset(25, 5) = shDistributedLoadTables.Range("BE3").End(xlDown).row - StartRow
        .offset(26, 5) = shDistributedLoadTables.Range("BL3").End(xlDown).row - StartRow
        .offset(27, 5) = shDistributedLoadTables.Range("BS3").End(xlDown).row - StartRow
        .offset(28, 5) = shDistributedLoadTables.Range("BZ3").End(xlDown).row - StartRow
        .offset(29, 5) = shDistributedLoadTables.Range("CG3").End(xlDown).row - StartRow
        .offset(30, 5) = shDistributedLoadTables.Range("CN3").End(xlDown).row - StartRow
        .offset(31, 5) = shDistributedLoadTables.Range("CU3").End(xlDown).row - StartRow
        .offset(32, 5) = shDistributedLoadTables.Range("DB3").End(xlDown).row - StartRow

        .offset(49, 5) = shDistributedLoadTables.Range("DI3").End(xlDown).row - StartRow
        .offset(50, 5) = shDistributedLoadTables.Range("DP3").End(xlDown).row - StartRow
        .offset(51, 5) = shDistributedLoadTables.Range("DW3").End(xlDown).row - StartRow
        .offset(52, 5) = shDistributedLoadTables.Range("ED3").End(xlDown).row - StartRow
        .offset(53, 5) = shDistributedLoadTables.Range("EK3").End(xlDown).row - StartRow
        .offset(54, 5) = shDistributedLoadTables.Range("ER3").End(xlDown).row - StartRow
        .offset(55, 5) = shDistributedLoadTables.Range("EY3").End(xlDown).row - StartRow
        .offset(56, 5) = shDistributedLoadTables.Range("FF3").End(xlDown).row - StartRow
        .offset(57, 5) = shDistributedLoadTables.Range("FM3").End(xlDown).row - StartRow
        .offset(58, 5) = shDistributedLoadTables.Range("FT3").End(xlDown).row - StartRow
        .offset(59, 5) = shDistributedLoadTables.Range("GA3").End(xlDown).row - StartRow
        .offset(60, 5) = shDistributedLoadTables.Range("GH3").End(xlDown).row - StartRow
        .offset(61, 5) = shDistributedLoadTables.Range("GO3").End(xlDown).row - StartRow
        .offset(62, 5) = shDistributedLoadTables.Range("GV3").End(xlDown).row - StartRow
        .offset(63, 5) = shDistributedLoadTables.Range("HC3").End(xlDown).row - StartRow
        .offset(64, 5) = shDistributedLoadTables.Range("HJ3").End(xlDown).row - StartRow
    
          '>>>>>>ICE DEAD DISTRIBUTED LOADS<<<<<<'
        .offset(66, 5) = shDistributedLoadTables.Range("HQ3").End(xlDown).row - StartRow
 
   End With


    With BLC_Header
         '>>>>>>DEAD AND ICE DEAD POINT LOADS<<<<<<'
        .offset(65, 3) = shPointLoadTables.Range("FE3").End(xlDown).row - StartRow
        .offset(66, 3) = shPointLoadTables.Range("FJ3").End(xlDown).row - StartRow
        
                
        '>>>>>>WIND AND ICE WIND POINT LOADS<<<<<<'
        .offset(1, 3) = shPointLoadTables.Range("A3").End(xlDown).row - StartRow
        .offset(2, 3) = shPointLoadTables.Range("F3").End(xlDown).row - StartRow
        .offset(3, 3) = shPointLoadTables.Range("K3").End(xlDown).row - StartRow
        .offset(4, 3) = shPointLoadTables.Range("P3").End(xlDown).row - StartRow
        .offset(5, 3) = shPointLoadTables.Range("U3").End(xlDown).row - StartRow
        .offset(6, 3) = shPointLoadTables.Range("Z3").End(xlDown).row - StartRow
        .offset(7, 3) = shPointLoadTables.Range("AE3").End(xlDown).row - StartRow
        .offset(8, 3) = shPointLoadTables.Range("AJ3").End(xlDown).row - StartRow
        .offset(9, 3) = shPointLoadTables.Range("AO3").End(xlDown).row - StartRow
        .offset(10, 3) = shPointLoadTables.Range("AT3").End(xlDown).row - StartRow
        .offset(11, 3) = shPointLoadTables.Range("AY3").End(xlDown).row - StartRow
        .offset(12, 3) = shPointLoadTables.Range("BD3").End(xlDown).row - StartRow
        .offset(13, 3) = shPointLoadTables.Range("BI3").End(xlDown).row - StartRow
        .offset(14, 3) = shPointLoadTables.Range("BN3").End(xlDown).row - StartRow
        .offset(15, 3) = shPointLoadTables.Range("BS3").End(xlDown).row - StartRow
        .offset(16, 3) = shPointLoadTables.Range("BX3").End(xlDown).row - StartRow
        
        .offset(33, 3) = shPointLoadTables.Range("CC3").End(xlDown).row - StartRow
        .offset(34, 3) = shPointLoadTables.Range("CH3").End(xlDown).row - StartRow
        .offset(35, 3) = shPointLoadTables.Range("CM3").End(xlDown).row - StartRow
        .offset(36, 3) = shPointLoadTables.Range("CR3").End(xlDown).row - StartRow
        .offset(37, 3) = shPointLoadTables.Range("CW3").End(xlDown).row - StartRow
        .offset(38, 3) = shPointLoadTables.Range("DB3").End(xlDown).row - StartRow
        .offset(39, 3) = shPointLoadTables.Range("DG3").End(xlDown).row - StartRow
        .offset(40, 3) = shPointLoadTables.Range("DL3").End(xlDown).row - StartRow
        .offset(41, 3) = shPointLoadTables.Range("DQ3").End(xlDown).row - StartRow
        .offset(42, 3) = shPointLoadTables.Range("DV3").End(xlDown).row - StartRow
        .offset(43, 3) = shPointLoadTables.Range("EA3").End(xlDown).row - StartRow
        .offset(44, 3) = shPointLoadTables.Range("EF3").End(xlDown).row - StartRow
        .offset(45, 3) = shPointLoadTables.Range("EK3").End(xlDown).row - StartRow
        .offset(46, 3) = shPointLoadTables.Range("EP3").End(xlDown).row - StartRow
        .offset(47, 3) = shPointLoadTables.Range("EU3").End(xlDown).row - StartRow
        .offset(48, 3) = shPointLoadTables.Range("EZ3").End(xlDown).row - StartRow


        '>>>>>>SEISMIC LOADS<<<<<<'
'        If Sheets("Loads and Equipment").Range("Code") = "H" Then
        .offset(67, 3) = shRISA_3D.Range("H71")
        .offset(68, 3) = shRISA_3D.Range("H72")
        .offset(69, 3) = shRISA_3D.Range("H73")
'        End If
        
        '>>>>>>MAINTENANCE LOADS<<<<<<'
        .offset(86, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("GD4:GD19"))
        .offset(87, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("GI4:GI19"))
        .offset(88, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("GN4:GN19"))
        .offset(89, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("GS4:GS19"))
        .offset(90, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("GX4:GX19"))
        .offset(91, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("HC4:HC19"))
        .offset(92, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("HH4:HH19"))
        .offset(93, 3) = WorksheetFunction.CountA(shPointLoadTables.Range("HM4:HM19"))

        If shCode.Range("QtySectors") > 3 Then
        .offset(45, 3) = shPointLoadTables.Range("DL3").End(xlDown).row - StartRow
        .offset(49, 3) = shPointLoadTables.Range("DQ3").End(xlDown).row - StartRow
        End If
        
        .offset(65, 10) = "-1.000000!"
        .offset(65, 8) = "1"
        .offset(66, 8) = "5"
    
        .offset(70, 9) = shRISA_3D.Range("D74")
        .offset(71, 9) = shRISA_3D.Range("D75")
        .offset(72, 9) = shRISA_3D.Range("D76")
        .offset(73, 9) = shRISA_3D.Range("D77")
        .offset(74, 9) = shRISA_3D.Range("D78")
        .offset(75, 9) = shRISA_3D.Range("D79")
        .offset(76, 9) = shRISA_3D.Range("D80")
        .offset(77, 9) = shRISA_3D.Range("D81")
        .offset(78, 9) = shRISA_3D.Range("D82")
        .offset(79, 9) = shRISA_3D.Range("D83")
        .offset(80, 9) = shRISA_3D.Range("D84")
        .offset(81, 9) = shRISA_3D.Range("D85")
        .offset(82, 9) = shRISA_3D.Range("D86")
        .offset(83, 9) = shRISA_3D.Range("D87")
        .offset(84, 9) = shRISA_3D.Range("D88")
        .offset(85, 9) = shRISA_3D.Range("D89")
                
        .offset(70, 10) = shRISA_3D.Range("E74")
        .offset(71, 10) = shRISA_3D.Range("E75")
        .offset(72, 10) = shRISA_3D.Range("E76")
        .offset(73, 10) = shRISA_3D.Range("E77")
        .offset(74, 10) = shRISA_3D.Range("E78")
        .offset(75, 10) = shRISA_3D.Range("E79")
        .offset(76, 10) = shRISA_3D.Range("E80")
        .offset(77, 10) = shRISA_3D.Range("E81")
        .offset(78, 10) = shRISA_3D.Range("E82")
        .offset(79, 10) = shRISA_3D.Range("E83")
        .offset(80, 10) = shRISA_3D.Range("E84")
        .offset(81, 10) = shRISA_3D.Range("E85")
        .offset(82, 10) = shRISA_3D.Range("E86")
        .offset(83, 10) = shRISA_3D.Range("E87")
        .offset(84, 10) = shRISA_3D.Range("E88")
        .offset(85, 10) = shRISA_3D.Range("E89")
                
        .offset(70, 11) = shRISA_3D.Range("F74")
        .offset(71, 11) = shRISA_3D.Range("F75")
        .offset(72, 11) = shRISA_3D.Range("F76")
        .offset(73, 11) = shRISA_3D.Range("F77")
        .offset(74, 11) = shRISA_3D.Range("F78")
        .offset(75, 11) = shRISA_3D.Range("F79")
        .offset(76, 11) = shRISA_3D.Range("F80")
        .offset(77, 11) = shRISA_3D.Range("F81")
        .offset(78, 11) = shRISA_3D.Range("F82")
        .offset(79, 11) = shRISA_3D.Range("F83")
        .offset(80, 11) = shRISA_3D.Range("F84")
        .offset(81, 11) = shRISA_3D.Range("F85")
        .offset(82, 11) = shRISA_3D.Range("F86")
        .offset(83, 11) = shRISA_3D.Range("F87")
        .offset(84, 11) = shRISA_3D.Range("F88")
        .offset(85, 11) = shRISA_3D.Range("F89")

        '>>>>>>AREA LOADS<<<<<<'
        If WorksheetFunction.CountA(shMaintenance_Loads.Range("L7:L10")) <> 0 Then
            .offset(65, 6) = WorksheetFunction.CountA(shAreaLoadTables.Range("A4:A7"))
            .offset(66, 6) = WorksheetFunction.CountA(shAreaLoadTables.Range("I4:I7"))
        Else
            .offset(65, 6) = 0
            .offset(66, 6) = 0
        End If
    
    End With

    'Range(BLC_Header.Offset(1, 5), BLC_Header.Offset(4, 5)).Value = Lastrow + 1 - FirstRow
 
    '>>>>>>POPULATES LOAD COMBINATIONS TABLE<<<<<<'
    'Finds "LC" section of Text file'
    Set LC_Header = Temp.Range("A:A").Find("[LOAD_COMBINATIONS]")
    cnt = inserts(4)
    
    Range(LC_Header.offset(1), LC_Header.offset(cnt)).Value = shRISA_3D.Range("N5", "N" & cnt + 4).Value
    LC_Header.offset(1, 1).Resize(cnt) = 0
    For r = 5 To cnt + 5
        If shRISA_3D.Range("O" & r) = "True" Then
            LC_Header.offset(1 + r - 5, 3) = 1
        Else
            LC_Header.offset(1 + r - 5, 3) = 0
        End If
        If shRISA_3D.Range("P" & r) = "Y" Then
            LC_Header.offset(1 + r - 5, 2) = 1
        Else
            LC_Header.offset(1 + r - 5, 2) = 0
        End If
    Next r
    LC_Header.offset(1, 4).Resize(cnt, 5) = 0
    LC_Header.offset(1, 6).Resize(cnt, 3) = 1
    LC_Header.offset(1, 9).Resize(cnt, 2) = 0
    LC_Header.offset(1, 11).Resize(cnt, 7) = 1
    LC_Header.offset(1, 18).Resize(cnt) = -1
    LC_Header.offset(1, 19).Resize(cnt, 5) = 1
    
    For i = 1 To cnt
        Range(LC_Header.offset(i, 24), LC_Header.offset(i, 33)).Value = Range(shRISA_3D.Range("R" & 4 + i), shRISA_3D.Range("AA" & 4 + i)).Value
    Next i
    
    'Copy LOAD_COMBINATIONS to array for more efficient updating of decimals and padding'
    Set LC_Range = Range(LC_Header.offset(1, 0), LC_Header.offset(cnt, 43))
    LC_List = LC_Range.Value
    For i = 1 To cnt
        LC_List(i, 1) = "^" & Strings.Left(LC_List(i, 1) & LCLongPadding, 79) & "^"
        For a = 0 To UBound(LC_Padding)
            LC_List(i, LC_Padding(a)) = "^" & Strings.Left(LC_List(i, LC_Padding(a)) & LCShortPadding, 32) & "^"
        Next a
        For a = 0 To UBound(LC_Decimals)
            LC_List(i, LC_Decimals(a)) = Strings.Format(LC_List(i, LC_Decimals(a)), "0.000000!")
        Next a
    Next i
    
    'Write updated array back to sheet'
    LC_Range.Value = LC_List
    
    'Populate any residual blank cells with decimal 0'
    On Error Resume Next
    LC_Range.SpecialCells(xlCellTypeBlanks).Value = "0.000000!"
    On Error GoTo 0
  
    Dim f As Range, firstAddr As String, tgt As Range

    With Temp.Columns("A")
        Set f = .Find(What:="[WOOD_SCHEDULE", LookIn:=xlValues, LookAt:=xlPart, _
                      SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        If Not f Is Nothing Then
            firstAddr = f.Address
            Do
                ' Require ":" in adjacent Column B (allowing stray spaces)
                If Trim$(CStr(f.offset(0, 1).Value)) = ":" Then
                    Set tgt = f.offset(1, 4)   ' one row down, 4 cols right (col E relative to A)
                    If Not IsError(tgt.Value) And Len(CStr(tgt.Value)) > 0 Then
                        tgt.NumberFormat = "@"                    ' store as Text
                        tgt.Value = Format$(CDbl(tgt.Value), "0.000")  ' write as text "0.000"
                    End If
                End If
                Set f = .FindNext(f)
            Loop While Not f Is Nothing And f.Address <> firstAddr
        End If
    End With
    
    Call text_to_RISA
    
    Start = Timer - Start
    Debug.Print Start
    
Handle:
    Select Case Err.Number
        Case 1004
        AppActivate Application.Caption
        MsgBox "It appears that member label """ & Tables_DistributedLoads.Cells(StartRow + i, t) & """ does not exist in the model. Please fix Appurtenance/Dish Table and Re-calcuate Point Loads.", vbInformation, "Error!"
        Exit Sub
    End Select
    
CleanExit:
    'RESTORE
    Application.StatusBar = prevSB
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevSU
    Exit Sub

CleanFail:
    'optional: MsgBox Err.Description
    Resume CleanExit

End Sub


Sub Maintenance_Only()

   Dim i As Long
   
   For i = 27 To 147

End Sub


Sub text_to_RISA()
 
Dim ExpRange As Range
Dim ExpRow As Range
Dim ExpCell As Range
Dim myValue As Variant
Dim Width As Integer
Dim exclaimPurge As Boolean
Dim RISAModel As String
Dim Answer As Integer
Dim Updated_Model As Byte
Dim RISA As Object
Dim Company As Range, ModelTitle As Range, Designer As Range, Job As Range, Checker As Range
Dim Company_End As Range, ModelTitle_End As Range, Designer_End As Range, Job_End As Range, Checker_End As Range

RISAModel = Range("Carrier_SiteNumber") & "_" & CleanFileName(Range("Carrier_SiteName")) & "_" & Range("Carrier") & "_" & Range("Standard") & "_" & Range("Method") & "_" & Format(Date, "YYYY-MM-DD") & "_" & Format(Now, "hh-mm-ss") & ".r3d"
RISAModel = Application.GetSaveAsFilename(RISAModel, "RISA-3D (*.r3d), *.r3d")

Open RISAModel For Output As #1

Sheets("Risa File").Activate
Set ExpRange = Sheets(Sheets.Count).UsedRange.Rows
Set Company = Sheets("Risa File").Range("A:A").Find("[..COMPANY_NAME]")
Set Company_End = Sheets("Risa File").Range("A:A").Find("[..END_COMPANY_NAME]")
Set ModelTitle = Sheets("Risa File").Range("A:A").Find("[..MODEL_TITLE]")
Set ModelTitle_End = Sheets("Risa File").Range("A:A").Find("[..END_MODEL_TITLE]")
Set Designer = Sheets("Risa File").Range("A:A").Find("[..DESIGNER_NAME]")
Set Designer_End = Sheets("Risa File").Range("A:A").Find("[..END_DESIGNER_NAME]")
Set Job = Sheets("Risa File").Range("A:A").Find("[..JOB_NUMBER]")
Set Job_End = Sheets("Risa File").Range("A:A").Find("[...END_JOB_NUMBER]")
'Set Checker = Sheets("Risa File").Range("A:A").Find("[..CHECKER_NAME]")                'Disabled
For Each ExpRow In ExpRange
        'Iterate only to right-most cell in this row'
        Width = Cells(ExpRow.row, Columns.Count).End(xlToLeft).Column
        'For Each ExpCell In ExpRow.Cells'
         For Each ExpCell In ExpRow.Resize(, Width).Cells
            If InStr(ExpCell, "POINT_LOADS]") Or InStr(ExpCell, "DIRECT_DISTRIBUTED_LOADS]") Or _
                InStr(ExpCell, "BASIC_LOAD_CASES]") Or InStr(ExpCell, "LOAD_COMBINATIONS]") Then exclaimPurge = Not exclaimPurge
            If Strings.Right(ExpCell, 1) = ";" Then
                ExpCell.Replace What:=";", Replacement:=""
            End If
            If InStr(ExpCell, "[") <> 0 Then
                myValue = ExpCell.Value
            ElseIf InStr(ExpCell, "]") <> 0 Then
                myValue = myValue & ExpCell.Value
            ElseIf InStr(ExpCell, ", ") <> 0 Then
                myValue = myValue & ExpCell.Value & " "
            ElseIf InStr(ExpCell, ":") <> 0 Then
                myValue = myValue & " " & ExpCell.Value & " "
            ElseIf InStr(ExpCell, "<") <> 0 Then
                myValue = myValue & " " & ExpCell.Value
            ElseIf IsEmpty(ExpCell) Then
                myValue = myValue
            ElseIf ModelTitle.offset(1, 0) = ExpCell Then
               With ModelTitle
                ModelTitle.offset(1, 0) = shCode.Range("Carrier_SiteNumber").Value & " - " & shCode.Range("Carrier_SiteName").Value
                ModelTitle.offset(1, 1).Resize(1, 10).ClearContents
                myValue = shCode.Range("Carrier_SiteNumber").Value & " - " & shCode.Range("Carrier_SiteName").Value
'                   .Offset(1, 0) = Strings.RTrim(.Offset(1, 0) & " " & .Offset(1, 1) & " " & .Offset(1, 2) & " " & .Offset(1, 3) & " " & .Offset(1, 4) & " " & .Offset(1, 5))
'                   .Offset(1, 1).Resize(1, 10).ClearContents
'                   myValue = Strings.RTrim(.Offset(1, 0))
               End With
            ElseIf Company.offset(1, 0) = ExpCell Then
                If Left(Sheets("Code").Range("B2"), 3) = "TKK" Then
                    Company.offset(1, 0) = "TKK Engineering"
                Else
                    Company.offset(1, 0) = "NB+C ES"
                End If
                Company.offset(1, 1).Resize(1, 10).ClearContents
                If Left(Sheets("Code").Range("B2"), 3) = "TKK" Then
                    myValue = "TKK Engineering"
                Else
                    myValue = "NB+C ES"
                End If
            ElseIf Designer.offset(1, 0) = ExpCell Then
               With Designer
                Designer.offset(1, 0) = "NB+C Engineer" 'LastAuthor
                Designer.offset(1, 1).Resize(1, 10).ClearContents
                myValue = "NB+C Engineer" 'LastAuthor
'                Designer.Offset(1, 0) = "NB+C Engineer"
'                Designer.Offset(1, 1).Resize(1, 10).ClearContents
'                myValue = "NB+C Engineer"
'                   .Offset(1, 0) = Strings.RTrim(.Offset(1, 0) & " " & .Offset(1, 1))
'                   .Offset(1, 1).Resize(1, 3).ClearContents
'                   myValue = Strings.RTrim(.Offset(1, 0))
               End With
            ElseIf Job.offset(1, 0) = ExpCell Then
                Job.offset(1, 0) = shCode.Range("ProjectNumber").Value
                Job.offset(1, 1).Resize(1, 10).ClearContents
                myValue = shCode.Range("ProjectNumber").Value
                myValue = Job.offset(1, 0)
            ElseIf Not Checker Is Nothing Then
               If Checker.offset(1, 0) = ExpCell Then
               With Checker
                Checker.offset(1, 0) = "NB+C"
                Checker.offset(1, 1).Resize(1, 10).ClearContents
                myValue = "RIM"
'                     .Offset(1, 0) = Strings.RTrim(.Offset(1, 0) & " " & .Offset(1, 1))
'                     .Offset(1, 1).Resize(1, 3).ClearContents
'                      myValue = Strings.RTrim(.Offset(1, 0))
                  End With
                Else
                    myValue = myValue & ExpCell.Value & " "
               End If
            Else
                If Application.IsNumber(ExpCell.Value) Then
                    myValue = myValue & ExpCell.Value & " "
                Else
                    myValue = myValue & """" & ExpCell.Value & """" & " "
                End If
            End If
        Next ExpCell
              
        If myValue = "" Then
            Print #1, myValue
        ElseIf Strings.Right(myValue, 1) = "]" Or Strings.Right(myValue, 1) = ">" Then
            Print #1, myValue
        ElseIf Company.offset(1, 0) = myValue Then
            Print #1, myValue
        ElseIf ModelTitle.offset(1, 0) = myValue Then
            Print #1, myValue
        ElseIf Designer.offset(1, 0) = myValue Then
            Print #1, myValue
        ElseIf Job.offset(1, 0) = myValue Then
            Print #1, Strings.LTrim(myValue)
        ElseIf Not Checker Is Nothing Then
            If Checker.offset(1, 0) = myValue Then
                Print #1, myValue
            Else
                'Print #1, Strings.Left(myValue, Strings.Len(myValue) - 1) & ";"   'added this line''
                GoTo Purge
           End If
        Else
Purge:
            If exclaimPurge Then
            
                'remove ! from numbers which ensured 6 decimals'
                myValue = Strings.Replace(myValue, "!", "")
                
                'remove quotes in all 4 sections'
                myValue = Strings.Replace(myValue, """", "")
                
                're-instate quotes around where supposed to be'
                myValue = Strings.Replace(myValue, "^", """")
            End If
            Print #1, Strings.Left(myValue, Strings.Len(myValue) - 1) & ";"
        End If
        myValue = ""
    Next ExpRow
Close #1
   AppActivate Application.Caption
    Answer = MsgBox("Loads have successfully been exported to a newly created RISA-3D model." & vbNewLine & vbNewLine & "Would you like to open the newly created RISA-3D model?", vbYesNo + vbDefaultButton1 + vbQuestion, "Export Complete")
    If Answer = 6 Then
        Set RISA = CreateObject("shell.application")
        Updated_Model = RISA.shellExecute(RISAModel)
    Else
    End If
    
    shGeometry.Activate
    shGeometry.Range("B6").Select
    Selection.TextToColumns Destination:=Range("B6"), dataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=False, Tab:=True, Semicolon _
        :=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, _
        1), TrailingMinusNumbers:=True

    
shCode.Activate


End Sub



