Attribute VB_Name = "Import_Members"
Option Explicit
'This Module contains the procedure used to import member information to
'the "Risa 3D Members" and "Section Sets" tables on the "Import Model"
'tab. The procedure imports the .txt version of the .r3d file in to excel
'and the finds specific headers to extract the appropriate data.
'This procedure is ONLY good for Hot Rolled and Cold Formed Steel.


Sub Import()
    Dim Answer As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    
        Answer = MsgBox("Before importing a RISA-3D Model:" & vbNewLine & vbNewLine & _
        "Ensure RISA-3D Model Member and Section Set labels are free from problematic characters, such as quotation marks." & vbNewLine & vbNewLine & "Continue?", vbExclamation + vbYesNo + vbDefaultButton2, "WARNING!")
        If Answer = vbNo Then
            Exit Sub
        ElseIf Answer = vbYes Then
            With shGeometry
                .Unprotect
                .Range("B6:V99999").ClearContents
                .Range("X6:AA31").ClearContents
            End With
            
            Call Import_Model
            Call TextToColumn
            Call RowCount
            
            shGeometry.Range("B6").Select
            Call Geometry_Lookup
            Call Geometry_Lookup
                
            Application.ScreenUpdating = True
            
            If Application.Range("UserInputCheck") = True Or Application.Range("FalseCheck") = True Then
                Answer = MsgBox("Please input shape data in the RISA-3D Section Sets table on the Geometry tab.", vbCritical + vbOKOnly, "Missing Shape Data")
                shGeometry.Range("X6").Select
            End If
        End If
    Call ModelEPA

    Answer = MsgBox("The imported RISA-3D model was modeled using the following parameters. Please ensure the following is correct. Otherwise, update the RISA-3D model with the correct parameters and re-import." _
    & vbNewLine & vbNewLine & _
    "       " & "Unit Length:  " & shCode.Range("RISA3D.Unit.Length") & vbNewLine & _
    "       " & "Unit Force:  " & shCode.Range("RISA3D.Unit.Force") & vbNewLine & _
    "       " & "Unit Linear Force:  " & shCode.Range("RISA3D.Unit.LinearForce") & vbNewLine & _
    "       " & "Unit Moment:  " & shCode.Range("RISA3D.Unit.Moment") & vbNewLine & _
    "       " & "Unit Area Load:  " & shCode.Range("RISA3D.Unit.AreaLoad") & vbNewLine & _
    "       " & "Model Axes:  " & shCode.Range("Axes") & vbNewLine & _
    "       " & "Steel Design Standard:  " & shCode.Range("RISA3D.AISC") & vbNewLine & _
    "       " & "RISA-3D Version:  " & shCode.Range("RISA3D.Version") & vbNewLine & _
    vbNewLine & "Please note that any combination or system of units may be used. Unit conversions are included in all calculations for all systems of units.", _
    vbInformation + vbOKOnly, "RISA-3D model was successfully imported!")

    '--- Prompt for up to 4 mount azimuths (comma-separated; blanks allowed)
    Dim s As String, parts() As String
    Dim targets As Variant, i As Long
    Dim v As String
    
    targets = Array( _
        Sheets("Discrete Loads").Range("M2"), _
        Sheets("Discrete Loads").Range("U2"), _
        Sheets("Discrete Loads").Range("AC2"), _
        Sheets("Discrete Loads").Range("AK2") _
    )
    
    s = Application.InputBox( _
        Prompt:="Enter up to four Mount Azimuth(s) separated by commas." & vbCrLf & _
                "Blanks allowed to skip a position (use double commas)." & vbCrLf & _
                "Examples:" & vbCrLf & _
                "  0, 120, 240" & vbCrLf & _
                "  0,, 240" & vbCrLf & _
                "  0, 90,, 270", _
        title:="Mount Azimuth(s)", _
        Type:=2 _
    )
    
    If s = "False" Then Exit Sub  'user cancelled
    
    'Clear existing cells first (optional but usually preferred)
    For i = LBound(targets) To UBound(targets)
        targets(i).Value = vbNullString
    Next i
    
    'Split into up to 4 slots; keep blanks if user types ",,"
    parts = Split(s, ",")
    
    For i = 0 To 3
        If i <= UBound(parts) Then
            v = Trim$(parts(i))
            If Len(v) > 0 Then
                'Optional: validate numeric
                If IsNumeric(v) Then
                    targets(i).Value = CDbl(v)
                Else
                    MsgBox "Invalid azimuth in position " & (i + 1) & ": '" & v & "'", vbExclamation
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True

End Sub


Sub Import_Model()
    Dim fileName As Variant
    Dim M_Header As Range, HR_Header As Range, CF_Header As Range, AL_Header As Range                   'Members Header, Hot_Rolled Header, Colf_Formed Header
    Dim Nodes_Header As Range, FirstNode As String, SecondNode As String, Units As Range
    Dim Members() As String, Member_Section() As String, HR_Section() As String, CF_Section() As String         'Members, Hot_Rolled_Sections_Sets, Cold_Formed_Section_Sets
    Dim HR_cnt As Integer, CF_cnt As Integer, AL_cnt, Nodes_cnt As Integer, M_cnt As Integer, cnt As Integer  'Hot_Rolled Section Set counter, Cold_Formed Section Set counter, Members counter
    Dim DesignCode As Range, Axis As Range
    Dim r As Long, n As Long, t As Byte
    Dim a As Byte, b As Byte, c As Byte
    Dim m, lastRow As Double
    Dim i, v As Integer
    Dim Temp As Worksheet                                                           'Tempoaray file imported from RISA 3D
    Dim Openbook As Workbook
    Dim Project As String, Folder As String
    Dim RisaFile As Worksheet
    Dim Answer As String
    Dim RISA_Version As Range
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New FileSystemObject
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
       
    Set RisaFile = Sheets("Risa File")
    
    On Error GoTo 0
    
    If Not RisaFile Is Nothing Then     'incase user decides to update analysis from a previous run.
        
        Worksheets("Risa File").Delete  'if user selects "yes', this deletes the previously used file
        GoTo OpenFile
    End If
    
OpenFile:
    fileName = Application.GetOpenFilename("RISA-3D Model (*.r3d),*.r3d")
        If fileName = False Then    'If user hits "Cancel" button
            Exit Sub
        End If
    Application.Workbooks.OpenText fileName, dataType:=xlDelimited, Space:=True
    
    shCode.Range("ModelPath") = fso.GetAbsolutePathName(fileName)
    Set Openbook = ActiveWorkbook
    Openbook.Sheets(1).Cells.Copy
    ThisWorkbook.Activate
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Worksheets(Sheets.Count)
    ActiveSheet.Name = "Risa File"
    ActiveSheet.Range("A1").PasteSpecial (xlPasteValues)
    Range("A1").Select
    ActiveSheet.Visible = True
    Application.CutCopyMode = False
    Openbook.Close False
    
'==========================================================================================================================================='

    m = shGeometry.Cells(Rows.Count, "B").End(xlUp).row - 5
        If m >= 0 Then
            shGeometry.Range("B6", "C" & m + 6).ClearContents        'clears Members
        End If
    m = shGeometry.Cells(Rows.Count, "U").End(xlUp).row - 5
        If m >= 0 Then
            shGeometry.Range("X6", "Y" & m + 6).ClearContents        'clears Section Set Label
            shGeometry.Range("X6", "Z" & m + 6).ClearContents        'clears Section Set Shape
        End If

'==========================================================================================================================================='

'++++Copy Members & Section Sets from Risa file++++
    Set Temp = Worksheets(Sheets.Count)
    Set M_Header = Temp.Range("A:A").Find("[.MEMBERS_MAIN_DATA]")           'Members labels
    Set Nodes_Header = Temp.Range("A:A").Find("[Nodes]")                    'Node Labels
    Set HR_Header = Temp.Range("A:A").Find("[.HR_STEEL_SECTION_SETS]")      'Hot Rolled Section Sets
    Set AL_Header = Temp.Range("A:A").Find("[.ALUMINUM_SECTION_SETS]")      'Aluminum Sections Sets
    Set CF_Header = Temp.Range("A:A").Find("[.CF_STEEL_SECTION_SETS]")      'Cold Formed Sections Sets
    Set Units = Temp.Range("A:A").Find("[UNITS]")                           'Units
    Set DesignCode = Temp.Range("A:A").Find("[.DESIGN_CODES]")              'Design code
    Set RISA_Version = Temp.Range("A:A").Find("[VERSION_NUMBER]")           'RISA-3D Version
    Set Axis = Temp.Range("A:A").Find("[.SOLUTION_PARAMETERS]")             'Axis

    cnt = Range(M_Header, M_Header.End(xlDown)).Rows.Count - 2

    For r = 1 To cnt
        If M_Header.offset(r, 10) = 1 Or M_Header.offset(r, 10) = 2 Or M_Header.offset(r, 10) = 6 Then    'Counts number of members if members are Hot Rolled or Cold Formed (i.e., No Rigid Links). 1=Hot Rolled, 2=Cold Formed
'        If M_Header.Offset(r, 10) <> 7 Then   'Counts number of members if members are Hot Rolled or Cold Formed (i.e., No Rigid Links). 1=Hot Rolled, 2=Cold Formed
            n = n + 1
        End If
    Next r
    ReDim Members(1 To n, 1 To 3)       'Declared size or Member's array

    n = 1: a = 1: b = 1: c = 1
    Nodes_Header.offset(1, 0).End(xlToRight).SpecialCells(xlCellTypeBlanks).Delete (xlToLeft)
    For r = 1 To cnt
        If M_Header.offset(r, 10) = 0 Then
        ElseIf M_Header.offset(r, 10) = 1 Then              'Populates Members array if Hot Rolled
            Members(n, 1) = RTrim(M_Header.offset(r))
            Members(n, 2) = RTrim(Application.WorksheetFunction.Index(Temp.Range("A" & HR_Header.row + 1, "A" & HR_Header.End(xlDown).row - 1), M_Header.offset(r, 7) + 1))
            
            'Find coordinates of Node to determine Members Lengths
            Dim n1 As Variant, n2 As Variant
            Dim x1 As Double, y1 As Double, z1 As Double
            Dim x2 As Double, y2 As Double, z2 As Double
            Dim l As Double
            Dim W As Double
            
            'Nodes_Header.Offset(1, 0).End(xlToRight).SpecialCells(xlCellTypeBlanks).Delete (xlToLeft)
            
            n1 = M_Header.offset(r, 3)
            n1 = Application.WorksheetFunction.Index(Temp.Range("A" & Nodes_Header.row + 1, "A" & Nodes_Header.End(xlDown).row - 1), n1)
            n1 = Application.WorksheetFunction.Match(n1, Temp.Range("A" & Nodes_Header.row + 1, "A" & Nodes_Header.End(xlDown).row - 1), 0)
            n2 = M_Header.offset(r, 4)
            n2 = Application.WorksheetFunction.Index(Temp.Range("A" & Nodes_Header.row + 1, "A" & Nodes_Header.End(xlDown).row - 1), n2)
            n2 = Application.WorksheetFunction.Match(n2, Temp.Range("A" & Nodes_Header.row + 1, "A" & Nodes_Header.End(xlDown).row - 1), 0)
            x1 = Nodes_Header.offset(n1, 1)
            y1 = Nodes_Header.offset(n1, 2)
            z1 = Nodes_Header.offset(n1, 3)
            x2 = Nodes_Header.offset(n2, 1)
            y2 = Nodes_Header.offset(n2, 2)
            z2 = Nodes_Header.offset(n2, 3)
            l = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2 + (z2 - z1) ^ 2)
            'W =
            Members(n, 3) = l
            n = n + 1
            

            With shCode
            If Axis.offset(1, 6) = 1 Then
                .Range("RISA3D.Axis") = "XZY"
            ElseIf Axis.offset(1, 6) = 2 Then
                .Range("RISA3D.Axis") = "YXZ"
            ElseIf Axis.offset(1, 6) = 3 Then
                .Range("RISA3D.Axis") = "ZYX"
            End If
                
            If DesignCode.offset(1, 0) = 0 Then
                .Range("RISA3D.AISC") = "None"
                ElseIf DesignCode.offset(1, 0) = 406 Then
                .Range("RISA3D.AISC") = "AISC 15th (360-16): ASD"
                ElseIf DesignCode.offset(1, 0) = 405 Then
                .Range("RISA3D.AISC") = "AISC 15th (360-16): LRFD"
                ElseIf DesignCode.offset(1, 0) = 404 Then
                .Range("RISA3D.AISC") = "AISC 14th (360-10): ASD"
                ElseIf DesignCode.offset(1, 0) = 403 Then
                .Range("RISA3D.AISC") = "AISC 14th (360-10): LRFD"
                ElseIf DesignCode.offset(1, 0) = 304 Then
                .Range("RISA3D.AISC") = "AISC 13th (360-05): ASD"
                ElseIf DesignCode.offset(1, 0) = 303 Then
                .Range("RISA3D.AISC") = "AISC 13th (360-05): LRFD"
                ElseIf DesignCode.offset(1, 0) = 3 Then
                .Range("RISA3D.AISC") = "AISC 3rd: LRFD"
                ElseIf DesignCode.offset(1, 0) = 2 Then
                .Range("RISA3D.AISC") = "AISC 2nd: LRFD"
                ElseIf DesignCode.offset(1, 0) = 9 Then
                .Range("RISA3D.AISC") = "AISC 9th: ASD"
            End If
            
            If Units.offset(1, 1) = 0 Then
               .Range("RISA3D.Unit.Length") = "ft"
               ElseIf Units.offset(1, 1) = 1 Then
               .Range("RISA3D.Unit.Length") = "in"
               ElseIf Units.offset(1, 1) = 2 Then
               .Range("RISA3D.Unit.Length") = "m"
               ElseIf Units.offset(1, 1) = 3 Then
               .Range("RISA3D.Unit.Length") = "cm"
               ElseIf Units.offset(1, 1) = 4 Then
               .Range("RISA3D.Unit.Length") = "mm"
            End If
        
            If Units.offset(1, 4) = 0 Then
               .Range("RISA3D.Unit.Force") = "kip"
               ElseIf Units.offset(1, 4) = 1 Then
               .Range("RISA3D.Unit.Force") = "lbf"
               ElseIf Units.offset(1, 4) = 2 Then
               .Range("RISA3D.Unit.Force") = "kN"
               ElseIf Units.offset(1, 4) = 3 Then
               .Range("RISA3D.Unit.Force") = "N"
               ElseIf Units.offset(1, 4) = 5 Then
               .Range("RISA3D.Unit.Force") = "kg"
            End If
            
            If Units.offset(1, 5) = 0 Then
               .Range("RISA3D.Unit.LinearForce") = "klf"
               ElseIf Units.offset(1, 5) = 1 Then
               .Range("RISA3D.Unit.LinearForce") = "kli"
               ElseIf Units.offset(1, 5) = 2 Then
               .Range("RISA3D.Unit.LinearForce") = "plf"
               ElseIf Units.offset(1, 5) = 3 Then
               .Range("RISA3D.Unit.LinearForce") = "pli"
               ElseIf Units.offset(1, 5) = 4 Then
               .Range("RISA3D.Unit.LinearForce") = "kN/m"
               ElseIf Units.offset(1, 5) = 5 Then
               .Range("RISA3D.Unit.LinearForce") = "kN/cm"
               ElseIf Units.offset(1, 5) = 6 Then
               .Range("RISA3D.Unit.LinearForce") = "kN/mm"
               ElseIf Units.offset(1, 5) = 7 Then
               .Range("RISA3D.Unit.LinearForce") = "N/m"
               ElseIf Units.offset(1, 5) = 8 Then
               .Range("RISA3D.Unit.LinearForce") = "N/cm"
               ElseIf Units.offset(1, 5) = 9 Then
               .Range("RISA3D.Unit.LinearForce") = "N/mm"
               ElseIf Units.offset(1, 5) = 13 Then
               .Range("RISA3D.Unit.LinearForce") = "kg/m"
               ElseIf Units.offset(1, 5) = 14 Then
               .Range("RISA3D.Unit.LinearForce") = "kg/cm"
               ElseIf Units.offset(1, 5) = 15 Then
               .Range("RISA3D.Unit.LinearForce") = "kg/mm"
            End If

            If Units.offset(1, 7) = 0 Then
               .Range("RISA3D.Unit.Moment") = "kip-ft"
               ElseIf Units.offset(1, 7) = 1 Then
               .Range("RISA3D.Unit.Moment") = "kip-in"
               ElseIf Units.offset(1, 7) = 2 Then
               .Range("RISA3D.Unit.Moment") = "lbf-ft"
               ElseIf Units.offset(1, 7) = 3 Then
               .Range("RISA3D.Unit.Moment") = "lbf-in"
               ElseIf Units.offset(1, 7) = 4 Then
               .Range("RISA3D.Unit.Moment") = "kN-m"
               ElseIf Units.offset(1, 7) = 5 Then
               .Range("RISA3D.Unit.Moment") = "kN-cm"
               ElseIf Units.offset(1, 7) = 6 Then
               .Range("RISA3D.Unit.Moment") = "kN-mm"
               ElseIf Units.offset(1, 7) = 7 Then
               .Range("RISA3D.Unit.Moment") = "N-m"
               ElseIf Units.offset(1, 7) = 8 Then
               .Range("RISA3D.Unit.Moment") = "N-cm"
               ElseIf Units.offset(1, 7) = 9 Then
               .Range("RISA3D.Unit.Moment") = "N-mm"
               ElseIf Units.offset(1, 7) = 13 Then
               .Range("RISA3D.Unit.Moment") = "kg-m"
               ElseIf Units.offset(1, 7) = 14 Then
               .Range("RISA3D.Unit.Moment") = "kg-cm"
               ElseIf Units.offset(1, 7) = 15 Then
               .Range("RISA3D.Unit.Moment") = "kg-mm"
            End If
            
            If Units.offset(1, 9) = 0 Then
               .Range("RISA3D.Unit.AreaLoad") = "ksf"
               ElseIf Units.offset(1, 9) = 1 Then
               .Range("RISA3D.Unit.AreaLoad") = "ksi"
               ElseIf Units.offset(1, 9) = 2 Then
               .Range("RISA3D.Unit.AreaLoad") = "psf"
               ElseIf Units.offset(1, 9) = 3 Then
               .Range("RISA3D.Unit.AreaLoad") = "psi"
               ElseIf Units.offset(1, 9) = 4 Then
               .Range("RISA3D.Unit.AreaLoad") = "MPa"
               ElseIf Units.offset(1, 9) = 5 Then
               .Range("RISA3D.Unit.AreaLoad") = "kPa"
               ElseIf Units.offset(1, 9) = 6 Then
               .Range("RISA3D.Unit.AreaLoad") = "Pa"
               ElseIf Units.offset(1, 9) = 8 Then
               .Range("RISA3D.Unit.AreaLoad") = "kg/m^2"
               ElseIf Units.offset(1, 9) = 9 Then
               .Range("RISA3D.Unit.AreaLoad") = "kg/mm^2"
            End If

            .Range("RISA3D.Version") = WorksheetFunction.Substitute(RISA_Version.offset(1, 0), ";", "")
            End With

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
        ElseIf M_Header.offset(r, 10) = 2 Then              'Populates Members array if Cold Formed
            Members(n, 1) = RTrim(M_Header.offset(r))
            Members(n, 2) = RTrim(Application.WorksheetFunction.Index(Temp.Range("A" & CF_Header.row + 1, "A" & CF_Header.End(xlDown).row - 1), M_Header.offset(r, 7) + 1))
            n = n + 1
        ElseIf M_Header.offset(r, 10) = 6 Then              'Populates Members array if Aluminum
            Members(n, 1) = RTrim(M_Header.offset(r))
            Members(n, 2) = RTrim(Application.WorksheetFunction.Index(Temp.Range("A" & AL_Header.row + 1, "A" & AL_Header.End(xlDown).row - 1), M_Header.offset(r, 7) + 1))
            n = n + 1
        End If
    Next r
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    shGeometry.Range("B6", "C" & 4 + n) = Members                'Writes Members array
    
    M_cnt = (Range(M_Header, M_Header.End(xlDown)).Rows.Count - 2)
    ReDim Member_Section(1 To M_cnt, 1 To 4)
    For r = 1 To M_cnt
        If M_Header.offset(r) = 0 Or M_Header.offset(r) <> 0 Then
            Member_Section(r, 1) = RTrim(M_Header.offset(r))
            Member_Section(r, 2) = RTrim(M_Header.offset(r, 3))
            Member_Section(r, 3) = RTrim(M_Header.offset(r, 4))
            Member_Section(r, 4) = RTrim(M_Header.offset(r, 5))
            On Error Resume Next
        End If
    Next r
    shGeometry.Range("E6", "E" & 5 + (M_cnt)) = Member_Section
    shGeometry.Range("E6", "F" & 5 + (M_cnt)) = Member_Section
    shGeometry.Range("E6", "G" & 5 + (M_cnt)) = Member_Section
    shGeometry.Range("E6", "H" & 5 + (M_cnt)) = Member_Section
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Nodes_cnt = (Range(Nodes_Header, Nodes_Header.End(xlDown)).Rows.Count - 2)   'Writes Nodes array
    ReDim Nodes_Section(1 To Nodes_cnt, 1 To 4)
    For r = 1 To Nodes_cnt
        If Nodes_Header.offset(r) <> 0 Then
            Nodes_Section(r, 1) = RTrim(Nodes_Header.offset(r))
            Nodes_Section(r, 2) = RTrim(Nodes_Header.offset(r, 1))
            Nodes_Section(r, 3) = RTrim(Nodes_Header.offset(r, 2))
            Nodes_Section(r, 4) = RTrim(Nodes_Header.offset(r, 3))
            On Error Resume Next
        End If
    Next r
    shGeometry.Range("S6", "T" & 5 + (Nodes_cnt)).Value = Nodes_Section
    shGeometry.Range("S6", "U" & 5 + (Nodes_cnt)).Value = Nodes_Section
    shGeometry.Range("S6", "V" & 5 + (Nodes_cnt)).Value = Nodes_Section
        
'================================================================================================================================================================='
        
    Dim Projected_Width, Projected_Depth, Dc As Integer
    Dim TotalRows, NextRow As Integer
    
    TotalRows = shGeometry.Range("X" & Rows.Count).End(xlUp).row
    
    HR_cnt = (Range(HR_Header, HR_Header.End(xlDown)).Rows.Count - 2)

    ReDim HR_Section(1 To HR_cnt, 1 To 9)

    For r = 1 To HR_cnt
        If HR_Header.offset(r, 4) = 1 Then                                              'Populates Hot Rolled Section Sets array
            HR_Section(r, 1) = RTrim(HR_Header.offset(r))
            HR_Section(r, 2) = RTrim(HR_Header.offset(r, 2))
            HR_Section(r, 3) = RTrim(HR_Header.offset(r, 1))
            On Error Resume Next
        End If
    Next r
    shGeometry.Range("X6", "Y" & 5 + (HR_cnt)) = HR_Section          'Writes Hot Rolled Sections Sets array
    shGeometry.Range("X6", "Z" & 5 + (HR_cnt)) = HR_Section

'==========================================================================================================================================='
    
'==========================================================================================================================================='

    CF_cnt = (Range(CF_Header, CF_Header.End(xlDown)).Rows.Count - 2)
    
    ReDim CF_Section(1 To CF_cnt, 1 To 4)                                                'Populated Cold Formed Section Sets only if used in members array (since CF is more rarely used)
    
    For r = 1 To cnt
        If M_Header.offset(r, 10) = 2 Then
            t = t + 1
        End If
    Next r

    lastRow = shGeometry.Range("X5").End(xlDown).row
    If t > 1 Then
        For r = 1 To CF_cnt                                                              'Writes Cold Formed Sections Sets array
                CF_Section(r, 1) = RTrim(CF_Header.offset(r))
                CF_Section(r, 2) = RTrim(CF_Header.offset(r, 2))
                CF_Section(r, 3) = RTrim(CF_Header.offset(r, 1))
        Next r
        shGeometry.Range("X" & lastRow + 1, "X" & lastRow + CF_cnt) = CF_Section     'Writes Cold Formed Sections Sets array
        shGeometry.Range("X" & lastRow + 1, "Y" & lastRow + CF_cnt) = CF_Section     'Writes Cold Formed Sections Sets array
        shGeometry.Range("X" & lastRow + 1, "Z" & lastRow + CF_cnt) = CF_Section
    End If

'==========================================================================================================================================='
    
'==========================================================================================================================================='

    AL_cnt = (Range(AL_Header, AL_Header.End(xlDown)).Rows.Count - 2)
    
    ReDim AL_Section(1 To AL_cnt, 1 To 4)                                                'Populated Cold Formed Section Sets only if used in members array (since CF is more rarely used)
    
    For r = 1 To cnt
        If M_Header.offset(r, 10) = 6 Then
            t = t + 1
        End If
    Next r

    lastRow = shGeometry.Range("X5").End(xlDown).row
    If t > 1 Then
        For r = 1 To AL_cnt                                                              'Writes Cold Formed Sections Sets array
                AL_Section(r, 1) = RTrim(AL_Header.offset(r))
                AL_Section(r, 2) = RTrim(AL_Header.offset(r, 2))
                AL_Section(r, 3) = RTrim(AL_Header.offset(r, 1))
        Next r
        shGeometry.Range("X" & lastRow + 1, "X" & lastRow + AL_cnt) = AL_Section     'Writes Cold Formed Sections Sets array
        shGeometry.Range("X" & lastRow + 1, "Y" & lastRow + AL_cnt) = AL_Section     'Writes Cold Formed Sections Sets array
        shGeometry.Range("X" & lastRow + 1, "Z" & lastRow + AL_cnt) = AL_Section
    End If

'==========================================================================================================================================='
    

    shGeometry.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

Exit Sub
'Call TextToColumn

shGeometry.Range("B6").Select

End Sub


Sub TextToColumn()
    
    With shGeometry
    
    .Range("E6").Select
    .Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=.Range("E6"), dataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    .Range("F6").Select
    .Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=.Range("F6"), dataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    .Range("G6").Select
    .Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=.Range("G6"), dataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    .Range("H6").Select
    .Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=.Range("H6"), dataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    .Range("B6").Select
    
    End With

End Sub


Sub Member_Geometry()
    
    Dim a As Integer                                            'Counter to detmerine size of array
    Dim b As Integer                                            'Counter to detmerine size of array
    Dim c As Integer                                            'Counter to find array of members
    Dim d As Integer                                            'Counter to find array of members
    Dim lastRow As Long                                         'Last row of member labels table

    Dim i_node As Double                                        'Member I Node
    Dim j_node As Double                                        'Member J Node
    Dim i_lookup As Double                                      'I Node Number VLookup
    Dim j_lookup As Double                                      'J Node Number VLookup
             
    With shGeometry
        lastRow = .Range("B6").CurrentRegion.End(xlDown).row
    For c = 6 To lastRow                                        'Determine size of array by filtering out "RIGID" members
        If .Range("C" & c).Value <> "RIGID" Then
            a = a + 1
        End If
    Next c
             
    i_lookup = Application.WorksheetFunction.VLookup(.Range("B" & c), .Range("AE6:AG10003"), 2, False)
    i_node = Application.WorksheetFunction.VLookup(i_lookup, .Range("I6:J10003"), 2, False)
    
    j_lookup = Application.WorksheetFunction.VLookup(.Range("B" & c), .Range("AE6:AG10003"), 3, False)
    j_node = Application.WorksheetFunction.VLookup(i_lookup, .Range("I6:J10003"), 3, False)
    
    End With

End Sub


Sub RowCount()
Application.ScreenUpdating = False

    Dim i As Integer, j As Long
    Dim TotalRows As Integer, StartRow As Long
    
    With shGeometry
    TotalRows = .Range("S" & Rows.Count).End(xlUp).row
        
    For i = 6 To TotalRows
    
    If .Cells(i, 19) <> "" Then
        .Cells(i, 18).Value = WorksheetFunction.CountA(Range("S6:S" & i))
    End If
    Next i
    Exit Sub
    End With
    
Application.ScreenUpdating = True
End Sub


Sub Calc_Length()
    Dim ws As Worksheet
    Dim rngLookUp As Range
    Dim m As Integer, n As Integer, o As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim NodeRows As Integer, MemberRows As Integer
    Dim TotalRows_i As Integer, TotalRows_j As Integer
    Dim TotalRows_o As Integer
    Dim X_Rows As Integer
    Dim Answer As String
    
    With shGeometry
    
    TotalRows_i = .Range("B" & Rows.Count).End(xlUp).row
    X_Rows = .Range("V" & Rows.Count).End(xlUp).row
    For i = 6 To TotalRows_i
        .Cells(i, 13).Value = WorksheetFunction.VLookup(WorksheetFunction.VLookup(.Cells(i, 2).Value, .Range("E6:H99999"), 2, False), .Range("R6:S99999"), 2, False)
        .Cells(i, 14).Value = WorksheetFunction.VLookup(WorksheetFunction.VLookup(.Cells(i, 2).Value, .Range("E6:H99999"), 3, False), .Range("R6:S99999"), 2, False)
        
        If Range("RISA3D.Unit.Length") = "ft" Then
            .Cells(i, 15).Value = Format(WorksheetFunction.Power((WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 2, False) - WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 2, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 3, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 3, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 4, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 4, False)), 2)), 0.5) _
                                  * 12, "0###.###")
        
        ElseIf Range("RISA3D.Unit.Length") = "in" Then
            .Cells(i, 15).Value = Format(WorksheetFunction.Power((WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 2, False) - WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 2, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 3, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 3, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 4, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 4, False)), 2)), 0.5), "0###.###")
        
        ElseIf Range("RISA3D.Unit.Length") = "m" Then
            .Cells(i, 15).Value = Format(WorksheetFunction.Power((WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 2, False) - WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 2, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 3, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 3, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 4, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 4, False)), 2)), 0.5) _
                                  / 0.3048, "0###.###")
       
        ElseIf Range("RISA3D.Unit.Length") = "cm" Then
            .Cells(i, 15).Value = Format(WorksheetFunction.Power((WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 2, False) - WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 2, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 3, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 3, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 4, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 4, False)), 2)), 0.5) _
                                  / 30.48, "0###.###")
       
        ElseIf Range("RISA3D.Unit.Length") = "mm" Then
            .Cells(i, 15).Value = Format(WorksheetFunction.Power((WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 2, False) - WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 2, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 3, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 3, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 4, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 4, False)), 2)), 0.5) _
                                  / 304.8, "0###.###")
       
        End If
    Next i
    End With
End Sub


Sub Geometry_Lookup()
    
    Dim ws As Worksheet
    Dim rngLookUp As Range
    Dim m As Integer, n As Integer, o As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim NodeRows As Integer, MemberRows As Integer
    Dim TotalRows_i As Integer, TotalRows_j As Integer
    Dim TotalRows_o As Integer
    Dim X_Rows As Integer
    Dim Answer As String
    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With shGeometry
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  UNIT CONVERSIONS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   MEMBER LENGTHS  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    TotalRows_i = .Range("B" & Rows.Count).End(xlUp).row
    X_Rows = .Range("V" & Rows.Count).End(xlUp).row
    For i = 6 To TotalRows_i
        .Cells(i, 13).Value = WorksheetFunction.VLookup(WorksheetFunction.VLookup(.Cells(i, 2).Value, .Range("E6:H99999"), 2, False), .Range("R6:S99999"), 2, False)
        .Cells(i, 14).Value = WorksheetFunction.VLookup(WorksheetFunction.VLookup(.Cells(i, 2).Value, .Range("E6:H99999"), 3, False), .Range("R6:S99999"), 2, False)
        
        If Range("RISA3D.Unit.Length") = "ft" Then
            .Cells(i, 15).Value = Format(WorksheetFunction.Power((WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 2, False) - WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 2, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 3, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 3, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 4, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 4, False)), 2)), 0.5) _
                                  * 12, "0###.###")
        
        ElseIf Range("RISA3D.Unit.Length") = "in" Then
            .Cells(i, 15).Value = Format(WorksheetFunction.Power((WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 2, False) - WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 2, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 3, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 3, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 4, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 4, False)), 2)), 0.5), "0###.###")
        
        ElseIf Range("RISA3D.Unit.Length") = "m" Then
            .Cells(i, 15).Value = Format(WorksheetFunction.Power((WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 2, False) - WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 2, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 3, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 3, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 4, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 4, False)), 2)), 0.5) _
                                  / 0.3048, "0###.###")
       
        ElseIf Range("RISA3D.Unit.Length") = "cm" Then
            .Cells(i, 15).Value = Format(WorksheetFunction.Power((WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 2, False) - WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 2, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 3, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 3, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 4, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 4, False)), 2)), 0.5) _
                                  / 30.48, "0###.###")
       
        ElseIf Range("RISA3D.Unit.Length") = "mm" Then
            .Cells(i, 15).Value = Format(WorksheetFunction.Power((WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 2, False) - WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 2, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 3, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 3, False)), 2) + _
                                  WorksheetFunction.Power((WorksheetFunction.VLookup(.Cells(i, 14), .Range("S6:V" & X_Rows), 4, False) - WorksheetFunction.VLookup(.Cells(i, 13), .Range("S6:V" & X_Rows), 4, False)), 2)), 0.5) _
                                  / 304.8, "0###.###")
       
        End If
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' MEMBER OREINATIONS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If WorksheetFunction.VLookup(.Cells(i, 2), .Range("E6:K99999"), 5, False) = 0 And WorksheetFunction.VLookup(.Cells(i, 2), .Range("E6:K99999"), 7, False) = 0 Then
                .Cells(i, 16).Value = 0
            ElseIf WorksheetFunction.degrees(WorksheetFunction.Atan2(1, WorksheetFunction.VLookup(.Cells(i, 2), .Range("E6:K99999"), 5, False))) = 0 Then
            .Cells(i, 16).Value = 90
            ElseIf WorksheetFunction.degrees(WorksheetFunction.Atan2(1, (WorksheetFunction.VLookup(.Cells(i, 2), .Range("E6:K99999"), 7, False) / WorksheetFunction.VLookup(.Cells(i, 2), .Range("E6:K99999"), 5, False)))) < 0 Then
                .Cells(i, 16).Value = Format(180 + WorksheetFunction.degrees(WorksheetFunction.Atan2(1, (WorksheetFunction.VLookup(.Cells(i, 2), .Range("E6:K99999"), 7, False) / WorksheetFunction.VLookup(.Cells(i, 2), .Range("E6:K99999"), 5, False)))), "0###")
            Else
                .Cells(i, 16).Value = Format(WorksheetFunction.degrees(WorksheetFunction.Atan2(1, (WorksheetFunction.VLookup(.Cells(i, 2), .Range("E6:K99999"), 7, False) / WorksheetFunction.VLookup(.Cells(i, 2), .Range("E6:K99999"), 5, False)))), "0###")
        End If
    
    Next i
   
    
    TotalRows_j = .Range("E" & Rows.Count).End(xlUp).row
    X_Rows = .Range("X" & Rows.Count).End(xlUp).row
    For j = 6 To TotalRows_j
        .Cells(j, 9).Value = Format(WorksheetFunction.VLookup(WorksheetFunction.VLookup(.Cells(j, 6), .Range("R6:S99999"), 2, False), .Range("S6:V99999"), 2, False) - _
                             WorksheetFunction.VLookup(WorksheetFunction.VLookup(.Cells(j, 7), .Range("R6:S99999"), 2, False), .Range("S6:V99999"), 2, False), "0###.###")
    
        .Cells(j, 10).Value = Format(WorksheetFunction.VLookup(WorksheetFunction.VLookup(.Cells(j, 7), .Range("R6:S99999"), 2, False), .Range("S6:V99999"), 3, False) - _
                             WorksheetFunction.VLookup(WorksheetFunction.VLookup(.Cells(j, 6), .Range("R6:S99999"), 2, False), .Range("S6:V99999"), 3, False), "0###.###")
    
        .Cells(j, 11).Value = Format(WorksheetFunction.VLookup(WorksheetFunction.VLookup(.Cells(j, 7), .Range("R6:S99999"), 2, False), .Range("S6:V99999"), 4, False) - _
                             WorksheetFunction.VLookup(WorksheetFunction.VLookup(.Cells(j, 6), .Range("R6:S99999"), 2, False), .Range("S6:V99999"), 4, False), "0###.###")
    
    Next j
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' MEMBER PROPERITIES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    TotalRows_o = .Range("X" & Rows.Count).End(xlUp).row
    For m = 6 To TotalRows_o
    On Error Resume Next
    'On Error GoTo Handle
'    Dim Answer As Integer

        If .Cells(m, 26).Value = "BAR" Or .Cells(m, 26).Value = "HSS Pipe A1085" Or .Cells(m, 26).Value = "HSS Pipe" Or .Cells(m, 26).Value = "Pipe" Then
            .Cells(m, 27).Value = "Round"
        ElseIf .Cells(m, 26).Value = "None" Then
            .Cells(m, 27).Value = "Please input member shape as Flat or Round"
        ElseIf .Cells(m, 26).Value = "SquareTube" Or .Cells(m, 26).Value = "SquareTube A1085" Then
            .Cells(m, 27).Value = "HSS Flat"
'        ElseIf .Cells(m, 26).Value = "#N/A" Then
'            .Cells(m, 27).Value = "Flat"
        ElseIf .Cells(m, 26) = "" Then
        Else
            .Cells(m, 27).Value = "Flat"
        End If
'
'
'        If .Cells(m, 27).Value = "HSS Flat" And .Cells(m, 27).Value <> "" Then
''            .Cells(m, 30).Value = "User Input"
'            .Cells(m, 31).Value = "=IFERROR(1.5 * AD" & m & " / AB" & m & ", " & Chr(34) & "User Input" & Chr(34) & ")"
'        Else
''            .Cells(m, 30).Value = "N/A"
'            .Cells(m, 31).Value = "N/A"
'        End If
'
'
'        If .Cells(m, 27).Value = "Round" Or .Cells(m, 26).Value = "BAR" And .Cells(m, 27).Value <> "" Then
'            .Cells(m, 32).Value = "=AB" & m
'        Else
'            .Cells(m, 32).Value = "=IFERROR(SQRT(AB" & m & "^2 + AC" & m & "^2, " & Chr(34) & "User Input" & Chr(34) & ")"
'            '.Cells(m, 32).Value = (.Cells(m, 28) ^ 2 + .Cells(m, 29) ^ 2) ^ 0.5
'        End If
        
'Handle:
'Answer = MsgBox("Does the RISA-3D model contain unistrut members?", vbYesNo, "Possible Cold Formed Member(s) Found.")
'If Answer = vbYes Then
''    .Cells(m, 26).Value = "CU"
''    .Cells(m, 27).Value = "Flat"
''    .Cells(m, 28).Value = 1.6
''    .Cells(m, 29).Value = 1.6
'ElseIf Answer = vbNo Then
'On Error Resume Next
'End If
    Next m
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End With
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub


Public Function GetNumber(s As String) As Long
    Dim b As Boolean, i As Long, t As String
    b = False
    t = ""
    For i = 1 To Len(s)
        If IsNumeric(Mid(s, i, 1)) Then
            b = True
            t = t & Mid(s, i, 1)
        Else
            If b Then
                GetNumber = CLng(t)
                Exit Function
            End If
        End If
    Next i
End Function


Sub ModelEPA()
Dim r, c As Long
Dim TotalRows As Integer
Dim lastRow As Long

With shGeometry
TotalRows = .Range("B" & Rows.Count).End(xlUp).row
lastRow = .Range("R" & Rows.Count).End(xlUp).row
.Range("AR6:BG" & lastRow).ClearContents

For r = 6 To TotalRows
    ' 0°
    .Range("AS" & r) = "=+IFERROR(IF(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0)=0,$O" & r & "/12*COS(PI()/180*($P" & r & "+0))^2,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))>0,$P" & r & "=0),$O" & r & "/12*COS(PI()/180*($P" & r & "+0))^2,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))>0,ABS($P" & r & ")>0),$O" & r & "/12*COS(PI()/180*($P" & r & "+0))^2))),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
    .Range("AT" & r) = "=+IFERROR(IF(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0)=0,0,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))>0,$P" & r & "=0),ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))*SIN(PI()/180*($P" & r & "+0))^2,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))>0,ABS($P" & r & ")>0),ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))*SIN(PI()/180*($P" & r & "+0))^2))),0)*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
    .Range("AU" & r) = "=+IFERROR((AS" & r & "+AT" & r & ")*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)*IF(VLOOKUP($C" & r & ",rngR3DSectionSets,4,0)=" & Chr(34) & "Flat" & Chr(34) & ",2,1.2),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)
    .Range("AV" & r) = "=+IFERROR(+IF(VLOOKUP($M" & r & ",$S$6:$V" & lastRow & ",3,0)=VLOOKUP($N" & r & ",$S$6:$V" & lastRow & ",3,0), ((((($O" & r & "/12+(2*0.5/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*0.5/12)))-($O" & r & "/12*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))*1.2*COS(PI()/180*($P" & r & "))^2)) + AU" & r & ", (((((AS" & r & "+(2*0.5/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*0.5/12)))-(AS" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))+(((AT" & r & "+(2*0.5/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*0.5/12)))-(AT" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12))))*1.2) + AU" & r & "),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
    .Range("AW" & r) = "=+IFERROR(+IF(VLOOKUP($M" & r & ",$S$6:$V" & lastRow & ",3,0)=VLOOKUP($N" & r & ",$S$6:$V" & lastRow & ",3,0), ((((($O" & r & "/12+(2*1.0/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*1.0/12)))-($O" & r & "/12*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))*1.2*COS(PI()/180*($P" & r & "))^2)) + AU" & r & ", (((((AS" & r & "+(2*1.0/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*1.0/12)))-(AS" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))+(((AT" & r & "+(2*1.0/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*1.0/12)))-(AT" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12))))*1.2) + AU" & r & "),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
    
    ' 90°
    .Range("AX" & r) = "=+IFERROR(IF(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0)=0,$O" & r & "/12*COS(PI()/180*($P" & r & "+90))^2,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))>0,$P" & r & "=0),$O" & r & "/12*COS(PI()/180*($P" & r & "+90))^2,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))>0,ABS($P" & r & ")>0),$O" & r & "/12*COS(PI()/180*($P" & r & "+90))^2))),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
    .Range("AY" & r) = "=+IFERROR(IF(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0)=0,0,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))>0,$P" & r & "=0),ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))*SIN(PI()/180*($P" & r & "+90))^2,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))>0,ABS($P" & r & ")>0),ABS(VLOOKUP($B" & r & ",$E$6:$K" & lastRow & ",6,0))*SIN(PI()/180*($P" & r & "+90))^2))),0)*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
    .Range("AZ" & r) = "=+IFERROR((AX" & r & "+AY" & r & ")*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)*IF(VLOOKUP($C" & r & ",rngR3DSectionSets,4,0)=" & Chr(34) & "Flat" & Chr(34) & ",2,1.2),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
    .Range("BA" & r) = "=+IFERROR(+IF(VLOOKUP($M" & r & ",$S$6:$V" & lastRow & ",3,0)=VLOOKUP($N" & r & ",$S$6:$V" & lastRow & ",3,0), ((((($O" & r & "/12+(2*0.5/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*0.5/12)))-($O" & r & "/12*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))*1.2*COS(PI()/180*($P" & r & "))^2)) + AZ" & r & ", (((((AX" & r & "+(2*0.5/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*0.5/12)))-(AX" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))+(((AY" & r & "+(2*0.5/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*0.5/12)))-(AY" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12))))*1.2) + AZ" & r & "),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
    .Range("BB" & r) = "=+IFERROR(+IF(VLOOKUP($M" & r & ",$S$6:$V" & lastRow & ",3,0)=VLOOKUP($N" & r & ",$S$6:$V" & lastRow & ",3,0), ((((($O" & r & "/12+(2*1.0/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*1.0/12)))-($O" & r & "/12*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))*1.2*COS(PI()/180*($P" & r & "))^2)) + AZ" & r & ", (((((AX" & r & "+(2*1.0/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*1.0/12)))-(AX" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))+(((AY" & r & "+(2*1.0/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*1.0/12)))-(AY" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12))))*1.2) + AZ" & r & "),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
    
'    ' 120°
'    .Range("BC" & r) = "=+IFERROR(IF(VLOOKUP($B" & r & ",$E$6:$K" & LastRow & ",6,0)=0,$O" & r & "/12*COS(PI()/180*($P" & r & "+120))^2,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & LastRow & ",6,0))>0,$P" & r & "=0),$O" & r & "/12*COS(PI()/180*($P" & r & "+120))^2,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & LastRow & ",6,0))>0,ABS($P" & r & ")>0),$O" & r & "/12*COS(PI()/180*($P" & r & "+120))^2))),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
'    .Range("BD" & r) = "=+IFERROR(IF(VLOOKUP($B" & r & ",$E$6:$K" & LastRow & ",6,0)=0,0,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & LastRow & ",6,0))>0,$P" & r & "=0),ABS(VLOOKUP($B" & r & ",$E$6:$K" & LastRow & ",6,0))*SIN(PI()/180*($P" & r & "+120))^2,IF(AND(ABS(VLOOKUP($B" & r & ",$E$6:$K" & LastRow & ",6,0))>0,ABS($P" & r & ")>0),ABS(VLOOKUP($B" & r & ",$E$6:$K" & LastRow & ",6,0))*SIN(PI()/180*($P" & r & "+120))^2))),0)*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
'    .Range("BE" & r) = "=+IFERROR((BB" & r & "+BC" & r & ")*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)*IF(VLOOKUP($C" & r & ",rngR3DSectionSets,4,0)=" & Chr(34) & "Flat" & Chr(34) & ",2,1.2),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)
'    .Range("BF" & r) = "=+IFERROR(+IF(VLOOKUP($M" & r & ",$S$6:$V" & LastRow & ",3,0)=VLOOKUP($N" & r & ",$S$6:$V" & LastRow & ",3,0), ((((($O" & r & "/12+(2*0.5/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*0.5/12)))-($O" & r & "/12*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))*1.2*COS(PI()/180*($P" & r & "))^2)) + AT" & r & ", (((((BB" & r & "+(2*0.5/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*0.5/12)))-(BB" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))+(((BC" & r & "+(2*0.5/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*0.5/12)))-(BC" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12))))*1.2) + AT" & r & "),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
'    .Range("BG" & r) = "=+IFERROR(+IF(VLOOKUP($M" & r & ",$S$6:$V" & LastRow & ",3,0)=VLOOKUP($N" & r & ",$S$6:$V" & LastRow & ",3,0), ((((($O" & r & "/12+(2*1.0/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*1.0/12)))-($O" & r & "/12*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))*1.2*COS(PI()/180*($P" & r & "))^2)) + AT" & r & ", (((((BB" & r & "+(2*1.0/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*1.0/12)))-(BB" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)))+(((BC" & r & "+(2*1.0/12))*((VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12)+(2*1.0/12)))-(BC" & r & "*(VLOOKUP($C" & r & ",rngR3DSectionSets,5,0)/12))))*1.2) + AT" & r & "),0)" '*IF(RISA3D.Unit.Length=" & Chr(34) & "in" & Chr(34) & ",1/12,1)"
    
    ' Weight
    .Range("BC" & r) = "=LET(" & Chr(10) & _
                        "ti, 0.0," & Chr(10) & _
                        "h,  XLOOKUP($C" & r & ",$X:$X,$AB:$AB,0,0,1)," & Chr(10) & _
                        "d,  XLOOKUP($C" & r & ",$X:$X,$AC:$AC,0,0,1)," & Chr(10) & _
                        "l,  $O" & r & " / 12," & Chr(10) & _
                        "Ag, XLOOKUP($C" & r & ",$X:$X,$AG:$AG,0,0,1)," & Chr(10) & _
                        "UW, 490," & Chr(10) & _
                        "Wt, UW*Ag/144*l," & Chr(10) & _
                        "Wi, IFERROR((((h+2*ti)*(d+2*ti)*(l*12+2*ti))-(h*d*l*12))*56/1728,0)," & Chr(10) & _
                        "(Wi+Wt)*IF($B" & r & "=" & Chr(34) & Chr(34) & ",0,1))"
    .Range("BD" & r) = "=LET(" & Chr(10) & _
                        "ti, 0.5," & Chr(10) & _
                        "h,  XLOOKUP($C" & r & ",$X:$X,$AB:$AB,0,0,1)," & Chr(10) & _
                        "d,  XLOOKUP($C" & r & ",$X:$X,$AC:$AC,0,0,1)," & Chr(10) & _
                        "l,  $O" & r & " / 12," & Chr(10) & _
                        "Ag, XLOOKUP($C" & r & ",$X:$X,$AG:$AG,0,0,1)," & Chr(10) & _
                        "UW, 490," & Chr(10) & _
                        "Wt, UW*Ag/144*l," & Chr(10) & _
                        "Wi, IFERROR((((h+2*ti)*(d+2*ti)*(l*12+2*ti))-(h*d*l*12))*56/1728,0)," & Chr(10) & _
                        "(Wi+Wt)*IF($B" & r & "=" & Chr(34) & Chr(34) & ",0,1))"
    .Range("BE" & r) = "=LET(" & Chr(10) & _
                        "ti, 1.0," & Chr(10) & _
                        "h,  XLOOKUP($C" & r & ",$X:$X,$AB:$AB,0,0,1)," & Chr(10) & _
                        "d,  XLOOKUP($C" & r & ",$X:$X,$AC:$AC,0,0,1)," & Chr(10) & _
                        "l,  $O" & r & " / 12," & Chr(10) & _
                        "Ag, XLOOKUP($C" & r & ",$X:$X,$AG:$AG,0,0,1)," & Chr(10) & _
                        "UW, 490," & Chr(10) & _
                        "Wt, UW*Ag/144*l," & Chr(10) & _
                        "Wi, IFERROR((((h+2*ti)*(d+2*ti)*(l*12+2*ti))-(h*d*l*12))*56/1728,0)," & Chr(10) & _
                        "(Wi+Wt)*IF($B" & r & "=" & Chr(34) & Chr(34) & ",0,1))"
    
    .Range("AS" & r).NumberFormat = "0.00"
    .Range("AT" & r).NumberFormat = "0.00"
    .Range("AU" & r).NumberFormat = "0.00"
    .Range("AV" & r).NumberFormat = "0.00"
    .Range("AW" & r).NumberFormat = "0.00"
    
    .Range("AX" & r).NumberFormat = "0.00"
    .Range("AY" & r).NumberFormat = "0.00"
    .Range("AZ" & r).NumberFormat = "0.00"
    .Range("BA" & r).NumberFormat = "0.00"
    .Range("BB" & r).NumberFormat = "0.00"
    
    .Range("BC" & r).NumberFormat = "0.00"
    .Range("BD" & r).NumberFormat = "0.00"
    .Range("BE" & r).NumberFormat = "0.00"
    .Range("BF" & r).NumberFormat = "0.00"
    .Range("BG" & r).NumberFormat = "0.00"
Next r
End With
End Sub


