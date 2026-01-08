Attribute VB_Name = "RISAModel"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String
Private regex As New RegExp


Private FileStartData As String

Private pSectionSetData As String
Private pSectionSets() As SectionSetGroup

Private FileMiddleData As String

Private pNodesData As String
Private pNodes() As Node
Private pNodesByName As New Scripting.Dictionary

Private BoundaryConditions As String

Private pShapesData As String
Private pShapes() As Shape
Private pShapesByName As New Scripting.Dictionary

Private CustomReportSections As String

Private MemberData As String
Private pMembers() As Member
Private pMembersByName As New Scripting.Dictionary

Private pPlates() As Plate

'Private pBLCs() As BasicLoadCase
'
'Private pPointLoads() As PointLoad
'
'Private pDistributedLoads() As DistributedLoad
'
'Private pAreaLoads() As AreaLoad
'
'Private pSurfaceLoads() As SurfaceLoad

Private FileMidEndData As String

'Private pLCs() As LoadCombination

Private FileEndData As String

Private OutputStr


Private MinX
Private MinY
Private MinZ

Private MaxX
Private MaxY
Private MaxZ

Public Property Get ModelMinX()
    ModelMinX = MinX
End Property
Public Property Get ModelMinY()
    ModelMinY = MinY
End Property
Public Property Get ModelMinZ()
    ModelMinZ = MinZ
End Property

Public Property Get ModelMaxX()
    ModelMaxX = MaxX
End Property
Public Property Get ModelMaxY()
    ModelMaxY = MaxY
End Property
Public Property Get ModelMaxZ()
    ModelMaxZ = MaxZ
End Property

'when we get to area loads, make it an option to keep the existing area loads, but don't automate them otherwise
'have a popup reminding the user to enter the area loads
'the popup will list all the options for grating weight and show the ice weight (if there is ice)

Private Sub Class_Initialize()
    Globals.LoadCatMap.Add "DL", 1
    Globals.LoadCatMap.Add "LL", 2
    Globals.LoadCatMap.Add "WLZ", 31
    Globals.LoadCatMap.Add "WLY", 30
    Globals.LoadCatMap.Add "WLX", 29
    Globals.LoadCatMap.Add "OL1", 16
    Globals.LoadCatMap.Add "OL2", 17
    Globals.LoadCatMap.Add "OL3", 18
    Globals.LoadCatMap.Add "ELZ", 28
    Globals.LoadCatMap.Add "ELY", 27
    Globals.LoadCatMap.Add "ELX", 26
    Globals.LoadCatMap.Add "None", 0
    
    
    Globals.DirMap.Add "x", 120
    Globals.DirMap.Add "y", 121
    Globals.DirMap.Add "z", 122
    Globals.DirMap.Add "X", 88
    Globals.DirMap.Add "Y", 89
    Globals.DirMap.Add "Z", 90
    Globals.DirMap.Add "SX", 10
    Globals.DirMap.Add "SY", 11
    Globals.DirMap.Add "SZ", 12
    Globals.DirMap.Add "PX", 76
    Globals.DirMap.Add "PY", 86
    Globals.DirMap.Add "PZ", 72
    
    
    Globals.MaterialTypeMap.Add 0, "GENERAL"
    Globals.MaterialTypeMap.Add 1, "HR_STEEL"
    Globals.MaterialTypeMap.Add 2, "CF_STEEL"
    Globals.MaterialTypeMap.Add 3, "WOOD"
    Globals.MaterialTypeMap.Add 5, "CONCRETE"
    Globals.MaterialTypeMap.Add 6, "ALUMINUM"
    
    With regex
        .Global = False
        .IgnoreCase = False
        .MultiLine = True
    End With

End Sub
'
'Public Function PointLoads()
'    PointLoads = pPointLoads
'End Function
'
'Public Function DistributedLoads()
'    DistributedLoads = pDistributedLoads
'End Function
'
'Public Function AreaLoads()
'    AreaLoads = pAreaLoads
'End Function
'
'Public Function SurfaceLoads()
'    SurfaceLoads = pSurfaceLoads
'End Function
'
'Public Function BLCs()
'    BLCs = pBLCs
'End Function
'
'Public Function LCs()
'    LCs = pLCs
'End Function


Public Function SectionSets()
    SectionSets = pSectionSets
End Function

Public Function Shapes()
    Shapes = pShapes
End Function

Public Function Nodes()
    Nodes = pNodes
End Function

Public Function Members()
    Members = pMembers
End Function

Public Function Plates()
    Plates = pPlates
End Function

Public Function MembersByName()
    Set MembersByName = pMembersByName
End Function

Public Function NodesByName()
    Set NodesByName = pNodesByName
End Function
'
'Public Sub SetBLCs(BLCs)
'    pBLCs = BLCs
'End Sub
'
'Public Sub SetLCs(LCs)
'    pLCs = LCs
'End Sub
'
'Public Sub SetPointLoads(PointLoads)
'    pPointLoads = PointLoads
'End Sub
'
'Public Sub SetDistributedLoads(DistributedLoads)
'    pDistributedLoads = DistributedLoads
'End Sub






Public Sub Parse(SourcePath)
    On Error Resume Next
        Open SourcePath For Input As #1
            Raw = Input(LOF(1), 1)
        Close #1
        
        If Err.Number <> 0 Then
            MsgBox ("The following path could not be found:" & vbNewLine & SourcePath & vbNewLine & "If this path contains a web address, please re-link the model by navigating to it from your C: drive (Starting at Desktop, Documents, etc.)")
            End
        End If
    On Error GoTo 0
    
    'start processing here
    ProcessFileStartData
    ProcessSectionSets
    ProcessFileMiddleData
    ProcessNodes
    ProcessBoundaryConditions
    ProcessShapes
    ProcessCustomReportSections
    ProcessMembers
    ProcessPlates
'    ProcessBLCs
'    ProcessPointLoads
'    ProcessDistributedLoads
'    ProcessAreaLoads
'    ProcessSurfaceLoads
    ProcessFileMidEndData
'    ProcessLCs
    ProcessFileEndData
    
End Sub



'Start All Processing Functions
Private Sub ProcessFileStartData()
    'set the pattern to match the first portion of the file that we don't edit or need data from
    regex.Global = False
    regex.pattern = "(.|\n)+\[END_MATERIAL_PROPERTIES\]"
    
    If regex.Test(Raw) Then
        FileStartData = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If
End Sub

Private Sub ProcessSectionSets()
    regex.Global = False
    regex.pattern = "\[SECTION_SETS\](.|\n)+\[END_SECTION_SETS\]"
    
    If regex.Test(Raw) Then
        pSectionSetData = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If

    'set the pattern to match the Section Set Data
    regex.Global = True
    'regEx.Pattern = "(\[\..{1,9}_SECTION_SETS\] <[0-9]+>)"
    regex.pattern = "(\[\..{1,9}_SECTION_SETS\] <[0-9]+>\r\n)((.+;\r\n)+)(\[\.END_.{1,9}_SECTION_SETS\])"
    'regEx.Pattern = "[\[\.][.]{1,9}[_SECTION_SETS\]](.|\n)+[\[\.END_][.]{1,9}[_SECTION_SETS\]]"
    
    If regex.Test(Raw) Then
        Set Data = regex.Execute(Raw)
    Else
        Exit Sub
    End If

    ReDim pSectionSets(1 To Data.Count)
    
    For i = 1 To Data.Count
        Set pSectionSets(i) = New SectionSetGroup
        pSectionSets(i).Define Data(i - 1).Value
    Next i
    
End Sub

Private Sub ProcessFileMiddleData()
    'Set the pattern to get the data between the section sets and nodes that we don't edit or need data from
    regex.Global = False
    regex.pattern = "\[WOOD_SCHEDULES\](.|\n)+\[END_CONNECTION_RULES\]"
    
    If regex.Test(Raw) Then
        FileMiddleData = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If
    
End Sub

Private Sub ProcessNodes()
    regex.Global = False
    regex.pattern = "\[NODES\](.|\n)+\[END_NODES\]"
    
    If regex.Test(Raw) Then
        pNodesData = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If
    
    ProcessList "NODES", "Node", pNodes
    MinX = pNodes(LBound(pNodes)).X
    MinY = pNodes(LBound(pNodes)).Y
    MinZ = pNodes(LBound(pNodes)).Z
    
    MaxX = pNodes(LBound(pNodes)).X
    MaxY = pNodes(LBound(pNodes)).Y
    MaxZ = pNodes(LBound(pNodes)).Z
    For i = LBound(pNodes) To UBound(pNodes)
        pNodes(i).Number = i
        NodesByName.Add pNodes(i).Name, pNodes(i)
        
        If pNodes(i).X < MinX Then
            MinX = pNodes(i).X
        ElseIf pNodes(i).X > MaxX Then
            MaxX = pNodes(i).X
        End If
        
        If pNodes(i).Y < MinY Then
            MinY = pNodes(i).Y
        ElseIf pNodes(i).Y > MaxY Then
            MaxY = pNodes(i).Y
        End If
        
        If pNodes(i).Z < MinZ Then
            MinZ = pNodes(i).Z
        ElseIf pNodes(i).Z > MaxZ Then
            MaxZ = pNodes(i).Z
        End If
    Next i

End Sub



Private Sub ProcessBoundaryConditions()
    'Set the pattern to get the data between the section sets and nodes that we don't edit or need data from
    regex.Global = False
    regex.pattern = "\[BOUNDARY_CONDITIONS\](.|\n)+\[END_BOUNDARY_CONDITIONS\]"
    
    If regex.Test(Raw) Then
        BoundaryConditions = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If
    
End Sub

Private Sub ProcessShapes()
    regex.Global = False
    regex.pattern = "\[SHAPES_LIST\](.|\n)+\[END_SHAPES_LIST\]"
    
    If regex.Test(Raw) Then
        pShapesData = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If


    ProcessList "SHAPES_LIST", "Shape", pShapes
    For i = LBound(pShapes) To UBound(pShapes)
        pShapesByName.Add pShapes(i).Name, pShapes(i)
    Next i
    
    ' Rigid shapes aren't saved in risa model files but we need to add it to the dictionary here for use later
    ' when we get member diameters
    Dim RigidShape As Shape
    Set RigidShape = New Shape
    
    pShapesByName.Add RigidShape.Name, RigidShape
End Sub

Private Sub ProcessCustomReportSections()
    'Set the pattern to get the data between the section sets and nodes that we don't edit or need data from
    regex.Global = False
    regex.pattern = "\[CUSTOM_REPORT_SECTIONS_3D\](.|\n)+\[END_CUSTOM_REPORT_SECTIONS_3D\]"
    
    If regex.Test(Raw) Then
        CustomReportSections = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If
End Sub

Private Sub ProcessMembers()
    'Gather member data including all subsections
    regex.Global = False
    regex.pattern = "\[MEMBERS\](.|\n)+\[END_MEMBERS\]"
    
    If regex.Test(Raw) Then
        MemberData = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If
    
    
    'Gather member main data for use in calculations
    regex.pattern = "\[.MEMBERS_MAIN_DATA\](.|\n)+\[.END_MEMBERS_MAIN_DATA\]"
    
    If regex.Test(Raw) Then
        RetStr = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If
    
    'Pattern to get each row of node data
    regex.Global = True
    regex.pattern = ".+;$"
    
    If regex.Test(RetStr) Then
        Set Data = regex.Execute(RetStr)
    Else
        Exit Sub
    End If
    
    ReDim pMembers(1 To Data.Count)
    For i = 1 To Data.Count
        Set pMembers(i) = New Member
        pMembers(i).Define Data(i - 1)
        pMembers(i).Number = i
        MembersByName.Add pMembers(i).Name, pMembers(i)
        CalcMemberDiameter pMembers(i)
        'CalcMemberLength pMembers(i)
    Next i
End Sub

Private Sub ProcessPlates()
    ProcessList "PLATES", "Plate", pPlates
End Sub
'
'Private Sub ProcessBLCs()
'    ProcessList "BASIC_LOAD_CASES", "BasicLoadCase", pBLCs
'End Sub
'
'Private Sub ProcessPointLoads()
'    ProcessList "POINT_LOADS", "PointLoad", pPointLoads
'End Sub
'
'
'Private Sub ProcessDistributedLoads()
'    ProcessList "DISTRIBUTED_LOADS", "DistributedLoad", pDistributedLoads
'End Sub
'
'Private Sub ProcessAreaLoads()
'    ProcessList "AREA_LOADS", "AreaLoad", pAreaLoads
'End Sub
'
'Private Sub ProcessSurfaceLoads()
'    ProcessList "SURFACE_LOADS", "SurfaceLoad", pSurfaceLoads
'End Sub

Private Sub ProcessFileMidEndData()
    regex.Global = False
    regex.pattern = "\[TIME_HISTORY_INPUT\](.|\n)+\[END_SPECTRA_SCALING_FACTOR\]"
    
    If regex.Test(Raw) Then
        FileMidEndData = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If
End Sub
'
'Private Sub ProcessLCs()
'    ProcessList "LOAD_COMBINATIONS", "LoadCombination", pLCs
'End Sub

Private Sub ProcessFileEndData()
    regex.Global = False
    regex.pattern = "\[FILE_INFORMATION\](.|\n)+\[FINISH\]"
    
    If regex.Test(Raw) Then
        FileEndData = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If
End Sub


Private Sub ProcessList(Label, dataType As String, ByRef DataList)
    'Pattern to get the String that is the list of nodes
    regex.Global = False
    regex.pattern = "\[" & Label & "\].+\r\n((.+;\r\n)+)\[END_" & Label & "\]"
    
    If regex.Test(Raw) Then
        RetStr = regex.Execute(Raw)(0)
    Else
        Exit Sub
    End If
    
    'Pattern to get each row of node data
    regex.Global = True
    regex.pattern = ".+;$"
    'regEx.Pattern = """.{32}""|[^ \t\r\n\f;]+"
    
    If regex.Test(RetStr) Then
        Set Data = regex.Execute(RetStr)
    Else
        Exit Sub
    End If
    
    ReDim DataList(1 To Data.Count)
    For i = 1 To Data.Count
        Set DataList(i) = NewObject(dataType)
        DataList(i).Define Data(i - 1).Value
    Next i
    
End Sub




Private Function NewObject(ClassName As String) As Object
    Set NewObject = Application.Run("Globals.c" & ClassName)
End Function


Private Sub CalcMemberDiameter(ByRef Mem As Member)

    Dim MemberShape As Shape
    Set MemberShape = pShapesByName(Mem.ShapeName)

    Mem.IsFlat = True
    Select Case MemberShape.ShapeType
        Case 1, 2, 4, 7, 8, 200, 201, 450, 457, 501
            ' square tube, angle, channel, wide flange
            ' Dimension(2) is height Dimension(4) is width
            Mem.Dc = (MemberShape.Dimensions(2) ^ 2 + MemberShape.Dimensions(4) ^ 2) ^ 0.5
        Case 3, 9, 460
            ' round tubes, SR Shapes
            ' Dimension(2) is Diameter
            Mem.Dc = MemberShape.Dimensions(2)
            Mem.IsFlat = False
        Case 5
            ' plates
            Mem.Dc = Application.WorksheetFunction.Max(MemberShape.Dimensions(2), MemberShape.Dimensions(4))
        Case 6
            ' double angles
            Mem.Dc = Application.WorksheetFunction.Max(MemberShape.Dimensions(2) * 2 + MemberShape.Dimensions(6), MemberShape.Dimensions(4))
        Case 0
            'Rigid (not official, i set this up as a default)
            Mem.Dc = 0
        Case Else
            ' unknown
            Mem.Dc = InputBox("Please enter the member width for the following shape: " & Trim(Replace(MemberShape.Name, """", "")), "Enter Member Width")
    End Select

End Sub

Private Sub CalcMemberLength(ByRef Mem As Member)
    Dim oINode As Node
    Dim oJNode As Node
    
    Set oINode = Nodes(Mem.iNode)
    Set oJNode = Nodes(Mem.jNode)
    
    Mem.Length = ((oINode.X - oJNode.X) ^ 2 + (oINode.Y - oJNode.Y) ^ 2 + (oINode.Z - oJNode.Z) ^ 2) ^ 0.5
End Sub

Public Function StrRep()
    StrRep = FileStartData
    
    If UBound(SectionSets) - LBound(SectionSets) + 1 > 0 Then
        StrRep = StrRep & vbNewLine & vbNewLine & pSectionSetData
    End If
    
    StrRep = StrRep & vbNewLine & vbNewLine & FileMiddleData
    
    If UBound(Nodes) - LBound(Nodes) + 1 > 0 Then
        StrRep = StrRep & vbNewLine & vbNewLine & pNodesData
    End If
    
    StrRep = StrRep & vbNewLine & vbNewLine & BoundaryConditions
    
    If UBound(Shapes) - LBound(Shapes) + 1 > 0 Then
        StrRep = StrRep & vbNewLine & vbNewLine & pShapesData
    End If
    
    StrRep = StrRep & vbNewLine & vbNewLine & CustomReportSections
    StrRep = StrRep & vbNewLine & vbNewLine & MemberData
    
    
    If UBound(Plates) - LBound(Plates) + 1 > 0 Then
        StrRep = StrRep & vbNewLine & vbNewLine & ListStrRep("PLATES", Plates)
    End If
    
    If UBound(BLCs) - LBound(BLCs) + 1 > 0 Then
        StrRep = StrRep & vbNewLine & vbNewLine & ListStrRep("BASIC_LOAD_CASES", pBLCs)
    End If
    
    If UBound(PointLoads) - LBound(PointLoads) + 1 > 0 Then
        StrRep = StrRep & vbNewLine & vbNewLine & ListStrRep("POINT_LOADS", pPointLoads)
    End If
    
    If UBound(DistributedLoads) - LBound(DistributedLoads) + 1 > 0 Then
        StrRep = StrRep & vbNewLine & vbNewLine & ListStrRep("DIRECT_DISTRIBUTED_LOADS", pDistributedLoads)
    End If
    
'    If UBound(AreaLoads) - LBound(AreaLoads) + 1 > 0 Then
'        StrRep = StrRep & vbNewLine & vbNewLine & ListStrRep("AREA_LOADS", pAreaLoads)
'    End If
    
'    If UBound(SurfaceLoads) - LBound(SurfaceLoads) + 1 > 0 Then
'        StrRep = StrRep & vbNewLine & vbNewLine & ListStrRep("SURFACE_LOADS", pSurfaceLoads)
'    End If

    StrRep = StrRep & vbNewLine & vbNewLine & FileMidEndData
    
    If UBound(LCs) - LBound(LCs) + 1 > 0 Then
        StrRep = StrRep & vbNewLine & vbNewLine & ListStrRep("LOAD_COMBINATIONS", pLCs)
    End If
    
    StrRep = StrRep & vbNewLine & vbNewLine & FileEndData
End Function



Private Function ListStrRep(Label, Data)
    ListStrRep = "[" & Label & "] <" & CStr(UBound(Data) - LBound(Data) + 1) & ">"
    For Each row In Data
        ListStrRep = ListStrRep & vbNewLine & row.StrRep()
    Next row
    ListStrRep = ListStrRep & vbNewLine & "[END_" & Label & "]"
End Function

Private Function ListStrRep2(Label, Data)
    ListStrRep2 = "[" & Label & "]"
    For Each row In Data
        ListStrRep2 = ListStrRep2 & vbNewLine & vbNewLine & row.StrRep()
    Next row
    ListStrRep2 = ListStrRep2 & vbNewLine & vbNewLine & "[END_" & Label & "]"
End Function



'----------------------------------Views----------------------------------
Public Sub InitViewData()

End Sub

