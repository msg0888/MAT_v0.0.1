Attribute VB_Name = "ViewElevation"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private ChartWidth
Private ChartHeight
Private buffer
Private ViewScale
Private Cht As Chart
Private Center
Private AR
Private X
Private Y
Private ChartName
Private ScaleAxis

Private ColorMap As New Scripting.Dictionary

Private MinX
Private MinY
Private MinZ

Private MaxX
Private MaxY
Private MaxZ

Private ExtraBufferTop
Private ExtraBufferLeft

Private CenterPoint
Private DisplayAngle

Public Sub SetGraphBoundaries(ModelMinX, ModelMinY, ModelMinZ, ModelMaxX, ModelMaxY, ModelMaxZ, Center)
    Set CenterPoint = Center
    
    MinX = ModelMinX
    MinY = ModelMinY
    MinZ = ModelMinZ
    
    MaxX = ModelMaxX
    MaxY = ModelMaxY
    MaxZ = ModelMaxZ
    
        
    If Application.Range("Axes").Value = "XYZ" Then
        MinX = ModelMinY
        MinY = ModelMinZ
        MinZ = ModelMinX
        MaxX = ModelMaxY
        MaxY = ModelMaxZ
        MaxZ = ModelMaxX
        
    End If
    
    If MaxX = MinX Then
        MaxX = MaxX + 1
        MinX = MinX - 1
    End If
    If MaxY = MinY Then
        MaxY = MaxY + 1
        MinY = MinY - 1
    End If
    If MaxZ = MinZ Then
        MaxZ = MaxZ + 1
        MinZ = MinZ - 1
    End If
    
    XScale = (ChartWidth - 2 * buffer) / (MaxX - MinX)
    YScale = (ChartHeight - 2 * buffer) / (MaxY - MinY)
    
    If Range("Scale_Elevation") <> "" Then
        ViewScale = 0.8 * Range("Scale_Elevation") * Application.Min(YScale, XScale)
    Else
        ViewScale = 0.8 * Application.Min(YScale, XScale)
    End If
    ScaleAxis = "X"
    If YScale < XScale Then
        ScaleAxis = "Y"
    End If
    
'    SingleSectorBuffer = 0
'    If Application.Range("QtySectors").Value = 1 Then
'        SingleSectorBuffer = -0.1 * ChartHeight
'    End If
    If Range("Buffer_TopElev") <> "" Then
        ExtraBufferTop = Range("Buffer_TopElev")
    Else
        ExtraBufferTop = (ChartHeight * 0.4 - (MaxY - MinY) * ViewScale) / 2 + (ChartHeight * 0.5 / 2)
    End If
    If Range("Buffer_LeftElev") <> "" Then
        ExtraBufferLeft = Range("Buffer_LeftElev")
    Else
        ExtraBufferLeft = 0 + (ChartWidth * 0.2 / 2)
    End If
    Sheets("Code").Range("EV_Height").Value = ChartHeight
    Sheets("Code").Range("EV_Width").Value = ChartWidth
    Sheets("Code").Range("EV_TopBuff").Value = (ChartHeight * 0.4 - (MaxY - MinY) * ViewScale) / 2 + (ChartHeight * 0.5 / 2)
    Sheets("Code").Range("EV_LeftBuff").Value = 0 + (ChartWidth * 0.2 / 2)
    If ScaleAxis = "Y" Then
        ExtraBufferTop = (ChartHeight * 0.4 / 2)
        ExtraBufferLeft = (ChartWidth * 0.4 - (MaxX - MinX) * ViewScale) / 2 + (ChartWidth * 0.5 / 2)
    End If
End Sub


Private Sub Class_Initialize()
    ChartName = "ElevationView"
    buffer = 8
    
    Dim Alpha As Range
    Set Alpha = Application.Range(ChartName & "Cell_Alpha")
    
    'Get Cell Location and Size
    Y = Alpha.MergeArea.Top + buffer
    X = Alpha.MergeArea.Left + buffer
    ChartWidth = Alpha.MergeArea.Width - 2 * buffer
    ChartHeight = Alpha.MergeArea.Height - 2 * buffer
    
    
    'Delete Existing Chart (If it exists)
    On Error Resume Next
        Sheets("Code").ChartObjects(ChartName).Delete
        Err.Clear
    On Error GoTo 0
    
    'Create Chart
    Set Cht = Sheets("Code").ChartObjects.Add(X, Y, ChartWidth, ChartHeight).Chart
    
    Cht.AutoScaling = False
    Cht.Parent.Name = ChartName
    Cht.ChartArea.Height = ChartHeight
    Cht.ChartArea.Width = ChartWidth
    Cht.ChartArea.Border.LineStyle = xlLineStyleNone
    DisplayAngle = (360 / Sheets("Code").Range("QtySectors").Value) + 2
    
    ColorMap.Add "MemberBorder", RGB(75, 75, 75)
    ColorMap.Add "MemberFill", RGB(150, 150, 150)

'    Answer = MsgBox("Use carrier colors for proposed equipment?", vbYesNo, "Placement Diagram Equipment Color Selection")
'    If Answer = vbNo Then
    If Range("Carrier_Colors") <> True Then
    
    Sheets("Code").Range("AO17").Borders.Color = RGB(75, 105, 32)
    Sheets("Code").Range("AO17").Interior.Color = RGB(253, 119, 0)

        
    'Proposed
    ColorMap.Add "AntennaBorder", RGB(75, 105, 32)
    ColorMap.Add "AntennaFill", RGB(253, 119, 0)
    
    ColorMap.Add "TMEBorder", RGB(75, 105, 32)
    ColorMap.Add "TMEFill", RGB(253, 119, 0)
    
    ColorMap.Add "DishBorder", RGB(75, 105, 32)
    ColorMap.Add "DishFill", RGB(253, 119, 0)

    'Existing
    ColorMap.Add "AntennaBorder_Existing", RGB(0, 0, 0)
    ColorMap.Add "AntennaFill_Existing", RGB(235, 235, 235)
    
    ColorMap.Add "TMEBorder_Existing", RGB(0, 0, 0)
    ColorMap.Add "TMEFill_Existing", RGB(235, 235, 235)
    
    ColorMap.Add "DishBorder_Existing", RGB(0, 0, 0)
    ColorMap.Add "DishFill_Existing", RGB(235, 235, 235)
    Else
    
    'Existing
    ColorMap.Add "AntennaBorder_Existing", RGB(0, 0, 0)
    ColorMap.Add "AntennaFill_Existing", RGB(235, 235, 235)
    
    ColorMap.Add "TMEBorder_Existing", RGB(0, 0, 0)
    ColorMap.Add "TMEFill_Existing", RGB(235, 235, 235)
    
    ColorMap.Add "DishBorder_Existing", RGB(0, 0, 0)
    ColorMap.Add "DishFill_Existing", RGB(235, 235, 235)
    
    If Application.Range("Carrier") = "T-Mobile" Then
        ColorMap.Add "AntennaBorder", RGB(153, 155, 158)
        ColorMap.Add "AntennaFill", RGB(237, 0, 140)

        ColorMap.Add "TMEBorder", RGB(153, 155, 158)
        ColorMap.Add "TMEFill", RGB(237, 0, 140)

        ColorMap.Add "DishBorder", RGB(153, 155, 158)
        ColorMap.Add "DishFill", RGB(237, 0, 140)
        
        Sheets("Code").Range("AO17").Borders.Color = RGB(153, 155, 0)
        Sheets("Code").Range("AO17").Interior.Color = RGB(237, 0, 140)

    ElseIf Range("Carrier") = "Verizon" Then
        ColorMap.Add "AntennaBorder", RGB(0, 0, 0)
        ColorMap.Add "AntennaFill", RGB(236, 28, 36)

        ColorMap.Add "TMEBorder", RGB(0, 0, 0)
        ColorMap.Add "TMEFill", RGB(236, 28, 36)

        ColorMap.Add "DishBorder", RGB(0, 0, 0)
        ColorMap.Add "DishFill", RGB(236, 28, 36)
        
        Sheets("Code").Range("AO17").Borders.Color = RGB(0, 0, 0)
        Sheets("Code").Range("AO17").Interior.Color = RGB(236, 28, 36)

    ElseIf Range("Carrier") = "AT&T" Then
        ColorMap.Add "AntennaBorder", RGB(0, 0, 0)
        ColorMap.Add "AntennaFill", RGB(0, 159, 219)

        ColorMap.Add "TMEBorder", RGB(0, 0, 0)
        ColorMap.Add "TMEFill", RGB(0, 159, 219)

        ColorMap.Add "DishBorder", RGB(0, 0, 0)
        ColorMap.Add "DishFill", RGB(0, 159, 219)
        
        Sheets("Code").Range("AO17").Borders.Color = RGB(0, 0, 0)
        Sheets("Code").Range("AO17").Interior.Color = RGB(0, 159, 219)

    ElseIf Range("Carrier") = "Dish" Then
        ColorMap.Add "AntennaBorder", RGB(0, 0, 0)
        ColorMap.Add "AntennaFill", RGB(225, 58, 62)

        ColorMap.Add "TMEBorder", RGB(0, 0, 0)
        ColorMap.Add "TMEFill", RGB(225, 58, 62)

        ColorMap.Add "DishBorder", RGB(0, 0, 0)
        ColorMap.Add "DishFill", RGB(225, 58, 62)
        
        Sheets("Code").Range("AO17").Borders.Color = RGB(0, 0, 0)
        Sheets("Code").Range("AO17").Interior.Color = RGB(225, 58, 62)

    ElseIf Range("Carrier") = "Shentel" Then
        ColorMap.Add "AntennaBorder", RGB(227, 112, 29)
        ColorMap.Add "AntennaFill", RGB(0, 69, 105)

        ColorMap.Add "TMEBorder", RGB(227, 112, 29)
        ColorMap.Add "TMEFill", RGB(253, 184, 19)

        ColorMap.Add "DishBorder", RGB(227, 112, 29)
        ColorMap.Add "DishFill", RGB(0, 69, 105)
        
        Sheets("Code").Range("AO17").Borders.Color = RGB(153, 155, 0)
        Sheets("Code").Range("AO17").Interior.Color = RGB(0, 69, 105)

    ElseIf Range("Carrier") = "Ericsson" Then
        ColorMap.Add "AntennaBorder", RGB(100, 100, 100)
        ColorMap.Add "AntennaFill", RGB(0, 37, 97)

        ColorMap.Add "TMEBorder", RGB(100, 100, 100)
        ColorMap.Add "TMEFill", RGB(0, 37, 97)

        ColorMap.Add "DishBorder", RGB(100, 100, 100)
        ColorMap.Add "DishFill", RGB(0, 37, 97)
        
        Sheets("Code").Range("AO17").Borders.Color = RGB(100, 100, 100)
        Sheets("Code").Range("AO17").Interior.Color = RGB(0, 37, 97)

    ElseIf Range("Carrier") = "Motorola" Then
        ColorMap.Add "AntennaBorder", RGB(92, 146, 250)
        ColorMap.Add "AntennaFill", RGB(10, 90, 163)

        ColorMap.Add "TMEBorder", RGB(92, 146, 250)
        ColorMap.Add "TMEFill", RGB(10, 90, 163)

        ColorMap.Add "DishBorder", RGB(92, 146, 250)
        ColorMap.Add "DishFill", RGB(10, 90, 163)
        
        Sheets("Code").Range("AO17").Borders.Color = RGB(92, 146, 250)
        Sheets("Code").Range("AO17").Interior.Color = RGB(10, 90, 163)

    ElseIf Range("Carrier") = "Amazon" Then
        ColorMap.Add "AntennaBorder", RGB(254, 189, 105)
        ColorMap.Add "AntennaFill", RGB(35, 47, 62)

        ColorMap.Add "TMEBorder", RGB(254, 189, 105)
        ColorMap.Add "TMEFill", RGB(35, 47, 62)

        ColorMap.Add "DishBorder", RGB(254, 189, 105)
        ColorMap.Add "DishFill", RGB(35, 47, 62)
        
        Sheets("Code").Range("AO17").Borders.Color = RGB(254, 189, 105)
        Sheets("Code").Range("AO17").Interior.Color = RGB(35, 47, 62)

    Else
        ColorMap.Add "AntennaBorder", RGB(75, 105, 32)
        ColorMap.Add "AntennaFill", RGB(253, 119, 0)

        ColorMap.Add "TMEBorder", RGB(75, 105, 32)
        ColorMap.Add "TMEFill", RGB(253, 119, 0)

        ColorMap.Add "DishBorder", RGB(75, 105, 32)
        ColorMap.Add "DishFill", RGB(253, 119, 0)
        
        Sheets("Code").Range("AO17").Borders.Color = RGB(75, 105, 32)
        Sheets("Code").Range("AO17").Interior.Color = RGB(253, 119, 0)
    
    End If
    
    End If

End Sub

Public Sub AddMember(ByRef Mem, ByRef iNode, ByRef jNode)

    Dim ConvertLength As String
    With Sheets("Code")
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength = 1
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength = 12
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength = 0.0254
    End If
    End With

    INodeX = iNode.X
    INodeY = iNode.Y
    INodeZ = iNode.Z
    JNodeX = jNode.X
    JNodeY = jNode.Y
    JNodeZ = jNode.Z
    
    ModelMinX = MinX
    ModelMinY = MinY
    ModelMinZ = MinZ
    ModelMaxY = MaxY
    If Application.Range("QtySectors").Value > 1 Then
        CenterPointX = CenterPoint.X
        CenterPointY = CenterPoint.Y
        CenterPointZ = CenterPoint.Z
    End If
    If Application.Range("Axes").Value = "XYZ" Then
        INodeX = iNode.Y
        INodeY = iNode.Z
        INodeZ = iNode.X
        JNodeX = jNode.Y
        JNodeY = jNode.Z
        JNodeZ = jNode.X

        If Application.Range("QtySectors").Value > 1 Then
            CenterPointX = CenterPoint.Y
            CenterPointY = CenterPoint.Z
            CenterPointZ = CenterPoint.X
        End If
    End If
    
    MemCenterX = (INodeX + JNodeX) / 2
    MemCenterY = (INodeY + JNodeY) / 2
    MemCenterZ = (INodeZ + JNodeZ) / 2
    
    With Application
        dX = .Max(INodeX, JNodeX) - MemCenterX
        If INodeX = .Max(INodeX, JNodeX) Then
            dY = INodeY - MemCenterY
        Else
            dY = JNodeY - MemCenterY
        End If
    End With
        
    If dY <> 0 Then
        'get angle in Degrees
        Angle = -Atn(dX / dY) * 180 / WorksheetFunction.Pi()
    Else
        Angle = 90
    End If
    TempName = Mem.Name

    
    INodeInRange = True
    JNodeInRange = True
    If Application.Range("QtySectors").Value > 1 Then
        INodeRadius = ((INodeX - CenterPointX) ^ 2 + (INodeZ - CenterPointZ) ^ 2) ^ 0.5
        JNodeRadius = ((JNodeX - CenterPointX) ^ 2 + (JNodeZ - CenterPointZ) ^ 2) ^ 0.5
        
        INodeInRange = INodeRadius <> 0
        JNodeInRange = JNodeRadius <> 0
        If JNodeRadius <> 0 And INodeRadius <> 0 Then
            INodeAngle = WorksheetFunction.Asin(Abs(INodeX - CenterPointX) / INodeRadius) * 180 / WorksheetFunction.Pi()
            JNodeAngle = WorksheetFunction.Asin(Abs(JNodeX - CenterPointX) / JNodeRadius) * 180 / WorksheetFunction.Pi()
            
            INodeInRange = INodeAngle < DisplayAngle / 2 And INodeAngle >= 0 And (INodeZ - CenterPointZ) >= 0
            JNodeInRange = JNodeAngle < DisplayAngle / 2 And JNodeAngle >= 0 And (JNodeZ - CenterPointZ) >= 0
        End If
    End If
    
    If INodeInRange Or JNodeInRange Then
        If dY = 0 And dX = 0 Then
            If Mem.IsFlat Then
                AddRectangle Mem.Name, (MemCenterX - ModelMinX) * ViewScale, (ModelMaxY - MemCenterY) * ViewScale, Mem.Dc / ConvertLength * ViewScale, Mem.Dc / ConvertLength * ViewScale
            Else
                AddCircle Mem.Name, (MemCenterX - ModelMinX) * ViewScale, (ModelMaxY - MemCenterY) * ViewScale, Mem.Dc / ConvertLength * ViewScale / 2
            End If
        Else
            AddRectangle Mem.Name, (MemCenterX - ModelMinX) * ViewScale, (ModelMaxY - MemCenterY) * ViewScale, Mem.Dc / ConvertLength * ViewScale, ((INodeX - JNodeX) ^ 2 + (INodeY - JNodeY) ^ 2) ^ 0.5 * ViewScale, Angle
        End If
    End If
End Sub


Public Sub AddEquipment(AppName, memberName, pos, Height, Width, Depth, AppType, Quantity, Condition)
    'Check if the member exists in the chart, if it doesn't, no need to add it
    On Error Resume Next
        Temp = Cht.Shapes(memberName).Name
        If Err.Number <> 0 Or Quantity = 0 Or Quantity = "" Then
            Err.Clear
            Exit Sub
        End If
        If Err.Number <> 0 Or Condition = 0 Or Condition = "" Then
            Err.Clear
            Exit Sub
        End If
    On Error GoTo 0


    Dim ConvertLength As String
    With Sheets("Code")
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength = 1
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength = 12
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength = 0.0254
    End If
    End With

    'Position = Cht.Shapes(MemberName).Height - Pos * ViewScale
    Position = pos * ViewScale
    If pos < 0 Then
        Position = Cht.Shapes(memberName).Height * (Abs(pos)) / 100
    End If

    'Find the coords based on the member position, then remove the buffer distances (It will get added back later)
    AppX = Cht.Shapes(memberName).Left + Position / ConvertLength * CustomFunctions.SIND(-Cht.Shapes(memberName).Rotation) + Cht.Shapes(memberName).Width / 2 - ExtraBufferLeft
    AppY = Cht.Shapes(memberName).Top + Position / ConvertLength * CustomFunctions.COSD(-Cht.Shapes(memberName).Rotation) - ExtraBufferTop
    AppY_Dish = Cht.Shapes(memberName).Top + Position / ConvertLength * CustomFunctions.COSD(-Cht.Shapes(memberName).Rotation) - ExtraBufferTop
    Debug.Print Cht.Shapes(memberName).Top

    If Condition = "Existing" Then
        Select Case AppType
        Case "Sphere"
            If Quantity > 1 Then
                AddCircle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Width / 2 / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing", 0
                AddCircle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Width / 2 / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing", 0
            Else
                AddCircle AppName, AppX, AppY, Width / ConvertLength * ViewScale, "AntennaBorder_Existing", "AntennaFill_Existing", 0
            End If
        Case "Round"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
            End If
        Case "Round (Back)"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
            End If
        Case "Round (Front)"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0.1
            End If
        Case "TME"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
            End If
        Case "TME (Back)"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
            End If
        Case "TME (Front)"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder_Existing", "TMEFill_Existing", 0.1
            End If
        Case "Antenna"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Width / ConvertLength / 2 + 2) * ViewScale, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "AntennaBorder_Existing", "AntennaFill_Existing", 0.4
                AddRectangle AppName & "b", AppX + (Width / ConvertLength / 2 + 2) * ViewScale, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "AntennaBorder_Existing", "AntennaFill_Existing", 0.4
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "AntennaBorder_Existing", "AntennaFill_Existing", 0.1
            End If
        Case Else
                AddCircle AppName, AppX, (AppY + Width / ConvertLength / 2), Height / 2 / ConvertLength * ViewScale, "AntennaBorder_Existing", "AntennaFill_Existing", 0.1
    End Select
    
    Else
    Select Case AppType
        Case "Sphere"
            If Quantity > 1 Then
                AddCircle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Width / ConvertLength * ViewScale, "TMEBorder", "TMEFill", 0
                AddCircle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Width / ConvertLength * ViewScale, "TMEBorder", "TMEFill", 0
            Else
                AddCircle AppName, AppX, AppY, Width / ConvertLength * ViewScale, "TMEBorder", "TMEFill", 0
            End If
        Case "Round"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
            End If
        Case "Round (Back)"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
            End If
        Case "Round (Front)"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0.1
            End If
        Case "TME"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
            End If
        Case "TME (Back)"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
            End If
        Case "TME (Front)"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
                AddRectangle AppName & "b", AppX + (Depth / ConvertLength / 2 + 2) * ViewScale, AppY, Depth / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0, "Back"
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "TMEBorder", "TMEFill", 0.1
            End If
        Case "Antenna"
            If Quantity > 1 Then
                AddRectangle AppName & "a", AppX - (Width / ConvertLength / 2 + 2) * ViewScale, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "AntennaBorder", "AntennaFill", 0.4
                AddRectangle AppName & "b", AppX + (Width / ConvertLength / 2 + 2) * ViewScale, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "AntennaBorder", "AntennaFill", 0.4
            Else
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Height / ConvertLength * ViewScale, 0, "AntennaBorder", "AntennaFill", 0.1
            End If
        Case Else
            AddCircle AppName, AppX, AppY, Height / ConvertLength / 2 * ViewScale, "DishBorder", "DishFill", 0.1
    End Select
    End If

End Sub


Private Sub AddCircle(ShapeName, XCoord, YCoord, Radius, Optional BorderColor = "MemberBorder", Optional FillColor = "MemberFill", Optional Transparency = 0)
    X = (XCoord - Radius)
    Y = (YCoord + Radius)
        
    Set NewShape = Cht.Shapes.AddShape(msoShapeOval, X + ExtraBufferLeft, ChartHeight - Y - ExtraBufferTop, Radius * 2, Radius * 2)
        NewShape.Rotation = -Rot
        NewShape.Left = X + ExtraBufferLeft
        NewShape.Top = Y - Radius * 2 + ExtraBufferTop
        NewShape.Height = Radius * 2
        NewShape.Width = Radius * 2
    
    NewShape.Fill.ForeColor.RGB = ColorMap(FillColor)
    NewShape.line.ForeColor.RGB = ColorMap(BorderColor)
    NewShape.Fill.Transparency = Transparency
    
    If FillColor = "MemberFill" Then
        With NewShape.Shadow
            .Visible = msoTrue
            .Style = msoShadowStyleInnerShadow   ' inner (inside) shadow
            .ForeColor.RGB = RGB(0, 0, 0)        ' shadow color (black)
            .Transparency = 0.6                  ' 0 = opaque, 1 = fully transparent
            .Blur = 6                            ' softness in points
            .OffsetX = -4                        ' negative = left
            .OffsetY = -0                        ' negative = up
        End With
    End If
    
    NewShape.Name = ShapeName
End Sub


Private Sub AddRectangle(ShapeName, XCoord, YCoord, Width, Height, Optional Rot = 0, Optional BorderColor = "MemberBorder", Optional FillColor = "MemberFill", Optional Transparency = 0, Optional Zorder = "None")
    
    If Width <> 0 And Height <> 0 Then
        ExtraAngle = Atn(Round(Width / Height, 2)) * 180 / WorksheetFunction.Pi()
        LengthToCorner = ((Width / 2) ^ 2 + (Height / 2) ^ 2) ^ 0.5
        X = XCoord - LengthToCorner * CustomFunctions.SIND(Rot + ExtraAngle)
        Y = YCoord - LengthToCorner * CustomFunctions.COSD(Rot + ExtraAngle)
        
        
        Set NewShape = Cht.Shapes.AddShape(msoShapeRectangle, XCoord, ChartHeight - YCoord, 2, 2)
        NewShape.Rotation = -Rot
        NewShape.Left = X + ExtraBufferLeft
        NewShape.Top = Y + ExtraBufferTop
        NewShape.Height = Height
        NewShape.Width = Width
        
        
        NewShape.Fill.ForeColor.RGB = ColorMap(FillColor)
        NewShape.Fill.BackColor.RGB = RGB(150, 150, 150)
        NewShape.line.ForeColor.RGB = ColorMap(BorderColor) 'NewShape.line.ForeColor.RGB = ColorMap(BorderColor)
        NewShape.Fill.Transparency = Transparency
        If Zorder = "Back" Then
            NewShape.Zorder msoSendToBack
        End If
        
        If FillColor = "MemberFill" Then
            With NewShape.Shadow
                .Visible = msoTrue
                .Style = msoShadowStyleInnerShadow   ' inner (inside) shadow
                .ForeColor.RGB = RGB(0, 0, 0)        ' shadow color (black)
                .Transparency = 0.6                  ' 0 = opaque, 1 = fully transparent
                .Blur = 6                            ' softness in points
                .OffsetX = -4                        ' negative = left
                .OffsetY = -0                        ' negative = up
            End With
        End If
        
        NewShape.Name = ShapeName
        
'        ' Tag the shape so we can find it later
'        On Error Resume Next
'        NewShape.Tags.Delete "Class": On Error GoTo 0
'        If ShapeType = "Member" Then
'            NewShape.Tags.Add "Class", "Member"
'        Else
'            NewShape.Tags.Add "Class", "Appurt"
'        End If

        Dim rngLabels As Range
        Dim rngAppurts As Range
        Dim isInList As Boolean
        
        Set rngLabels = ThisWorkbook.Worksheets("Code").Range("AE45:AE71")
        Set rngAppurts = ThisWorkbook.Worksheets("Discrete Loads").Range("A4:A503")
        
        isInList = (WorksheetFunction.CountIf(rngLabels, CStr(ShapeName)) > 0) And _
                   (WorksheetFunction.CountIf(rngAppurts, CStr(ShapeName)) = 0)
        
        If isInList _
           And (InStr(Left(ShapeName, 2), "MP") > 0 Or InStr(Left(ShapeName, 1), "A") > 0 _
           Or InStr(Left(ShapeName, 1), "B") > 0 Or InStr(Left(ShapeName, 1), "C") > 0 _
           Or InStr(Left(ShapeName, 1), "D") > 0 Or InStr(Left(ShapeName, 1), "G") > 0 _
           Or InStr(Left(ShapeName, 1), "M") > 0) Then
           'IsNumeric(Right(ShapeName, 1)) Then
            
            'add text box
            BoxWidth = 40
            BoxHeight = 20
            Set TextBox = Cht.Shapes.AddTextbox( _
                Orientation:=msoTextOrientationHorizontal, _
                Left:=NewShape.Left - BoxWidth / 2 + Width / 2, _
                Top:=ChartHeight - BoxHeight, _
                Width:=BoxWidth, _
                Height:=BoxHeight)
                
            TextBox.Name = ShapeName & "_Text"
            'TextBox.ShapeStyle = msoShapeStylePreset72
            TextBox.TextFrame.Characters.Text = ShapeName
            TextBox.AutoShapeType = msoShapeRoundedRectangle
            TextBox.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            TextBox.TextFrame2.VerticalAnchor = msoAnchorMiddle
            TextBox.TextFrame2.TextRange.Font.Bold = msoTrue
            TextBox.TextFrame2.TextRange.Font.Size = 9
        End If
    End If
    
    On Error Resume Next
    If Range("MCL_Check") = True Then
        DrawChartCenterDash Cht, "MCL", ColorMap("MemberFill")
    End If

End Sub


Public Sub DrawChartCenterDash(ByVal Cht As Chart, _
                               ByVal labelText As String, _
                               ByVal dashColor As Long)

    Const nameLine$ = "CenterDash"
    Const nameText$ = "CenterDash_Text"
    Dim CenterNode As Range, SectionSet As Range, nY As Range, ProjectHeight As Range

    ' Clean up previous instances
    On Error Resume Next
    Cht.Shapes(nameLine).Delete
    Cht.Shapes(nameText).Delete
    On Error GoTo 0

    ' Plot Area coordinates
    Dim l As Double, t As Double, W As Double, h As Double
    Set CenterNode = Range("CenterPoint")
    Set SectionSet = WorksheetFunction.XLookup(Sheets("Code").Range("BB39"), Sheets("Geometry").Range("B:B"), Sheets("Geometry").Range("C:C"))
    With Cht.PlotArea
        l = .InsideLeft
        t = .InsideTop - WorksheetFunction.XLookup(CenterNode, Sheets("Geometry").Range("S:S"), Sheets("Geometry").Range("U:U")).Value - WorksheetFunction.XLookup(SectionSet, Sheets("Geometry").Range("X:X"), Sheets("Geometry").Range("AB:AB")).Value
        W = .InsideWidth
        h = .InsideHeight - WorksheetFunction.XLookup(CenterNode, Sheets("Geometry").Range("S:S"), Sheets("Geometry").Range("U:U")).Value
        '- WorksheetFunction.XLookup(SectionSet, Sheets("Geometry").Range("X:X"), Sheets("Geometry").Range("AB:AB")).Value
    End With

    ' Draw dashed centerline (horizontal across plot area's middle)
    Dim yMid As Double: yMid = t + h / 2

    Dim lineSh As Excel.Shape
    Set lineSh = Cht.Shapes.AddLine(l, yMid, l + W, yMid)
    With lineSh
        .Name = nameLine
        With .line
            .Visible = msoTrue
            .DashStyle = msoLineDash
            .Weight = 1.25
            .ForeColor.RGB = dashColor
        End With
    End With

    ' --- Styled label like your AddRectangle text box ---
    Dim BoxWidth As Double:  BoxWidth = 36
    Dim BoxHeight As Double: BoxHeight = 20
    Dim padX As Double: padX = -12      ' nudge right from left edge
    Dim padY As Double: padY = -10      ' nudge up from the line

    Dim lbl As Excel.Shape
    Set lbl = Cht.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        l + padX - BoxWidth / 2, _
        yMid - BoxHeight - padY, _
        BoxWidth, BoxHeight)

    With lbl
        .Name = nameText
        '.ShapeStyle = msoShapeStylePreset72
        .TextFrame.Characters.Text = labelText

        ' Match your AddRectangle styling
        .AutoShapeType = msoShapeRoundedRectangle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 8

        .Zorder msoBringToFront
    End With
End Sub


Public Sub ListClearances(ByVal Cht As Chart, _
                          ByVal OutWS As Worksheet, _
                          ByVal StartCell As String)

    Dim sh As Excel.Shape
    Dim Members() As Variant, appurts() As Variant
    Dim mCnt As Long, aCnt As Long

    ' Collect rectangles with our tags
    For Each sh In Cht.Shapes
        If sh.Type = msoAutoShape And sh.AutoShapeType = msoShapeRectangle Then
            Dim cls As String
            On Error Resume Next
            cls = sh.Tags("Class")
            On Error GoTo 0

            If cls = "" Then
                ' Fallback: treat names starting with "MP" as appurtenances
                If Left$(sh.Name, 2) = "MP" Then
                    cls = "Appurt"
                Else
                    cls = "Member"
                End If
            End If

            If cls = "Member" Then
                mCnt = mCnt + 1
                ReDim Preserve Members(1 To 6, 1 To mCnt)
                Members(1, mCnt) = sh.Name
                Members(2, mCnt) = sh.Left
                Members(3, mCnt) = sh.Top
                Members(4, mCnt) = sh.Width
                Members(5, mCnt) = sh.Height
                Members(6, mCnt) = sh.Rotation ' for info only
            ElseIf cls = "Appurt" Then
                aCnt = aCnt + 1
                ReDim Preserve appurts(1 To 6, 1 To aCnt)
                appurts(1, aCnt) = sh.Name
                appurts(2, aCnt) = sh.Left
                appurts(3, aCnt) = sh.Top
                appurts(4, aCnt) = sh.Width
                appurts(5, aCnt) = sh.Height
                appurts(6, aCnt) = sh.Rotation
            End If
        End If
    Next sh

    Dim r As Range
    Set r = OutWS.Range(StartCell)

    '==== A) Appurtenance ? Appurtenance pairwise clearances
    r.offset(0, 0).Value = "Appurtenance-to-Appurtenance Clearances"
    r.offset(1, 0).Resize(1, 6).Value = Array("Appurt A", "Appurt B", _
                                              "Horiz Clear (in)", "Vert Clear (in)", _
                                              "A Overlaps Vert?", "A Overlaps Horiz?")

    Dim i As Long, j As Long, rowOff As Long
    rowOff = 2
    For i = 1 To aCnt
        For j = i + 1 To aCnt
            Dim Aleft As Double, Atop As Double, Aright As Double, Abot As Double
            Dim Bleft As Double, Btop As Double, Bright As Double, Bbot As Double

            Aleft = appurts(2, i): Atop = appurts(3, i)
            Aright = Aleft + appurts(4, i): Abot = Atop + appurts(5, i)

            Bleft = appurts(2, j): Btop = appurts(3, j)
            Bright = Bleft + appurts(4, j): Bbot = Btop + appurts(5, j)

            ' Overlap tests (Excel Y grows downward)
            Dim overlapVert As Boolean, overlapHoriz As Boolean
            overlapVert = Not (Abot < Btop Or Bbot < Atop)
            overlapHoriz = Not (Aright < Bleft Or Bright < Aleft)

            ' Edge-to-edge clearances (>=0)
            Dim hClear As Double, vClear As Double
            If Aright <= Bleft Then
                hClear = Bleft - Aright
            ElseIf Bright <= Aleft Then
                hClear = Aleft - Bright
            Else
                hClear = 0
            End If

            If Abot <= Btop Then
                vClear = Btop - Abot
            ElseIf Bbot <= Atop Then
                vClear = Atop - Bbot
            Else
                vClear = 0
            End If

            r.offset(rowOff, 0).Resize(1, 6).Value = _
                Array(appurts(1, i), appurts(1, j), hClear, vClear, overlapVert, overlapHoriz)
            rowOff = rowOff + 1
        Next j
    Next i

    '==== B) Appurtenance ? nearest Member (edge of steel) clearances
    Dim startB As Range
    Set startB = r.offset(rowOff + 2, 0)
    startB.Value = "Appurtenance to Nearest Member (Edge of Steel)"
    startB.offset(1, 0).Resize(1, 6).Value = _
        Array("Appurt", "Nearest Member (Horiz)", "Horiz Clear (in)", _
              "Nearest Member (Vert)", "Vert Clear (in)", "Notes")

    rowOff = rowOff + 4

    For i = 1 To aCnt
        Aleft = appurts(2, i): Atop = appurts(3, i)
        Aright = Aleft + appurts(4, i): Abot = Atop + appurts(5, i)

        Dim minH As Double: minH = 1E+30
        Dim minHMember As String: minHMember = ""
        Dim minV As Double: minV = 1E+30
        Dim minVMember As String: minVMember = ""
        Dim note As String: note = ""

        For j = 1 To mCnt
            Bleft = Members(2, j): Btop = Members(3, j)
            Bright = Bleft + Members(4, j): Bbot = Btop + Members(5, j)

            ' Horizontal clear only if vertical ranges overlap
            overlapVert = Not (Abot < Btop Or Bbot < Atop)
            If overlapVert Then
                If Aright <= Bleft Then
                    hClear = Bleft - Aright
                ElseIf Bright <= Aleft Then
                    hClear = Aleft - Bright
                Else
                    hClear = 0
                End If
                If hClear < minH Then
                    minH = hClear: minHMember = Members(1, j)
                End If
            End If

            ' Vertical clear only if horizontal ranges overlap
            overlapHoriz = Not (Aright < Bleft Or Bright < Aleft)
            If overlapHoriz Then
                If Abot <= Btop Then
                    vClear = Btop - Abot
                ElseIf Bbot <= Atop Then
                    vClear = Atop - Bbot
                Else
                    vClear = 0
                End If
                If vClear < minV Then
                    minV = vClear: minVMember = Members(1, j)
                End If
            End If
        Next j

        If minH = 1E+30 Then minH = VBA.CDbl(Application.Caller) * 0 + 0: note = note & "No vert overlap for horiz clear. "
        If minV = 1E+30 Then minV = VBA.CDbl(Application.Caller) * 0 + 0: note = note & "No horiz overlap for vert clear. "

        OutWS.Range(StartCell).offset(rowOff, 0).Resize(1, 6).Value = _
            Array(appurts(1, i), minHMember, minH, minVMember, minV, Trim$(note))
        rowOff = rowOff + 1
    Next i

End Sub

