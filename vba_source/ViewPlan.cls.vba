Attribute VB_Name = "ViewPlan"
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

Private MountAzimuth


Public Sub SetGraphBoundaries(ModelMinX, ModelMinY, ModelMinZ, ModelMaxX, ModelMaxY, ModelMaxZ)
    
    
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
    YScale = (ChartHeight - 2 * buffer) / (MaxZ - MinZ)
    
    If Range("Scale_Plan") <> "" Then
        ViewScale = 0.8 * Range("Scale_Plan") * Application.Min(YScale, XScale)
    Else
        ViewScale = 0.8 * Application.Min(YScale, XScale)
    End If
    ScaleAxis = "X"
    If YScale < XScale Then
        ScaleAxis = "Y"
    End If
    
    SingleSectorBuffer = 0
    If Application.Range("QtySectors").Value = 1 Then
        SingleSectorBuffer = -0.1 * ChartHeight
    End If
    If Range("Buffer_TopPlan") <> "" Then
        ExtraBufferTop = Range("Buffer_TopPlan")
    Else
        ExtraBufferTop = SingleSectorBuffer + (ChartHeight * 0.5 - (MaxZ - MinZ) * ViewScale) / 2 + (ChartHeight * 0.4 / 2)
    End If
    If Range("Buffer_LeftPlan") <> "" Then
        ExtraBufferLeft = Range("Buffer_LeftPlan")
    Else
        ExtraBufferLeft = 0 + (ChartWidth * 0.2 / 2)
    End If
    Sheets("Code").Range("PV_Height").Value = ChartHeight
    Sheets("Code").Range("PV_Width").Value = ChartWidth
    Sheets("Code").Range("PV_TopBuff").Value = SingleSectorBuffer + (ChartHeight * 0.5 - (MaxZ - MinZ) * ViewScale) / 2 + (ChartHeight * 0.4 / 2)
    Sheets("Code").Range("PV_LeftBuff").Value = 0 + (ChartWidth * 0.2 / 2)
    If ScaleAxis = "Y" Then
        ExtraBufferTop = SingleSectorBuffer + (ChartHeight * 0.2 / 2)
        ExtraBufferLeft = (ChartWidth * 0.45 - (MaxX - MinX) * ViewScale) / 2 + (ChartWidth * 0.5 / 2)
    End If
End Sub


Private Sub Class_Initialize()
    ChartName = "PlanView"
    buffer = 8
    
    Dim Answer As String
    Dim r As Range
    Set r = Application.Range(ChartName & "Cell")
    
    'Get Cell Location and Size
    Y = r.MergeArea.Top + buffer
    X = r.MergeArea.Left + buffer
    ChartWidth = r.MergeArea.Width - 2 * buffer
    ChartHeight = r.MergeArea.Height - 2 * buffer
    
    
    'Delete Existing Chart (If it exists)
    On Error Resume Next
        Sheets("Code").ChartObjects(ChartName).Delete
    On Error GoTo 0
    
    'Create Chart
    Set Cht = Sheets("Code").ChartObjects.Add(X, Y, ChartWidth, ChartHeight).Chart
    
    MountAzimuth = Application.Range("AlphaMountAzimuth").Value
    
    Cht.AutoScaling = False
    Cht.Parent.Name = ChartName
    Cht.ChartArea.Height = ChartHeight
    Cht.ChartArea.Width = ChartWidth
    Cht.ChartArea.Border.LineStyle = xlLineStyleNone
    Cht.ChartArea.Fill.Visible = msoFalse
    
    ColorMap.Add "MemberBorder", RGB(75, 75, 75)
'    ColorMap.Add "BackMemberFill", RGB(85, 85, 85)
    ColorMap.Add "MemberFill", RGB(150, 150, 150)
            
'    Answer = MsgBox("Use carrier colors for proposed equipment?", vbYesNo, "Placement Diagram Equipment Color Selection")
'    If Answer = vbNo Then
    If Range("Carrier_Colors") <> True Then
    
    Sheets("Code").Range("AT17").Borders.Color = RGB(75, 105, 32)
    Sheets("Code").Range("AT17").Interior.Color = RGB(253, 119, 0)

        
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
        
        Sheets("Code").Range("AT17").Borders.Color = RGB(153, 155, 0)
        Sheets("Code").Range("AT17").Interior.Color = RGB(237, 0, 140)

    ElseIf Range("Carrier") = "Verizon" Then
        ColorMap.Add "AntennaBorder", RGB(0, 0, 0)
        ColorMap.Add "AntennaFill", RGB(236, 28, 36)

        ColorMap.Add "TMEBorder", RGB(0, 0, 0)
        ColorMap.Add "TMEFill", RGB(236, 28, 36)

        ColorMap.Add "DishBorder", RGB(0, 0, 0)
        ColorMap.Add "DishFill", RGB(236, 28, 36)
        
        Sheets("Code").Range("AT17").Borders.Color = RGB(0, 0, 0)
        Sheets("Code").Range("AT17").Interior.Color = RGB(236, 28, 36)

    ElseIf Range("Carrier") = "AT&T" Then
        ColorMap.Add "AntennaBorder", RGB(0, 0, 0)
        ColorMap.Add "AntennaFill", RGB(0, 159, 219)

        ColorMap.Add "TMEBorder", RGB(0, 0, 0)
        ColorMap.Add "TMEFill", RGB(0, 159, 219)

        ColorMap.Add "DishBorder", RGB(0, 0, 0)
        ColorMap.Add "DishFill", RGB(0, 159, 219)
        
        Sheets("Code").Range("AT17").Borders.Color = RGB(0, 0, 0)
        Sheets("Code").Range("AT17").Interior.Color = RGB(0, 159, 219)

    ElseIf Range("Carrier") = "Dish" Then
        ColorMap.Add "AntennaBorder", RGB(0, 0, 0)
        ColorMap.Add "AntennaFill", RGB(225, 58, 62)

        ColorMap.Add "TMEBorder", RGB(0, 0, 0)
        ColorMap.Add "TMEFill", RGB(225, 58, 62)

        ColorMap.Add "DishBorder", RGB(0, 0, 0)
        ColorMap.Add "DishFill", RGB(225, 58, 62)
        
        Sheets("Code").Range("AT17").Borders.Color = RGB(0, 0, 0)
        Sheets("Code").Range("AT17").Interior.Color = RGB(225, 58, 62)

    ElseIf Range("Carrier") = "Shentel" Then
        ColorMap.Add "AntennaBorder", RGB(227, 112, 29)
        ColorMap.Add "AntennaFill", RGB(0, 69, 105)

        ColorMap.Add "TMEBorder", RGB(227, 112, 29)
        ColorMap.Add "TMEFill", RGB(253, 184, 19)

        ColorMap.Add "DishBorder", RGB(227, 112, 29)
        ColorMap.Add "DishFill", RGB(0, 69, 105)
        
        Sheets("Code").Range("AT17").Borders.Color = RGB(153, 155, 0)
        Sheets("Code").Range("AT17").Interior.Color = RGB(0, 69, 105)

    ElseIf Range("Carrier") = "Ericsson" Then
        ColorMap.Add "AntennaBorder", RGB(100, 100, 100)
        ColorMap.Add "AntennaFill", RGB(0, 37, 97)

        ColorMap.Add "TMEBorder", RGB(100, 100, 100)
        ColorMap.Add "TMEFill", RGB(0, 37, 97)

        ColorMap.Add "DishBorder", RGB(100, 100, 100)
        ColorMap.Add "DishFill", RGB(0, 37, 97)
        
        Sheets("Code").Range("AT17").Borders.Color = RGB(100, 100, 100)
        Sheets("Code").Range("AT17").Interior.Color = RGB(0, 37, 97)

    ElseIf Range("Carrier") = "Motorola" Then
        ColorMap.Add "AntennaBorder", RGB(92, 146, 250)
        ColorMap.Add "AntennaFill", RGB(10, 90, 163)

        ColorMap.Add "TMEBorder", RGB(92, 146, 250)
        ColorMap.Add "TMEFill", RGB(10, 90, 163)

        ColorMap.Add "DishBorder", RGB(92, 146, 250)
        ColorMap.Add "DishFill", RGB(10, 90, 163)
        
        Sheets("Code").Range("AT17").Borders.Color = RGB(92, 146, 250)
        Sheets("Code").Range("AT17").Interior.Color = RGB(10, 90, 163)

    ElseIf Range("Carrier") = "Amazon" Then
        ColorMap.Add "AntennaBorder", RGB(254, 189, 105)
        ColorMap.Add "AntennaFill", RGB(35, 47, 62)

        ColorMap.Add "TMEBorder", RGB(254, 189, 105)
        ColorMap.Add "TMEFill", RGB(35, 47, 62)

        ColorMap.Add "DishBorder", RGB(254, 189, 105)
        ColorMap.Add "DishFill", RGB(35, 47, 62)
        
        Sheets("Code").Range("AT17").Borders.Color = RGB(254, 189, 105)
        Sheets("Code").Range("AT17").Interior.Color = RGB(35, 47, 62)
        
    Else
        ColorMap.Add "AntennaBorder", RGB(87, 90, 93)
        ColorMap.Add "AntennaFill", RGB(237, 125, 49)

        ColorMap.Add "TMEBorder", RGB(87, 90, 93)
        ColorMap.Add "TMEFill", RGB(237, 125, 49)

        ColorMap.Add "DishBorder", RGB(87, 90, 93)
        ColorMap.Add "DishFill", RGB(237, 125, 49)
        
        Sheets("Code").Range("AT17").Borders.Color = RGB(87, 90, 93)
        Sheets("Code").Range("AT17").Interior.Color = RGB(237, 125, 49)
    
    End If
    
    End If
    
End Sub


Public Sub AddMember(ByRef Mem, ByRef iNode, ByRef jNode)
    INodeX = iNode.X
    INodeY = iNode.Y
    INodeZ = iNode.Z
    JNodeX = jNode.X
    JNodeY = jNode.Y
    JNodeZ = jNode.Z
    
    ModelMinX = MinX
    ModelMinY = MinY
    ModelMinZ = MinZ
    
    If Application.Range("Axes").Value = "XYZ" Then
        INodeX = iNode.Y
        INodeY = iNode.Z
        INodeZ = iNode.X
        JNodeX = jNode.Y
        JNodeY = jNode.Z
        JNodeZ = jNode.X

    End If
    
    MemCenterX = (INodeX + JNodeX) / 2
    MemCenterY = (INodeY + JNodeY) / 2
    MemCenterZ = (INodeZ + JNodeZ) / 2
    
    With Application
        dX = .Max(INodeX, JNodeX) - MemCenterX
        If INodeX = .Max(INodeX, JNodeX) Then
            dZ = INodeZ - MemCenterZ
        Else
            dZ = JNodeZ - MemCenterZ
        End If
    End With
        
    If dZ <> 0 Then
        'get angle in Degrees
        Angle = Atn(dX / dZ) * 180 / WorksheetFunction.Pi()
    Else
        Angle = 90
    End If

    Dim ConvertLength As String
    With shCode
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength = 1
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength = 12
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength = 0.0254
    End If
    End With
    
    Dim dimScale As String
    dimScale = 1 + Range("dimScale_Member").Value / 100

    If dZ = 0 And dX = 0 Then
        If Mem.IsFlat Then
            AddRectangle Mem.Name, (MemCenterX - ModelMinX) * ViewScale, (MemCenterZ - ModelMinZ) * ViewScale, Mem.Dc / ConvertLength * ViewScale, Mem.Dc / ConvertLength * ViewScale
        Else
            AddCircle Mem.Name, (MemCenterX * dimScale - ModelMinX) * ViewScale, (MemCenterZ * dimScale - ModelMinZ) * ViewScale, Mem.Dc / ConvertLength * ViewScale / 2
        End If
    Else
        AddRectangle Mem.Name, (MemCenterX - ModelMinX) * ViewScale, (MemCenterZ - ModelMinZ) * ViewScale, Mem.Dc / ConvertLength * ViewScale, ((INodeX - JNodeX) ^ 2 + (INodeZ - JNodeZ) ^ 2) ^ 0.5 * ViewScale, Angle
    End If
End Sub


Public Sub AddEquipment(AppName, memberName, Height, Width, Depth, AppType, Quantity, Condition, Rotation)
    'Check if the member exists in the chart, if it doesn't, no need to add it
    On Error Resume Next
        Temp = Cht.Shapes(memberName).Top
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
    Dim ConvertLength2 As String
    With shCode
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength = 1
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength = 12
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength = 0.0254
    End If
    End With

    Dim dimScale As String
    dimScale = 1 + Range("dimScale_Member").Value / 100 + Range("dimScale_Appurtenance").Value / 10

    'Find the coords based on the member position, then remove the buffer distances (It will get added back later)
    MemberDepth = Cht.Shapes(memberName).Height
    MemberWidth = Cht.Shapes(memberName).Width
    MemberX = Cht.Shapes(memberName).Left + MemberWidth - ExtraBufferLeft
    MemberY = Cht.Shapes(memberName).Top + MemberDepth - ExtraBufferTop
    Rot = -(Rotation - MountAzimuth)
    
    If Condition = "Existing" Then
        Select Case AppType
        Case "Antenna"
            If Quantity > 1 Then
                Radius = ((MemberWidth / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberDepth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((MemberWidth / 2 + Width * ViewScale / ConvertLength / 2) / (Depth * ViewScale / ConvertLength / 2 + MemberDepth / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "a", AppX1, AppY1, Width / ConvertLength * ViewScale, Depth / ConvertLength * ViewScale, Rot, "AntennaBorder_Existing", "AntennaFill_Existing"
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "b", AppX2, AppY2, Width / ConvertLength * ViewScale, Depth / ConvertLength * ViewScale, Rot, "AntennaBorder_Existing", "AntennaFill_Existing"
            Else
                Radius = MemberDepth / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Depth / ConvertLength * ViewScale, Rot, "AntennaBorder_Existing", "AntennaFill_Existing"
            End If
        
        Case "Sphere"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width / ConvertLength * ViewScale / 2) ^ 2 + (Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width / ConvertLength * ViewScale / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "a", AppX1, AppY1, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "b", AppX2, AppY2, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
            Else
                Radius = MemberDepth / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * -Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName, AppX, AppY, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
            End If
        
        Case "Round"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width / ConvertLength * ViewScale / 2) ^ 2 + (Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width / ConvertLength * ViewScale / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "a", AppX1, AppY1, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "b", AppX2, AppY2, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
            Else
                Radius = MemberDepth / 2 + Depth / ConvertLength * ViewScale / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * -Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName, AppX, AppY, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
            End If
        
        Case "Round (Back)"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width / ConvertLength * ViewScale / 2) ^ 2 + (Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width / ConvertLength * ViewScale / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "a", AppX1, AppY1, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "b", AppX2, AppY2, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
            Else
                Radius = MemberDepth / 2 + Depth / ConvertLength * ViewScale / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * -Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName, AppX, AppY, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
            End If
        
        Case "Round (Front)"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width / ConvertLength * ViewScale / 2) ^ 2 + (Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width / ConvertLength * ViewScale / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddCircle AppName & "a", AppX1, AppY1, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddCircle AppName & "b", AppX2, AppY2, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
            Else
                Radius = MemberDepth / 2 + Depth / ConvertLength * ViewScale / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * Radius * dimScale '/ ConvertLength
                AddCircle AppName, AppX, AppY, Width / ConvertLength * ViewScale, "TMEBorder_Existing", "TMEFill_Existing"
            End If
        
        Case "TME"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width / ConvertLength * ViewScale / 2) ^ 2 + (Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width / ConvertLength * ViewScale / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "a", AppX1, AppY1, Depth / ConvertLength * ViewScale, Width / ConvertLength * ViewScale, Rot, "TMEBorder_Existing", "TMEFill_Existing"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "b", AppX2, AppY2, Depth / ConvertLength * ViewScale, Width * ViewScale / ConvertLength, Rot, "TMEBorder_Existing", "TMEFill_Existing"
            Else
                Radius = MemberDepth / 2 + Depth / ConvertLength * ViewScale / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * -Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Depth / ConvertLength * ViewScale, Rot, "TMEBorder_Existing", "TMEFill_Existing"
            End If
        
        Case "TME (Back)"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width / ConvertLength * ViewScale / 2) ^ 2 + (Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width / ConvertLength * ViewScale / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "a", AppX1, AppY1, Depth / ConvertLength * ViewScale, Width / ConvertLength * ViewScale, Rot, "TMEBorder_Existing", "TMEFill_Existing"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "b", AppX2, AppY2, Depth / ConvertLength * ViewScale, Width / ConvertLength * ViewScale, Rot, "TMEBorder_Existing", "TMEFill_Existing"
            Else
                Radius = MemberDepth / 2 + Depth / ConvertLength * ViewScale / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * -Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName, AppX, AppY, Width / ConvertLength * ViewScale, Depth / ConvertLength * ViewScale, Rot, "TMEBorder_Existing", "TMEFill_Existing"
            End If
        
        Case "TME (Front)"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width / ConvertLength * ViewScale / 2) ^ 2 + (Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth / ConvertLength * ViewScale / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width / ConvertLength * ViewScale / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "a", AppX1, AppY1, Depth / ConvertLength * ViewScale, Width / ConvertLength * ViewScale, Rot, "TMEBorder_Existing", "TMEFill_Existing"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "b", AppX2, AppY2, Depth * ViewScale / ConvertLength, Width * ViewScale / ConvertLength, Rot, "TMEBorder_Existing", "TMEFill_Existing"
            Else
                Radius = MemberDepth / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "TMEBorder_Existing", "TMEFill_Existing"
            End If
                
        Case Else
            If Quantity > 1 Then
                Radius = ((MemberWidth / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberDepth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((MemberWidth / 2 + Width * ViewScale / ConvertLength / 2) / (Depth * ViewScale / ConvertLength / 2 + MemberDepth / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                If Range("Shape_Dish") = True Then
                    AddSemiCircle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "DishBorder_Existing", "DishFill_Existing"
                Else
                    AddRectangle AppName & "a", AppX1, AppY1, Width * ViewScale / ConvertLength, Depth / 3 * ViewScale / ConvertLength, Rot, "DishBorder_Existing", "DishFill_Existing"
                End If
            Else
                Radius = MemberDepth / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * Radius * dimScale '/ ConvertLength
                If Range("Shape_Dish") = True Then
                    AddSemiCircle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "DishBorder_Existing", "DishFill_Existing"
                Else
                    AddRectangle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth / 3 * ViewScale / ConvertLength, Rot, "DishBorder_Existing", "DishFill_Existing"
                End If
            End If
            
    End Select

    Else
    Select Case AppType
        Case "Sphere"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width * ViewScale / ConvertLength / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "a", AppX1, AppY1, Width * ViewScale / ConvertLength / ConvertLength, "TMEBorder", "TMEFill"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "b", AppX2, AppY2, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
            Else
                Radius = MemberDepth / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * -Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName, AppX, AppY, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
            End If
        
        Case "Round"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width * ViewScale / ConvertLength / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "a", AppX1, AppY1, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "b", AppX2, AppY2, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
            Else
                Radius = MemberDepth / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * -Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName, AppX, AppY, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
            End If
        
        Case "Round (Back)"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width * ViewScale / ConvertLength / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "a", AppX1, AppY1, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName & "b", AppX2, AppY2, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
            Else
                Radius = MemberDepth / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * -Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * -Radius * dimScale '/ ConvertLength
                AddCircle AppName, AppX, AppY, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
            End If
        
        Case "Round (Front)"
            If Quantity > 1 Then
                Radius = ((MemberDepth / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width * ViewScale / ConvertLength / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddCircle AppName & "a", AppX1, AppY1, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddCircle AppName & "b", AppX2, AppY2, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
            Else
                Radius = MemberDepth / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * Radius * dimScale '/ ConvertLength
                AddCircle AppName, AppX, AppY, Width * ViewScale / ConvertLength, "TMEBorder", "TMEFill"
            End If
        
        Case "TME"
            If Quantity > 1 Then
                Radius = ((MemberDepth / ConvertLength / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width * ViewScale / ConvertLength / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "a", AppX1, AppY1, Depth * ViewScale / ConvertLength, Width * ViewScale / ConvertLength, Rot, "TMEBorder", "TMEFill"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "b", AppX2, AppY2, Depth * ViewScale / ConvertLength, Width * ViewScale / ConvertLength, Rot, "TMEBorder", "TMEFill"
            Else
                Radius = MemberDepth / ConvertLength / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * -Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "TMEBorder", "TMEFill"
            End If
        
        Case "TME (Back)"
            If Quantity > 1 Then
                Radius = ((MemberDepth / ConvertLength / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width * ViewScale / ConvertLength / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "a", AppX1, AppY1, Depth * ViewScale / ConvertLength, Width * ViewScale / ConvertLength, Rot, "TMEBorder", "TMEFill"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "b", AppX2, AppY2, Depth * ViewScale / ConvertLength, Width * ViewScale / ConvertLength, Rot, "TMEBorder", "TMEFill"
            Else
                Radius = MemberDepth / ConvertLength / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * -Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * -Radius * dimScale '/ ConvertLength
                AddRectangle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "TMEBorder", "TMEFill"
            End If
        
        Case "TME (Front)"
            If Quantity > 1 Then
                Radius = ((MemberDepth / ConvertLength / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((Depth * ViewScale / ConvertLength / 2 + MemberWidth / 2) / (MemberDepth / 2 + Width * ViewScale / ConvertLength / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "a", AppX1, AppY1, Depth * ViewScale / ConvertLength, Width * ViewScale / ConvertLength, Rot, "TMEBorder", "TMEFill"
                
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "b", AppX2, AppY2, Depth * ViewScale / ConvertLength, Width * ViewScale / ConvertLength, Rot, "TMEBorder", "TMEFill"
            Else
                Radius = MemberDepth / ConvertLength / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "TMEBorder", "TMEFill"
            End If
        
        Case "Antenna"
            If Quantity > 1 Then
                Radius = ((MemberWidth / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberDepth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((MemberWidth / 2 + Width * ViewScale / ConvertLength / 2) / (Depth * ViewScale / ConvertLength / 2 + MemberDepth / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "a", AppX1, AppY1, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "AntennaBorder", "AntennaFill"
                
                AppX2 = MemberX + CustomFunctions.SIND(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY2 = MemberY + CustomFunctions.COSD(Rot + ExtraAngle) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName & "b", AppX2, AppY2, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "AntennaBorder", "AntennaFill"
            Else
                Radius = MemberDepth / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * Radius * dimScale '/ ConvertLength
                AddRectangle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "AntennaBorder", "AntennaFill"
            End If
        
        Case Else
            If Quantity > 1 Then
                Radius = ((MemberWidth / 2 + Width * ViewScale / ConvertLength / 2) ^ 2 + (Depth * ViewScale / ConvertLength / 2 + MemberDepth / 2) ^ 2) ^ 0.5
                ExtraAngle = Atn((MemberWidth / 2 + Width * ViewScale / ConvertLength / 2) / (Depth * ViewScale / ConvertLength / 2 + MemberDepth / 2)) * 180 / WorksheetFunction.Pi()
                AppX1 = MemberX + CustomFunctions.SIND(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                AppY1 = MemberY + CustomFunctions.COSD(Rot - ExtraAngle) * Radius * dimScale '/ ConvertLength
                If Range("Shape_Dish") = True Then
                    AddSemiCircle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "DishBorder", "DishFill"
                Else
                    AddRectangle AppName & "a", AppX1, AppY1, Width * ViewScale / ConvertLength, Depth / 3 * ViewScale / ConvertLength, Rot, "DishBorder", "DishFill"
                End If
            Else
                Radius = MemberDepth / 2 + Depth * ViewScale / ConvertLength / 2
                AppX = MemberX + CustomFunctions.SIND(Rot) * Radius * dimScale '/ ConvertLength
                AppY = MemberY + CustomFunctions.COSD(Rot) * Radius * dimScale '/ ConvertLength
                If Range("Shape_Dish") = True Then
                    AddSemiCircle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth * ViewScale / ConvertLength, Rot, "DishBorder", "DishFill"
                Else
                    AddRectangle AppName, AppX, AppY, Width * ViewScale / ConvertLength, Depth / 3 * ViewScale / ConvertLength, Rot, "DishBorder", "DishFill"
                End If
            End If
            
    End Select
    End If

End Sub


Private Sub AddCircle(ShapeName, XCoord, YCoord, Radius, Optional BorderColor = "MemberBorder", Optional FillColor = "MemberFill", Optional Transparency = 0)
    'Plots a circle based on its center
    X = (XCoord - Radius)
    Y = (YCoord - Radius)
    
    
    Set NewShape = Cht.Shapes.AddShape(msoShapeOval, X + ExtraBufferLeft, Y + ExtraBufferTop, Radius * 2, Radius * 2)
    
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
            .OffsetX = -0                        ' negative = left
            .OffsetY = -0                        ' negative = up
        End With
    End If
    
    NewShape.Name = ShapeName
End Sub

Private Sub AddRectangle(ShapeName, XCoord, YCoord, Width, Height, Optional Rotation = 0, Optional BorderColor = "MemberBorder", Optional FillColor = "MemberFill", Optional Transparency = 0)

    If Width <> 0 And Height <> 0 Then
        'plots a rectangle based on its center
        Rot = Rotation
        If ShapeType <> "Member" Then
            'Rot = -Rotation
        End If
        ExtraAngle = Atn(Round(Width / Height, 2)) * 180 / WorksheetFunction.Pi()
        LengthToCorner = ((Width / 2) ^ 2 + (Height / 2) ^ 2) ^ 0.5
        X = XCoord - LengthToCorner * CustomFunctions.SIND(Rot + ExtraAngle)
        Y = YCoord - LengthToCorner * CustomFunctions.COSD(Rot + ExtraAngle)
        
        If FillColor = "MemberFill" Then
            Set NewShape = Cht.Shapes.AddShape(msoShapeRectangle, XCoord, ChartHeight - YCoord, 2, 2)
            NewShape.Rotation = -Rot
            NewShape.Left = X + ExtraBufferLeft
            NewShape.Top = Y + ExtraBufferTop
            NewShape.Height = Height
            NewShape.Width = Width
        Else
            Set NewShape = Cht.Shapes.AddShape(msoShapeRoundedRectangle, XCoord, ChartHeight - YCoord, 2, 2)
            NewShape.Rotation = -Rot
            NewShape.Left = X + ExtraBufferLeft
            NewShape.Top = Y + ExtraBufferTop
            NewShape.Height = Height
            NewShape.Width = Width
        End If
        
        NewShape.Fill.ForeColor.RGB = ColorMap(FillColor)
        NewShape.Fill.BackColor.RGB = RGB(150, 150, 150)
        NewShape.line.ForeColor.RGB = ColorMap(BorderColor)
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
        
        Set rngLabels = ThisWorkbook.Worksheets("Code").Range("AE45:AH71")
        Set rngAppurts = ThisWorkbook.Worksheets("Discrete Loads").Range("A4:A503")
        
        isInList = (WorksheetFunction.CountIf(rngLabels, CStr(ShapeName)) > 0) And _
                   (WorksheetFunction.CountIf(rngAppurts, CStr(ShapeName)) = 0)
        
'        If isInList _
'           And (InStr(Left(ShapeName, 2), "MP") > 0 Or InStr(Left(ShapeName, 1), "A") > 0 _
'           Or InStr(Left(ShapeName, 1), "B") > 0 Or InStr(Left(ShapeName, 1), "C") > 0 _
'           Or InStr(Left(ShapeName, 1), "D") > 0 Or InStr(Left(ShapeName, 1), "G") > 0 _
'           Or InStr(Left(ShapeName, 1), "M") > 0) Then
'           'IsNumeric(Right(ShapeName, 1)) Then
            
'            'add text box
'            BoxWidth = 40
'            BoxHeight = 20
'            Set TextBox = Cht.Shapes.AddTextbox( _
'                Orientation:=msoTextOrientationHorizontal, _
'                Left:=NewShape.Left - BoxWidth / 2 + Width / 2, _
'                Top:=ChartHeight - BoxHeight, _
'                Width:=BoxWidth, _
'                Height:=BoxHeight)
'
'            TextBox.Name = ShapeName & "_Text"
'            'TextBox.ShapeStyle = msoShapeStylePreset72
'            TextBox.TextFrame.Characters.Text = ShapeName
'            TextBox.AutoShapeType = msoShapeRoundedRectangle
'            TextBox.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
'            TextBox.TextFrame2.VerticalAnchor = msoAnchorMiddle
'            TextBox.TextFrame2.TextRange.Font.Bold = msoTrue
'            TextBox.TextFrame2.TextRange.Font.Size = 9
'        End If
    End If



End Sub

Private Sub AddSemiCircle(ShapeName, XCoord, YCoord, Width, Height, Optional Rot = 0, Optional BorderColor = "MemberBorder", Optional FillColor = "MemberFill", Optional Transparency = 0)
    If Width <> 0 And Height <> 0 Then
        'plots a rectangle based on its center
        ExtraAngle = Atn(Round(Width / Height, 2)) * 180 / WorksheetFunction.Pi()
        LengthToCorner = ((Width / 2) ^ 2 + (Height / 2) ^ 2) ^ 0.5
        X = XCoord - LengthToCorner * CustomFunctions.SIND(Rot + ExtraAngle)
        Y = YCoord - LengthToCorner * CustomFunctions.COSD(Rot + ExtraAngle)
        
        
        Set NewShape = Cht.Shapes.AddShape(msoShapePie, XCoord, ChartHeight - YCoord, 2, 2)
        NewShape.Adjustments(1) = 90
        NewShape.Rotation = -Rot + 90
        NewShape.Left = X + Width + ExtraBufferLeft 'I don't know why, but for the pie shape it rotates around the corner instead of the center like the other shapes I used
        NewShape.Top = Y + ExtraBufferTop
        NewShape.Height = Height
        NewShape.Width = Width / 3
        
        
'        NewShape.Fill.BackColor.RGB = ColorMap(FillColor)
'        NewShape.Fill.TwoColorGradient msoGradientVertical, 1
        NewShape.Fill.ForeColor.RGB = ColorMap(FillColor)
        NewShape.line.ForeColor.RGB = ColorMap(BorderColor)
        NewShape.Fill.Transparency = Transparency
        
        NewShape.Name = ShapeName
    
'        If InStr(Left(ShapeName, 2), "MP") > 0 Then
'        Or InStr(Left(ShapeName, 1), "A") > 0 And IsNumeric(Right(ShapeName, 1)) _
'        Or InStr(Left(ShapeName, 1), "B") > 0 And IsNumeric(Right(ShapeName, 1)) _
'        Or InStr(Left(ShapeName, 1), "C") > 0 And IsNumeric(Right(ShapeName, 1)) _
'        Or InStr(Left(ShapeName, 1), "D") > 0 And IsNumeric(Right(ShapeName, 1)) _
'            'add text box
'            BoxWidth = 60
'            BoxHeight = 20
'            Set TextBox = Cht.Shapes.AddTextbox(msoTextOrientationHorizontal, NewShape.Left - BoxWidth / 2 + Width / 2, ChartHeight - BoxHeight, BoxWidth, BoxHeight)
'            TextBox.Name = ShapeName & "_Text"
'            TextBox.ShapeStyle = msoShapeStylePreset72
'            TextBox.TextFrame.Characters.Text = ShapeName
'            TextBox.AutoShapeType = msoShapeRoundedRectangle
'            TextBox.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
'        End If
    End If
End Sub






