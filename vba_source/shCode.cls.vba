Attribute VB_Name = "shCode"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton1, 23, 1, MSForms, CommandButton"

Private Sub LinkModel_Click()
    
    Dim fd As Office.FileDialog
    ' Make the dialog picker
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        ' Only allow the user to select one file
        .AllowMultiSelect = False
        
        ' Set the title of the dialog box.
        .title = "Choose Model"
        
        ' Set the initial path to current path
        .InitialFileName = ThisWorkbook.path
        
        ' Clear any existing filters, and add our own
        .Filters.Clear
        .Filters.Add "RISA Model Files", "*.r3d"
        .Filters.Add "All Files", "*.*"
        
        
        If .Show Then
            Application.Range("ModelPath").Value = .SelectedItems(1) 'return the chosen file path name
'            If InStr(Application.Range("MountType").Value, "Platform") <> 0 Then
'                Application.Range("CP").Value = InputBox("Please enter the label for the center point of the model")
'            Else
'                Application.Range("CP").Value = ""
'            End If
            MsgBox ("Model Successfully Chosen.")
        Else
            MsgBox ("Action canceled by user.")
        End If
    End With
End Sub

Private Sub CommandButton1_Click()
    LoadModel False
End Sub

Private Sub LoadDefault_Click()
    LoadModel False
End Sub

'Private Sub LoadDefault_Click()
'    LoadModel False
'End Sub

Private Sub LoadCustom_Click()
    LoadModel True
End Sub

Private Sub LoadModel(LoadWithCustom)
    
    Set Globals.LoadCatMap = Nothing
    Set Globals.DirMap = Nothing
    Set Globals.MaterialTypeMap = Nothing
    Set Globals.options = Nothing
    
    Dim ModelData As RISAModel
    Dim Elevation As ViewElevation
    Dim Plan As ViewPlan
    
    'Determine Path where the model to be loaded is located
    ModelPath = Application.Range("ModelPath").Value
    Do While ModelPath = ""
        LinkModel_Click
        ModelPath = Application.Range("ModelPath").Value
    Loop
        
    
    'Show the menu for the user to choose additional loading options, if necessary
    If LoadWithCustom Then
    
    End If

    'Parse the model file
    Set ModelData = New RISAModel
    On Error GoTo Handle:
    ModelData.Parse ModelPath
    
    'Add Members to Graphic
    Set Elevation = New ViewElevation
    Set Plan = New ViewPlan
    On Error Resume Next

        Set CP = ModelData.NodesByName(shCode.Range("CenterPoint").Value)

        If Err.Number <> 0 And Application.Range("QtySectors").Value > 1 Then
            MsgBox ("The Node Label """ & shCode.Range("CenterPoint").Value & """ does not exist in the model, please select a valid Center Point Label.")
            End
        ElseIf Err.Number <> 0 Then
            Set CP = New Node
        End If
    On Error GoTo 0

    
    Elevation.SetGraphBoundaries ModelData.ModelMinX, ModelData.ModelMinY, ModelData.ModelMinZ, ModelData.ModelMaxX, ModelData.ModelMaxY, ModelData.ModelMaxZ, CP
    Plan.SetGraphBoundaries ModelData.ModelMinX, ModelData.ModelMinY, ModelData.ModelMinZ, ModelData.ModelMaxX, ModelData.ModelMaxY, ModelData.ModelMaxZ
    
    Dim m As Member
    For i = LBound(ModelData.Members) To UBound(ModelData.Members)
        Debug.Print ModelData.Nodes(ModelData.Members(i).iNode).Name
        Elevation.AddMember ModelData.Members(i), ModelData.Nodes(ModelData.Members(i).iNode), ModelData.Nodes(ModelData.Members(i).jNode)
        Plan.AddMember ModelData.Members(i), ModelData.Nodes(ModelData.Members(i).iNode), ModelData.Nodes(ModelData.Members(i).jNode)
    Next i
    
    
    'Create a LoadGenerator to create the loads for the chosen model
    Dim Generator As New LoadGenerator
    
    CodeRev = Right(Application.Range("TIA").Value, 1)
    Generator.GatherLoadData CodeRev
    Generator.DetermineLoadPlacements

    Set AppTableObj = ThisWorkbook.Sheets("Filter").ListObjects("tblAPPURT")
    AppTable = AppTableObj.DataBodyRange.Value
    ConditionTable = AppTableObj.DataBodyRange.Value
    NameColumnData = AppTableObj.ListColumns("Name").Range.Value
    ConditionColumnData = AppTableObj.ListColumns("Condition").Range.Value
    
    Set AppTableDim = ThisWorkbook.Sheets("Filter").ListObjects("tblDimensions")
    DimTable = AppTableDim.DataBodyRange.Value
    HeightColumnData = AppTableDim.ListColumns("Height / Diameter (in)").Range.Value


    HeightIndex = AppTableDim.ListColumns("Height / Diameter (in)").Index
    WidthIndex = AppTableDim.ListColumns("Width (in)").Index
    DepthIndex = AppTableDim.ListColumns("Depth (in)").Index
    ConditionIndex = AppTableObj.ListColumns("Condition").Index
    TypeIndex = AppTableObj.ListColumns("Type").Index
    
    For Each MemName In Generator.GraphicPlacement.Keys
        For Each AppName In Generator.GraphicPlacement(MemName).Keys
            With Generator
                Condition = "Existing"
                Position = .GraphicPlacement(MemName)(AppName)("Location Top (in)")
                AppType = "TME"
                Position = .GraphicPlacement(MemName)(AppName)("Location Top (in)")
                AppType = "Round"
                Position = .GraphicPlacement(MemName)(AppName)("Location Top (in)")
                AppType = "Antenna"

                Condition = "Proposed"
                Position = .GraphicPlacement(MemName)(AppName)("Location Top (in)")
                AppType = "TME"
                Position = .GraphicPlacement(MemName)(AppName)("Location Top (in)")
                AppType = "Round"
                Position = .GraphicPlacement(MemName)(AppName)("Location Top (in)")
                AppType = "Antenna"

                
                'match the antenna name to the dimensions in the appurtenance table
                For i = LBound(NameColumnData) To UBound(NameColumnData)
                    If NameColumnData(i, 1) = .GraphicPlacement(MemName)(AppName)("Name") Then
                        rowNum = i - 1
                    End If
                Next i
                
                'Antennas
                If AppTable(rowNum, TypeIndex) <> "N/A" Then
                    AppType = AppTable(rowNum, TypeIndex)
                End If
                
                If ConditionTable(rowNum, ConditionIndex) <> "" Then
                    Condition = ConditionTable(rowNum, ConditionIndex)
                End If

                Elevation.AddEquipment AppName, MemName, Position, _
                    DimTable(rowNum, HeightIndex), DimTable(rowNum, WidthIndex), DimTable(rowNum, DepthIndex), AppType, _
                    .GraphicPlacement(MemName)(AppName)("Quantity"), Condition

                Plan.AddEquipment AppName, MemName, _
                    DimTable(rowNum, HeightIndex), DimTable(rowNum, WidthIndex), DimTable(rowNum, DepthIndex), AppType, _
                    .GraphicPlacement(MemName)(AppName)("Quantity"), Condition, .GraphicPlacement(MemName)(AppName)("Azimuth (º)")
                
            End With
        Next AppName
    Next MemName
    
    SuccessMessage = "Placement Diagrams Successfully Created!" & vbNewLine & vbNewLine & "Please verify that appurtenances are correctly positioned. The appurtenances may not be generate correctly if the appurtenance is installed on a non-mount pipe member and may require manual adjustment."

    MsgBox SuccessMessage, vbInformation, "Placement Diagrams Successfully Created!"

Handle:
If Err.Number > 0 Then MsgBox "There is an error with a custom member label. Please update the RISA-3D model and try again. Importing the updated RISA-3D model is not necessary.", vbCritical + vbOKOnly, "RISA-3D Member Label Error!"

End Sub






Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    
    
    
    ' Hide Sheets
    Dim Sheet As Worksheet
    On Error Resume Next
    If Not Intersect(Target, Application.Range("TIA")) Is Nothing Then
        For Each Sheet In Application.Sheets
        
'            If InStr(Sheet.Name, "- G") > 0 Or InStr(Sheet.Name, "- H") > 0 Then
'                Sheet.Visible = xlSheetHidden
'            End If
            
'            If InStr(Sheet.Name, "- " & Application.Range("TIA").Value) > 0 Then
'                Sheet.Visible = xlSheetVisible
'            End If
        Next Sheet
    End If
    
    
    'Hide Inputs
    Application.Range("BaseElevation").EntireRow.Hidden = Application.Range("TIA").Value = "ANSI/TIA-222-G"
    Application.Range("TopoData").EntireRow.Hidden = Application.Range("Kzt").Value = 1
    Application.Range("RooftopData").EntireRow.Hidden = Application.Range("TowerType") <> "Building"
    Inputs.Shapes("TopoImage").Visible = Application.Range("Kzt").Value <> 1
    Inputs.Shapes("RooftopImage").Visible = Application.Range("TowerType") = "Building"
    Inputs.Range("TankType").EntireRow.Hidden = Inputs.Range("TowerType").Value <> "Water Tank"
    Inputs.Range("CenterPoint").EntireRow.Hidden = Inputs.Range("QtySectors") = 1
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
