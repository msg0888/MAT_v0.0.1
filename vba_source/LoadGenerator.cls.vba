Attribute VB_Name = "LoadGenerator"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Data
Private RowsWithData()
Private TempLocationData()
Public GraphicPlacement As New Scripting.Dictionary

Private Const NumPlacementColumns = 5
Private NumSectors
Private CodeRev

'The Final Data Structure that holds the placement data
'   LocationData(,1) -> Quantity (adjusted for number of mounting locations)
'   LocationData(,2) -> Absolute Azimuth of the equipment
'   LocationData(,3) -> Member Label where the load is being placed
'   LocationData(,4) -> Position along member where load is being placed (negative is a % length)
'   LocationData(,5) -> Row Number that the load originated from in the placement table
'                               (used for reference to pull forces from the calcs table)
Private LocationData()

Private SelectedModel As RISAModel


Private AlphaSector
Private BetaSector
Private GammaSector
Private DeltaSector

Private NameIndex
Private ConditionIndex
Private WeightIndex
Private IceWeightIndex
Private NormalForceIndex
Private TransverseForceIndex
Private IceNormalForceIndex
Private IceTransverseForceIndex
Private SeismicIndex

Private QuantityIndex
Private AzimuthIndex
Private Member1PrefixIndex
Private Member1NumberIndex
Private Member2PrefixIndex
Private Member2NumberIndex
Private Location1Index
Private Location2Index

'Private BLCs() As BasicLoadCase
'Private LCs() As LoadCombination
'
'Private PointLoads() As PointLoad
'Private DistributedLoads() As DistributedLoad
'Private AreaLoads() As AreaLoad
'Private SurfaceLoads() As SurfaceLoad

Private SelfWeightsBLCNum
Private WindLoadsBLCNum
Private IceWeightsBLCNum
Private IceWindLoadsBLCNum
Private SeismicLoadsBLCNum
Private LiveLoadsBLCNum
Private MaintenanceLoadsBLCNum



Private Sub Class_Initialize()
    'determine the number of sectors that were modeled that need loading
    NumSectors = Application.Range("QtySectors").Value
End Sub



Public Sub GatherLoadData(CodeRevision)
    CodeRev = CodeRevision
    Dim CalcSheetTable As ListObject
    Dim PlacementSheet As Worksheet
    
    Set CalcSheetTable = ThisWorkbook.Worksheets("Filter").ListObjects("tblAPPURT")
    Set DimSheetTable = ThisWorkbook.Worksheets("Filter").ListObjects("tblDimensions")
    
    'Collect actual Data
    Data = CalcSheetTable.DataBodyRange.Value
    AlphaSector = Application.Transpose(shFilter.ListObjects("tblAlpha").DataBodyRange.Value)
    BetaSector = Application.Transpose(shFilter.ListObjects("tblBeta").DataBodyRange.Value)
    GammaSector = Application.Transpose(shFilter.ListObjects("tblGamma").DataBodyRange.Value)
    DeltaSector = Application.Transpose(shFilter.ListObjects("tblDelta").DataBodyRange.Value)
    
    
    'Find Column Numbers for relevant data in tables
    With CalcSheetTable
        NameIndex = .ListColumns("Name").Index
        ConditionIndex = .ListColumns("Condition").Index
    End With
    
    With DimSheetTable
        HeightIndex = .ListColumns("Height / Diameter (in)").Index
        WeightIndex = .ListColumns("Weight (lbs)").Index
    End With
    
    With shFilter.ListObjects("tblAlpha")
        QuantityIndex = .ListColumns("Quantity").Index
        AzimuthIndex = .ListColumns("Azimuth (°)").Index
        Member1PrefixIndex = .ListColumns("Position Top").Index
        Member2PrefixIndex = .ListColumns("Position Btm").Index
        Location1Index = .ListColumns("Location Top (in)").Index
        Location2Index = .ListColumns("Location Btm (in)").Index
    End With
    
End Sub

Public Sub DetermineLoadPlacements()


    'Determine which Rows have appurtenances here
    i = 0
    For j = LBound(Data) To UBound(Data)
        If Data(j, NameIndex) <> "" Then
            i = i + 1
            ReDim Preserve RowsWithData(1 To i)
            RowsWithData(i) = j
        End If
    Next j


    'Reformat Data into streamlined form for later use
    ReDim TempLocationData(1 To NumPlacementColumns, 1 To 1)
    CompressData AlphaSector
    If NumSectors > 1 Then
        CompressData BetaSector, UBound(TempLocationData, 2)
    End If
    If NumSectors > 2 Then
        CompressData GammaSector, UBound(TempLocationData, 2)
    End If
    If NumSectors > 3 Then
        CompressData DeltaSector, UBound(TempLocationData, 2)
    End If

    ReDim LocationData(1 To UBound(TempLocationData, 2), 1 To NumPlacementColumns)
    LocationData = Application.Transpose(TempLocationData)


End Sub

Private Sub CompressData(ByRef SectorData, Optional JStart = 0)
    j = JStart
    For Each row In RowsWithData
        If SectorData(QuantityIndex, row) <> 0 Then
            If SectorData(Location1Index, row) <> "" Then
                j = j + 1
                ReDim Preserve TempLocationData(1 To NumPlacementColumns, 1 To j)
                TempLocationData(1, j) = SectorData(QuantityIndex, row)
                TempLocationData(2, j) = SectorData(AzimuthIndex, row)
                TempLocationData(3, j) = CStr(SectorData(Member1PrefixIndex, row)) '& CStr(SectorData(Member1NumberIndex, Row))
                TempLocationData(4, j) = SectorData(Location1Index, row)
                TempLocationData(5, j) = row
            End If
            If SectorData(Location2Index, row) <> "" Then
                j = j + 1
                ReDim Preserve TempLocationData(1 To NumPlacementColumns, 1 To j)
                
                'if there is a value in Location 2, divide the quantities for both locations by 2
                '                                   since it is mounted in two places
                'Location 1 IS REQUIRED TO HAVE a value, if there is one in location 2
                TempLocationData(1, j - 1) = TempLocationData(1, j - 1) / 2
                TempLocationData(1, j) = SectorData(QuantityIndex, row) / 2
                TempLocationData(2, j) = SectorData(AzimuthIndex, row)
                TempLocationData(3, j) = CStr(SectorData(Member2PrefixIndex, row)) '& CStr(SectorData(Member2NumberIndex, Row))
                TempLocationData(4, j) = SectorData(Location2Index, row)
                TempLocationData(5, j) = row
            End If
        End If
        
        MountMember = CStr(SectorData(Member1PrefixIndex, row)) '& CStr(SectorData(Member1NumberIndex, Row))
        EquipmentName = Data(row, NameIndex) & "_" & row
        
        'Get data for Graphics
        If Not GraphicPlacement.Exists(MountMember) Then
            GraphicPlacement.Add MountMember, New Scripting.Dictionary
        End If
        
        If Not GraphicPlacement(MountMember).Exists(EquipmentName) Then
            GraphicPlacement(MountMember).Add EquipmentName, New Scripting.Dictionary
            GraphicPlacement(MountMember)(EquipmentName).Add "Name", Data(row, NameIndex)
            GraphicPlacement(MountMember)(EquipmentName).Add "Quantity", SectorData(QuantityIndex, row)
            GraphicPlacement(MountMember)(EquipmentName).Add "Azimuth (º)", SectorData(AzimuthIndex, row)
            GraphicPlacement(MountMember)(EquipmentName).Add "Location Top (in)", SectorData(Location1Index, row)
            GraphicPlacement(MountMember)(EquipmentName).Add "Location Btm (in)", SectorData(Location2Index, row)
        Else
            GraphicPlacement(MountMember)(EquipmentName)("Quantity") = GraphicPlacement(MountMember)(EquipmentName)("Quantity") + SectorData(QuantityIndex, row)
        End If
        
        
'        If Not GraphicPlacement(MountMember).Exists(EquipmentName) Then
'            GraphicPlacement(MountMember).Add EquipmentName, New Scripting.Dictionary
'            GraphicPlacement(MountMember)(EquipmentName).Add "Name", Data(Row, NameIndex)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Quantity", SectorData(QuantityIndex_Beta, Row)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Azimuth (º)", SectorData(AzimuthIndex_Beta, Row)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Location Top (in)", SectorData(Location1Index_Beta, Row)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Location Btm (in)", SectorData(Location2Index_Beta, Row)
'        Else
'            GraphicPlacement(MountMember)(EquipmentName)("Quantity") = GraphicPlacement(MountMember)(EquipmentName)("Quantity") + SectorData(QuantityIndex_Beta, Row)
'        End If
'
'
'        If Not GraphicPlacement(MountMember).Exists(EquipmentName) Then
'            GraphicPlacement(MountMember).Add EquipmentName, New Scripting.Dictionary
'            GraphicPlacement(MountMember)(EquipmentName).Add "Name", Data(Row, NameIndex)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Quantity", SectorData(QuantityIndex_Gamma, Row)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Azimuth (º)", SectorData(AzimuthIndex_Gamma, Row)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Location Top (in)", SectorData(Location1Index_Gamma, Row)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Location Btm (in)", SectorData(Location2Index_Gamma, Row)
'        Else
'            GraphicPlacement(MountMember)(EquipmentName)("Quantity") = GraphicPlacement(MountMember)(EquipmentName)("Quantity") + SectorData(QuantityIndex_Gamma, Row)
'        End If
'
'
'        If Not GraphicPlacement(MountMember).Exists(EquipmentName) Then
'            GraphicPlacement(MountMember).Add EquipmentName, New Scripting.Dictionary
'            GraphicPlacement(MountMember)(EquipmentName).Add "Name", Data(Row, NameIndex)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Quantity", SectorData(QuantityIndex_Delta, Row)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Azimuth (º)", SectorData(AzimuthIndex_Delta, Row)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Location Top (in)", SectorData(Location1Index_Delta, Row)
'            GraphicPlacement(MountMember)(EquipmentName).Add "Location Btm (in)", SectorData(Location2Index_Delta, Row)
'        Else
'            GraphicPlacement(MountMember)(EquipmentName)("Quantity") = GraphicPlacement(MountMember)(EquipmentName)("Quantity") + SectorData(QuantityIndex_Delta, Row)
'        End If
    
    Next row
End Sub
