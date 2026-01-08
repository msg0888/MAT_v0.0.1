Attribute VB_Name = "ModelLoadOptions"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private pOverwritePointLoads
Private pOverwriteDistributedLoads
Private pOverwriteAreaLoads
Private pOverwriteSurfaceLoads

Private pOverwriteBLCs
Private pOverwriteLCs

Private pAzimuthList()
Private pAxes

Private pServiceLoadMag
Private pMaintenanceLoadMag

Private pGenerateServiceLoads
Private pGenerateMaintenanceLoads

Private Const MustBeBoolError = 513
Private Const MustBeBoolErrorMsg = "Type Mismatch, Assigned value must be Boolean"

Private Const AxesError = 530
Private Const AxesErrorMsg = "Assigned value must be either ""ZXY"" or ""XYZ"""




Private Sub Class_Initialize()
    pOverwritePointLoads = True
    pOverwriteDistributedLoads = True
    pOverwriteAreaLoads = True
    pOverwriteSurfaceLoads = True
    pOverwriteBLCs = True
    pOverwriteLCs = True
    pGenerateServiceLoads = True
    pGenerateMaintenanceLoads = True
    
    pServiceLoadMag = -250
    pMaintenanceLoadMag = -500
    
    ReDim pAzimuthList(1 To 8)
    pAzimuthList = Array(0, 30, 45, 60, 90, 120, 135, 150)
    
    If Application.Range("Axes").Value <> "" Then
        pAxes = Application.Range("Axes").Value
    Else
        pAxes = "ZXY"
    End If
End Sub


'Setters
Public Property Let OverwritePointLoads(val)
    If VarType(val) = vbBoolean Then
        pOverwritePointLoads = val
    Else
        Err.Raise MustBeBoolError, , MustBeBoolErrorMsg
    End If

End Property
Public Property Let OverwriteDistributedLoads(val)
    If VarType(val) = vbBoolean Then
        pOverwriteDistributedLoads = val
    Else
        Err.Raise MustBeBoolError, , MustBeBoolErrorMsg
    End If

End Property
Public Property Let OverwriteAreaLoads(val)
    If VarType(val) = vbBoolean Then
        pOverwriteAreaLoads = val
    Else
        Err.Raise MustBeBoolError, , MustBeBoolErrorMsg
    End If

End Property
Public Property Let OverwriteSurfaceLoads(val)
    If VarType(val) = vbBoolean Then
        pOverwriteSurfaceLoads = val
    Else
        Err.Raise MustBeBoolError, , MustBeBoolErrorMsg
    End If

End Property
Public Property Let OverwriteBLCs(val)
    If VarType(val) = vbBoolean Then
        pOverwriteBLCs = val
    Else
        Err.Raise MustBeBoolError, , MustBeBoolErrorMsg
    End If

End Property
Public Property Let OverwriteLCs(val)
    If VarType(val) = vbBoolean Then
        pOverwriteLCs = val
    Else
        Err.Raise MustBeBoolError, , MustBeBoolErrorMsg
    End If

End Property
Public Property Let Axes(val)
    If val = "ZXY" Or val = "XYZ" Then
        pAxes = val
    Else
        Err.Raise AxesError, , AxesErrorMsg
    End If
End Property


'Getters
Public Property Get OverwritePointLoads()
    OverwritePointLoads = pOverwritePointLoads
End Property

Public Property Get OverwriteDistributedLoads()
    OverwriteDistributedLoads = pOverwriteDistributedLoads
End Property

Public Property Get OverwriteAreaLoads()
    OverwriteAreaLoads = pOverwriteAreaLoads
End Property

Public Property Get OverwriteSurfaceLoads()
    OverwriteSurfaceLoads = pOverwriteSurfaceLoads
End Property

Public Property Get OverwriteBLCs()
    OverwriteBLCs = pOverwriteBLCs
End Property

Public Property Get OverwriteLCs()
    OverwriteLCs = pOverwriteLCs
End Property

Public Property Get Azimuths()
    Azimuths = pAzimuthList
End Property

Public Property Get Axes()
    Axes = pAxes
End Property

Public Property Get ServiceLoad()
    ServiceLoad = pServiceLoadMag
End Property

Public Property Get MaintenanceLoad()
    MaintenanceLoad = pMaintenanceLoadMag
End Property

Public Property Get GenerateServiceLoads()
    GenerateServiceLoads = pGenerateServiceLoads
End Property
