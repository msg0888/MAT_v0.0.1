Attribute VB_Name = "BasicLoadCase"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String

Private BLCNum As Long

Private BLCName As String

Private JointLoads As Long
Private PointLoads As Long
Private pDistributedLoads As Long
Private DistributedLoads As Long
Private AreaLoads As Long
Private SurfaceLoads As Long

Private LoadCat As Long

Private XGrav As Double
Private YGrav As Double
Private ZGrav As Double

Private Unknown As Long


Private Const BoundsError = 550
Private Const BoundsErrorMsg = "Illegal Value assigned"


Public Property Let Num(Number)
    BLCNum = Number
End Property


Private Sub Class_Initialize()

    BLCNum = 0
    
    BLCName = CustomFunctions.BufferString("", 32)
    
    NumJointLoads = 0
    NumPointLoads = 0
    pDistributedLoads = 0
    NumDistributedLoads = 0
    NumAreaLoads = 0
    NumSurfaceLoads = 0
    
    LoadCategory = 0
    
    XGrav = 0#
    YGrav = 0#
    ZGrav = 0#

    Unknown = -1
End Sub

Public Sub Define(dRaw)
    Raw = dRaw
    
    Process
End Sub

Private Sub Process()
    Dim regex As New RegExp
    With regex
        .Global = True
        .IgnoreCase = False
        .MultiLine = False
    End With

    regex.pattern = """.{32}""|[^ \t\r\n\f;]+"
    If regex.Test(Raw) Then
        Set Data = regex.Execute(Raw)
    Else
        Exit Sub
    End If
    
    BLCNum = CLng(Data(0))
    BLCName = Data(1)
    NumJointLoads = CLng(Data(2))
    NumPointLoads = CLng(Data(3))
    pDistributedLoads = CLng(Data(4))
    NumDistributedLoads = CLng(Data(5))
    NumAreaLoads = CLng(Data(6))
    NumSurfaceLoads = CLng(Data(7))
    LoadCategory = CLng(Data(8))
    XGrav = CDbl(Data(9))
    YGrav = CDbl(Data(10))
    ZGrav = CDbl(Data(11))
    Unknown = CLng(Data(12))
        
    Set regex = Nothing
End Sub

Public Sub Create(Number, Name, Optional LoadC = "None", Optional X = 0, Optional Y = 0, Optional Z = 0, Optional numJoint = 0, Optional NumPoint = 0, Optional NumDist = 0, Optional NumArea = 0, Optional NumSurface = 0)
    BLCNum = Number
    BLCName = CustomFunctions.BufferString(Name, 32)
    LoadCat = Globals.LoadCatMap(LoadC)
    
    XGrav = X
    YGrav = Y
    ZGrav = Z
    
    NumJointLoads = numJoint
    NumPointLoads = NumPoint
    NumDistributedLoads = NumDist
    NumAreaLoads = NumArea
    NumSurfaceLoads = NumSurface
End Sub

Public Property Let NumJointLoads(Num)
    If Num >= 0 Then
        JointLoads = Num
    Else
        Err.Raise BoundsError, , BoundsErrorMsg
    End If
End Property
Public Property Let NumPointLoads(Num)
    If Num >= 0 Then
        PointLoads = Num
    Else
        Err.Raise BoundsError, , BoundsErrorMsg
    End If
End Property
Public Property Let NumDistributedLoads(Num)
    If Num >= 0 Then
        DistributedLoads = Num
    Else
        Err.Raise BoundsError, , BoundsErrorMsg
    End If
End Property
Public Property Let NumAreaLoads(Num)
    If Num >= 0 Then
        AreaLoads = Num
    Else
        Err.Raise BoundsError, , BoundsErrorMsg
    End If
End Property
Public Property Let NumSurfaceLoads(Num)
    If Num >= 0 Then
        SurfaceLoads = Num
    Else
        Err.Raise BoundsError, , BoundsErrorMsg
    End If
End Property

Public Property Let XGravity(val)
    XGrav = val
End Property
Public Property Let YGravity(val)
    YGrav = val
End Property
Public Property Let ZGravity(val)
    ZGrav = val
End Property
Public Property Let Category(val)
    LoadCat = LoadCatMap(val)
End Property


Public Property Get Num()
    Num = BLCNum
End Property

Public Function StrRep()
    Temp = Array(BLCNum, """" & CustomFunctions.BufferString(BLCName, 32) & """", JointLoads, PointLoads, pDistributedLoads, DistributedLoads, _
                    AreaLoads, SurfaceLoads, LoadCat, XGrav, YGrav, ZGrav, Unknown)
                    
    StrRep = Join(Temp, " ") & ";"
End Function

