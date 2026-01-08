Attribute VB_Name = "Node"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String

Private NodeLabel As String

Private pX As Double
Private pY As Double
Private pZ As Double

Private NodeTemperature As Double
Private SelectFlag As Long
Private DiaphragmConnectivity As Long
Private RISAFloorInteraction As Long
Private BIMID As Long
Private DiaphragmOffset As Long
Private WallFlag As Long

Dim Num


Private Sub Class_Initialize()
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
    
    NodeLabel = Data(0)
    pX = CDbl(Data(1))
    pY = CDbl(Data(2))
    pZ = CDbl(Data(3))
    NodeTemperature = CDbl(Data(4))
    SelectFlag = CLng(Data(5))
    DiaphragmConnectivity = CLng(Data(6))
    RISAFloorInteraction = CLng(Data(7))
    BIMID = CLng(Data(8))
    DiaphragmOffset = CLng(Data(9))
    WallFlag = CLng(Data(10))
    
    Set regex = Nothing
End Sub


Public Function StrRep()
    StrRep = Join(Array(NodeLabel, pX, pY, pZ, NodeTempurature, SelectFlag, DiaphragmConnectivity, RISAFloorInteraction, BIMID, DiaphragmOffset, WallFlag), " ") & ";"
End Function

Public Property Get Name()
    Name = Trim(Replace(NodeLabel, """", ""))
End Property

Public Property Get Number()
    Number = Num
End Property
Public Property Get X()
    X = pX
End Property
Public Property Get Y()
    Y = pY
End Property
Public Property Get Z()
    Z = pZ
End Property


Public Property Let Number(dNum)
    Num = dNum
End Property

