Attribute VB_Name = "AreaLoad"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String

Private JointA As Long
Private JointB As Long
Private JointC As Long
Private JointD As Long

Private LoadDirection As Long
    '88 = Global X
    '89 = Global Y
    '90 = Global Z
    '76 = Projected Global X
    '86 = Projected Global Y
    '72 = Projected Global Z


Private DistributioinDirection As Long
    '0 = Two Way
    '1 = A-B
    '2 = B-C
    '3 = C-D
    '4 = A-D
    '5 = A-C
    '6 = B-D

Private Magnitude As Double

Private BIMID As Long
Private RISAFloor As Long


Private Sub Class_Initialize()
    JointA = 0
    JointB = 0
    JointC = 0
    JointD = 0
    LoadDirection = 0
    DistributedDirection = 0
    Magnitude = 0#
    BIMID = -1
    RISAFloor = 0
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
    
    JointA = CLng(Data(0))
    JointB = CLng(Data(1))
    JointC = CLng(Data(2))
    JointD = CLng(Data(3))
    LoadDirection = CLng(Data(4))
    DistributionDirection = CLng(Data(5))
    Magnitude = CDbl(Data(6))
    BIMID = CLng(Data(7))
    RISAFloor = CLng(Data(8))
    
    Set regex = Nothing
End Sub


Public Sub Create(a, b, c, d, LoadDir, DistDir, Mag)
    JointA = a
    JointB = b
    JointC = c
    JointD = d
    LoadDirection = Globals.LoadCatMap(LoadDir)
    DistributionDirection = DistDir
    Magnitude = Mag
End Sub

Public Function StrRep()
    StrRep = Join(Array(JointA, JointB, JointC, JointD, LoadDirection, DistributionDirection, Magnitude, BIMID, RISAFloor), " ") & ";"
End Function


