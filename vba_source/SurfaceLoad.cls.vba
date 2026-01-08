Attribute VB_Name = "SurfaceLoad"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String

Private PlateNumber As Long
Private LoadDirection As Long
Private LoadMagnitude As Long

Private BIMID As Long
Private ElementType As Long

Private TopMagnitude As Double
Private BottomMagnitude As Double
Private Height As Double

Private DiaphragmI As Long
Private DiaphragmJ As Long
Private DiaphragmK As Double
Private DiaphragmL As Double
Private DiaphragmM As Double
Private DiaphragmN As Double
Private DiaphragmO As Double


Private Sub Class_Initialize()
    ' Set RISA Defaults
    PlateNumber = 0
    LoadMagnitude = 0
    LoadDirection = 0
    
    BIMID = -1
    ElementType = 3
    
    TopMagnitude = 0#
    BottomMagnitude = 0#
    Height = 0#
    
    DiaphragmI = 0
    DiaphragmJ = 0
    DiaphragmK = 0#
    DiaphragmL = 0#
    DiaphragmM = 0#
    DiaphragmN = 0#
    DiaphragmO = 0#
    
    Set regex = Nothing
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
    
    PlateNumber = CLng(Data(0))
    LoadMagnitude = CLng(Data(1))
    LoadDirection = CLng(Data(2))
    BIMID = CLng(Data(3))
    ElementType = CLng(Data(4))
    TopMagnitude = CDbl(Data(5))
    BottomMagnitude = CDbl(Data(6))
    Height = CDbl(Data(7))
    DiaphragmI = CLng(Data(8))
    DiaphragmJ = CLng(Data(9))
    DiaphragmK = CDbl(Data(10))
    DiaphragmL = CDbl(Data(11))
    DiaphragmM = CDbl(Data(12))
    DiaphragmN = CDbl(Data(13))
    DiaphragmO = CDbl(Data(14))
    
    
End Sub
Public Sub Create(PlateNum, Magnitude, LoadDir)
    ' Create a New Surface Load
    PlateNumber = PlateNum
    LoadMagnitude = Magnitude
    LoadDirection = Globals.LoadCatMap(LoadDir)
End Sub

Public Function StrRep()
    Temp = Array(PlateNumber, LoadDirection, LoadMagnitude, BIMID, ElementType, TopMagnitude, _
                    BottomMagnitude, Height, DiaphragmI, DiaphragmJ, DiaphragmK, "|", _
                    DiaphragmL, DiaphragmM, DiaphragmN, DiaphragmO)
    StrRep = Join(Temp, " ") & ";"
End Function
