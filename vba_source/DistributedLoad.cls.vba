Attribute VB_Name = "DistributedLoad"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String


Private MemberNumber As Long

Private LoadDirection  As Long
    '10 = Surface Load Global X
    '11 = Surface Load Global Y
    '12 = Surface Load Global Z
    '88 = Global X
    '89 = Global Y
    '90 = Global Z
    '120 = Local x
    '121 = Local y
    '122 = Local z
    '76 = Projected Global X
    '86 = Projected Global Y
    '72 = Projected Global Z

Private StartMagnitude As Double
Private EndMagnitude As Double

Private StartLocation As Double
Private EndLocation As Double
' -Number is % along length
' +Number is location in base units

Private RISAFloor As Long
Private BIMID As Long
Private StartEle As Double
Private EndEle As Double
Private ElementType As Long
' These will not change

Private ParamL As Double
Private ParamM As Double
Private ParamN As Double
' These are not noted in the risa provided legend
' I have no idea what they do so i am just using the defaults I saw when i open the text file


Private Sub Class_Initialize()
    
    ' creates an instance of the object and assigns it default values
    MemberNumber = 0
    LoadDirection = 0
    StartMagnitude = 0#
    EndMagnitude = 0#
    StartLocation = 0#
    EndLocation = 0#
    RISAFloor = 0
    BIMID = -1
    StartEle = 0#
    EndEle = 0#
    ElementType = 2
    ParamL = 0#
    ParamM = 0#
    ParamN = 0#
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
    
    MemberNumber = CLng(Data(0))
    LoadDirection = CLng(Data(1))
    StartMagnitude = CDbl(Data(2))
    EndMagnitude = CDbl(Data(3))
    StartLocation = CDbl(Data(4))
    EndLocation = CDbl(Data(5))
    RISAFloor = CLng(Data(6))
    BIMID = CLng(Data(7))
    StartEle = CDbl(Data(8))
    EndEle = CDbl(Data(9))
    ElementType = CLng(Data(10))
    ParamL = CDbl(Data(11))
    ParamM = CDbl(Data(12))
    ParamN = CDbl(Data(13))
    
    Set regex = Nothing
End Sub

Public Sub Create(MemNum, LoadDir, StartMag, EndMag, StartLoc, EndLoc)
    ' method to assign values to the created object
    MemberNumber = MemNum
    LoadDirection = Globals.DirMap(LoadDir)
    StartMagnitude = StartMag
    EndMagnitude = EndMag
    StartLocation = StartLoc
    EndLocation = EndLoc
End Sub

Public Function StrRep()
    StrRep = Join(Array(MemberNumber, LoadDirection, StartMagnitude, EndMagnitude, StartLocation, EndLocation, RISAFloor, BIMID, StartEle, EndEle, ElementType, ParamL, ParamM, ParamN), " ") & ";"
End Function
