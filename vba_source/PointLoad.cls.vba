Attribute VB_Name = "PointLoad"
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
    '88 = Global X
    '89 = Global Y
    '90 = Global Z
    '120 = Local x
    '121 = Local y
    '122 = Local z
    '125 = Local Mx
    '126 = Local My
    '127 Local Mz

Private Magnitude As Double
Private Location As Double

Private RISAFloor As Long
Private BIMID As Long
Private WallHeight As Long
Private MemberType As Long
    '2 = Member
    '7 Wall Panel

Private ParamI As Long
Private ParamJ As Long

Private BLCNumber As Long
' This is not a value required by RISA, this is mostly for organization
' RISA applies a BLC to a load based off the number of loads specified and the order of the loads in the file (which is a stupid way to do it)
    

Private Sub Class_Initialize()
    ' creates an instance of the object and assigns it default values
    MemberNumber = 0
    LoadDirection = 0
    Magnitude = 0#
    Location = 0#
    RISAFloor = 0
    BIMID = -1
    WallHeight = 0
    MemberType = 2
    ParamI = 0
    ParamJ = 0

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
    
    MemberNumber = CLng(Data(0))
    LoadDirection = CLng(Data(1))
    Magnitude = CDbl(Data(2))
    Location = CDbl(Data(3))
    RISAFloor = CLng(Data(4))
    BIMID = CLng(Data(5))
    WallHeight = CLng(Data(6))
    MemberType = CLng(Data(7))
    ParamI = CLng(Data(8))
    ParamJ = CLng(Data(9))
    
End Sub

Public Sub Create(MemNum, LoadDir, Mag, Loc, BLC, Optional MemType = 2, Optional Height = 0)
    ' method to assign values to the created object
    MemberNumber = MemNum
    LoadDirection = Globals.DirMap(LoadDir)
    Magnitude = Mag
    Location = Loc
    MemberType = MemType
    WallHeight = Height
    
    BLCNumber = BLC
End Sub

Public Function StrRep()
    StrRep = Join(Array(MemberNumber, LoadDirection, Magnitude, Location, RISAFloor, BIMID, WallHeight, MemberType, ParamI, ParamJ), " ") & ";"
End Function
