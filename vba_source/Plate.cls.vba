Attribute VB_Name = "Plate"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String

Private PlateLabel As String

Private AInput As Long

Private NodeA As Long
Private NodeB As Long
Private NodeC As Long
Private NodeD As Long

Private MaterialType  As Long
Private MaterialOffset As Long
Private PlateThickness As Double
Private LegacyI As Double
Private LegacyJ As Double
Private LegacyK As Double
Private LegacyL As Double
Private LegacyM As Double
Private LegacyN As Double
 
Private SelectionStatus As Long
Private PInput As Long
Private PlaneStressPlate As Long
Private RInput As Long
Private SInput As Long

Private ARelease As Long
Private BRelease As Long
Private CRelease As Long
Private DRelease As Long

Private XInput As Long
Private YInput As Long
Private ZInput As Long
Private AAInput As Long
Private ABInput As Double
Private ACInput As Long


Private Sub Class_Initialize()
    ' Set RISA Defaults
    PlateLabel = CustomFunctions.BufferString("", 32)
    
    AInput = 51
    
    NodeA = 0
    NodeB = 0
    NodeC = 0
    NodeD = 0
    
    MaterialType = 0
    MaterialOffset = 5
    
    LegacyI = -1#
    LegacyJ = -1#
    LegacyK = -1#
    LegacyL = -1#
    LegacyM = -1#
    LegacyN = -1#
    
    SelectionStatus = 65535
    PInput = 0
    PlaneStressPlate = 0
    RInput = -1
    SInput = -1
    
    ARelease = 0
    BRelease = 0
    CRelease = 0
    DRelease = 0
    
    XInput = 0
    YInput = -1
    ZInput = -1
    AAInput = -1
    ABInput = 0#
    ACInput = -1
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
    
    PlateLabel = Data(0)
    AInput = CLng(Data(1))
    NodeA = CLng(Data(2))
    NodeB = CLng(Data(3))
    NodeC = CLng(Data(4))
    NodeD = CLng(Data(5))
    MaterialType = CLng(Data(6))
    MaterialOffset = CLng(Data(7))
    LegacyI = CDbl(Data(8))
    LegacyJ = CDbl(Data(9))
    LegacyK = CDbl(Data(10))
    LegacyL = CDbl(Data(11))
    LegacyM = CDbl(Data(12))
    LegacyN = CDbl(Data(13))
    SelectionStatus = CLng(Data(14))
    PInput = CLng(Data(15))
    PlaneStressPlate = CLng(Data(16))
    RInput = CLng(Data(17))
    SInput = CLng(Data(18))
    ARelease = CLng(Data(19))
    BRelease = CLng(Data(20))
    CRelease = CLng(Data(21))
    DRelease = CLng(Data(22))
    XInput = CLng(Data(23))
    YInput = CLng(Data(24))
    ZInput = CLng(Data(25))
    AAInput = CLng(Data(26))
    ABInput = CDbl(Data(27))
    ACInput = CLng(Data(28))
    
    Set regex = Nothing
End Sub



Public Sub Create(Name, ANum, BNum, CNum, dNum, thickness)
    ' Create a 4-sided Plate
    Label = Name
    NodeA = ANum
    NodeB = BNum
    NodeC = CNum
    NodeD = dNum
    PlateThickness = thickness
    
    ' Find the vector normal to the face of the plate
'    FindNormalVector
    
End Sub


Public Function StrRep()
' Return the string representation of the Plate (in RISA format)
    StrRep = Raw
'    StrRep = """" & PlateLabel & " " & _
'             CStr(AInput) & " " & _
'             CStr(NodeA) & " " & CStr(NodeB) & " " & CStr(NodeC) & " " & CStr(NodeD) & " " & _
'             CStr(MaterialType) & " " & _
'             CStr(MaterialOffset) & " " & _
'             CStr(PlateThickness) & " " & _
'             CStr(LegacyI) & " " & _
'             CStr(LegacyJ) & " " & _
'             CStr(LegacyK) & " " & _
'             CStr(LegacyL) & " " & _
'             CStr(LegacyM) & " " & _
'             CStr(LegacyN) & " " & _
'             CStr(SelectionStatus) & " " & _
'             CStr(PInput) & " " & _
'             CStr(PlaneStressPlate) & " " & _
'             CStr(RInput) & " " & _
'             CStr(SInput) & " " & _
'             CStr(ARelease) & " " & CStr(BRelease) & " " & CStr(CRelease) & " " & CStr(DRelease) & " " & _
'             CStr(XInput) & " " & _
'             CStr(YInput) & " " & _
'             CStr(ZInput) & " " & _
'             CStr(AAInput) & " " & _
'             CStr(ABInput) & " " & _
'             CStr(ACInput) & ";"
End Function
