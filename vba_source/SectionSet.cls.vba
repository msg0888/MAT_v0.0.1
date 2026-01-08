Attribute VB_Name = "SectionSet"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String

Private SectionLabel As String
Private DesignList As String
Private ShapeLabel As String


Private MemberType As Long
    '0 = None
    '1 = Beam
    '2 = Column
    '3 = HBrace
    '4 = VBrace

Private MaterialType As Long
    '0 = General Material
    '1 = Hot Rolled Steel
    '2 = Cold Formed Steel
    '3 = Wood
    '4 = Not Used
    '5 = Concrete
    
Private MaterialOffset As Long
    '0 is first Material, 1 is second, etc.
Private ShapeLock As Long
    '0 = Can be redesigned
    '1 = locked
Private RedesignRules As Long
    '-1 = no rules
    '0+ design rule offset by user design rules

Private Area As Double
Private Iyy As Double
Private Izz As Double
Private j As Double

Private Unknown As Long


Private Sub Class_Initialize()
End Sub


Public Sub DefineSection(dRaw)
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

    'Pattern to match each piece of data in the section
    regex.pattern = """.{32}""|[^ \t\r\n\f;]+"
    
    If regex.Test(Raw) Then
        Set Data = regex.Execute(Raw)
    Else
        Exit Sub
    End If
    
    SectionLabel = Data(0)
    DesignList = Data(1)
    ShapeLabel = Data(2)
    MemberType = CLng(Data(3))
    MaterialType = CLng(Data(4))
    MaterialOffset = CLng(Data(5))
    ShapeLock = CLng(Data(6))
    RedesignRules = CLng(Data(7))
    Area = CDbl(Data(8))
    Iyy = CDbl(Data(9))
    Izz = CDbl(Data(10))
    j = CDbl(Data(11))
    Unknown = CLng(Data(12))
End Sub

Public Function StrRep()
    StrRep = Join(Array(SectionLabel, DesignList, ShapeLabel, MemberType, MaterialType, MaterialOffset, ShapeLock, RedesignRules, Area, Iyy, Izz, j, Unknown), " ") & ";"
End Function
