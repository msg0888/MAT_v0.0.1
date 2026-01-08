Attribute VB_Name = "Member"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String

Private MemberLabel As String
Private DesignList As String
Private ShapeLabel As String

Private pINode As Long ' a
Private pJNode As Long ' b
Private KNode As Long ' c

Private Rotation As Double ' d.dd

Private SectionSetOffset As Long ' e
Private TensionType As Long ' f
    ' 1 = compression
    ' 2 = tension
    ' 3 = Euler Buckling

Private DesignRules As Long ' g
Private MaterialType As Long ' h
    ' 0 = General Material
    ' 1 = Hot Rolled Steel
    ' 2 = Cold Formed Steel
    ' 3 = Wood
    ' 4 = Glulam
    ' 5 = Concrete
    ' 6 = Aluminum
    
Private MaterialOffset As Long ' i
Private EndReleases As Long ' j
Private IEndOffset As Double ' k
Private JEndOffset As Double ' l
Private isPhysical As Long ' m
    ' 0 = no
    ' 1 = yes
    
Private TopofMemberOffset As Long ' o
    ' 0 = no
    ' 1 = yes
    
Private MemberActivation As Long ' p
    ' 19 = active
    ' 1 = inactive
    ' 3 = excluded
    
Private Function3D As Long ' q
    ' 0 = gravity
    ' 1 = lateral

Private SeismicDesignRules As Long ' r
Private IEndConnectionRule As Long ' s
Private JEndConnectionRule As Long ' t

Private Num

Private pIsFlat
Private MemberDiameter
Private pLength


Private Sub Class_Initialize()


End Sub

Public Sub Define(dRaw)
    Raw = dRaw
    
    Process
End Sub


Sub Process()
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
    
    MemberLabel = Data(0)
    DesignList = Data(1)
    ShapeLabel = Data(2)
    pINode = CLng(Data(3))
    pJNode = CLng(Data(4))
    KNode = CLng(Data(5))
    Rotation = CDbl(Data(6))
    SectionSetOffset = CLng(Data(7))
    TensionType = CLng(Data(8))
    DesignRules = CLng(Data(9))
    MaterialType = CLng(Data(10))
    MaterialOffset = CLng(Data(11))
    EndReleases = CLng(Data(12))
    IEndOffset = CDbl(Data(13))
    JEndOffset = CDbl(Data(14))
    isPhysical = CLng(Data(15))
    TopofMemberOffset = CLng(Data(16))
    MemberActivation = CLng(Data(17))
    Function3D = CLng(Data(18))
    SeismicDesignRules = CLng(Data(19))
    IEndConnectionRule = CLng(Data(20))
    JEndConnectionRule = CLng(Data(21))
        
    Set regex = Nothing
End Sub


Public Function StrRep()
    StrRep = Raw
'    StrRep = Join(Array(MemberLabel, DesignList, ShapeLabel, INode, JNode, KNode, Rotation, SectionSetOffset, TensionType, _
'                        DesignRules, MaterialType, MaterialOffset, EndReleases, IEndOffset, JEndOffset, isPhysical, TopofMemberOffset, _
'                        MemberActivation, Function3D, SeismicDesignRules, IEndConnectionRule, JEndConnectionRule), " ") & ";"
End Function


Public Property Get Name()
    Name = Trim(Replace(MemberLabel, """", ""))
End Property
Public Property Get ShapeName()
    ShapeName = ShapeLabel
End Property
Public Property Get Number()
    Number = Num
End Property
Public Property Get SectionOffset()
    SectionOffset = SectionSetOffset
End Property
Public Property Get Material()
    Material = MaterialType
End Property
Public Property Get Dc()
    Dc = MemberDiameter
End Property
Public Property Get IsFlat()
    IsFlat = pIsFlat
End Property
Public Property Get iNode()
    iNode = pINode
End Property
Public Property Get jNode()
    jNode = pJNode
End Property
Public Property Get Length()
    Length = pLength
End Property



Public Property Let Dc(val)
    MemberDiameter = val
End Property
Public Property Let IsFlat(val)
    pIsFlat = val
End Property
Public Property Let Number(dNum)
    Num = dNum
End Property
Public Property Let Length(dLen)
    pLength = dLen
End Property

