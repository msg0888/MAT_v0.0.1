Attribute VB_Name = "Shape"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String

Private ShapeName As String

Private Database As Long
Private pShapeType As Long
Private MaterialType As Long


Private p(1 To 23) As Double
Private Unknown As Double
'various member dimensions not outlined in the RISA File Legend

Private Sub Class_Initialize()
    ShapeName = """" & CustomFunctions.BufferString("", 32) & """"
    
    Database = 0
    pShapeType = 0
    pMaterialType = 0
    
    For i = LBound(p) To UBound(p)
        p(i) = 0
    Next i
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

    'Pattern to match each piece of data in the shape
    regex.pattern = """.{32}""|[^ \t\r\n\f;]+"
    
    If regex.Test(Raw) Then
        Set Data = regex.Execute(Raw)
    Else
        Exit Sub
    End If
    
    ShapeName = Data(0)
    Database = CLng(Data(1))
    pShapeType = CLng(Data(2))
    MaterialType = CLng(Data(3))
    
    For i = LBound(p) To UBound(p)
        p(i) = CDbl(Data(i + 3))
    Next i
    
    Unknown = CDbl(Data(i + 3))
    
    Set regex = Nothing
End Sub

Public Property Get Name()
    Name = ShapeName
End Property
Public Property Get ShapeType()
    ShapeType = pShapeType
End Property
Public Property Get Dimensions()
    Dimensions = p
End Property

Public Function StrRep()
    Temp = ""
    For Each Num In p
        Temp = Temp & " " & CStr(Num)
    Next Num
    StrRep = Join(Array(ShapeName, Database, ShapeType, MaterialType), " ") & " " & Temp & " " & CStr(Unknown) & ";"
End Function
