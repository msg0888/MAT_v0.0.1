Attribute VB_Name = "SectionSetGroup"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw

Private SetName As String

Private NumSections As Long

Private Sections() As SectionSet


Private Sub Class_Initialize()
End Sub

Public Sub Define(dRaw)
    Raw = dRaw
    
    Process
End Sub


Private Sub Process()
    Dim regex As New RegExp
    With regex
        .Global = False
        .IgnoreCase = False
        .MultiLine = True
    End With

    'Pattern to get the Name of the section set Group
    regex.pattern = "(\[\.)(.+)(_SECTION_SETS\])"
    
    If regex.Test(Raw) Then
        SetName = regex.Execute(Raw)(0).SubMatches(1)
    Else
        Exit Sub
    End If
    
    
    
    'Pattern to get the data for each Section Set
    regex.Global = True
    regex.pattern = ".+;$"
    
    'Start Matching
    If regex.Test(Raw) Then
        Set Data = regex.Execute(Raw)
    Else
        Exit Sub
    End If
        
    'Determine Number of section sets in group
    NumSections = Data.Count
    ReDim Sections(1 To NumSections)
    
    'Precess each section set in group
    For i = LBound(Sections) To UBound(Sections)
        Set Sections(i) = New SectionSet
        Sections(i).DefineSection Data(i - 1).Value
    Next i

End Sub


Public Function StrRep()

    StrRep = "[." & SetName & "_SECTION_SETS] <" & CStr(NumSections) & ">" & vbNewLine
    For Each Sect In Sections
        StrRep = StrRep & Sect.StrRep & vbNewLine
    Next Sect
    StrRep = StrRep & "[.END_" & SetName & "_SECTION_SETS]"
End Function


Public Property Get Name()
    Name = SetName
End Property
