Attribute VB_Name = "LoadCombination"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Raw As String

Private Label As String

Private SRSS As Long
Private PDelta As Long
Private Env As Long
Private Reverse As Long
Private Solved As Long

Private ASIF As Double
Private CD As Double
Private ABIF As Double

Private PhiSeismic As Long
Private Service As Long

Private HotRolledCheck As Long
Private ColdFormedCheck As Long
Private WoodCheck As Long
Private SteelProductCheck_ As Long
Private WoodProductCheck_ As Long
Private ConcreteCheck As Long
Private FootingsCheck As Long

Private Flag1 As Long
Private Flag2 As Long
Private Flag3 As Long
Private Flag4 As Long
Private Flag5 As Long
Private Flag6 As Long

Private Combos(1 To 10)

Private Sub Class_Initialize()
   
    Label = """" & CustomFunctions.BufferString("", 79) & """"
    
    SRSS = 0
    PDelta = 1
    Env = 1
    Reverse = 0
    Solved = 0
    
    ASIF = 1#
    CD = 1#
    ABIF = 1#
    
    PhiSeismic = 0
    Service = 0
    
    HotRolledCheck = 1
    ColdFormedCheck = 1
    WoodCheck = 1
    SteelProductCheck_ = 1
    WoodProductCheck_ = 1
    ConcreteCheck = 1
    FootingsCheck = 1
    
    Flag1 = -1
    Flag2 = 1
    Flag3 = 1
    Flag4 = 1
    Flag5 = 0
    Flag6 = 1
    
    For i = LBound(Combos) To UBound(Combos)
        temp1 = """" & CustomFunctions.BufferString("", 32) & """"
        temp2 = 0#
        
        Combos(i) = Array(temp1, temp2)
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

    regex.pattern = """.{79}""|"".{32}""|[^ \t\r\n\f;]+"
    If regex.Test(Raw) Then
        Set Data = regex.Execute(Raw)
    Else
        Exit Sub
    End If
    
    Label = Data(0)
    SRSS = CLng(Data(1))
    PDelta = CLng(Data(2))
    Env = CLng(Data(3))
    Reverse = CLng(Data(4))
    Solved = CLng(Data(5))
    ASIF = CDbl(Data(6))
    CD = CDbl(Data(7))
    ABIF = CDbl(Data(8))
    PhiSeismic = CLng(Data(9))
    Service = CLng(Data(10))
    HotRolledCheck = CLng(Data(11))
    ColdFormedCheck = CLng(Data(12))
    WoodCheck = CLng(Data(13))
    SteelProductCheck = CLng(Data(14))
    WoodProductCheck = CLng(Data(15))
    ConcreteCheck = CLng(Data(16))
    FootingsCheck = CLng(Data(17))
    Flag1 = CLng(Data(18))
    Flag2 = CLng(Data(19))
    Flag3 = CLng(Data(20))
    Flag4 = CLng(Data(21))
    Flag5 = CLng(Data(22))
    Flag6 = CLng(Data(23))
    
    k = 24
    For i = k To Data.Count - 1
        Temp = Data(i)
        If InStr(Temp, """") = 0 Then
            Temp = CDbl(Temp)
        End If
        Ci = WorksheetFunction.Ceiling((i - k + 1) / 2, 1)
        Cj = i Mod 2
        Combos(Ci)(Cj) = Temp
    Next i
    
    Set regex = Nothing
End Sub

Public Sub Create(Name)
    Label = CustomFunctions.BufferString(Name, 79)
End Sub

Public Sub AddBLC(BLCType, factor)

        For i = LBound(Combos) To UBound(Combos)
            If Combos(i)(0) = """" & CustomFunctions.BufferString("", 32) & """" Then
                Combos(i)(0) = """" & CustomFunctions.BufferString(BLCType, 32) & """"
                Combos(i)(1) = factor
                Exit Sub
            End If
        Next i

End Sub

Public Function StrRep()
    Temp = Array("""" & CustomFunctions.BufferString(Label, 79) & """", SRSS, PDelta, Env, Reverse, Solved, ASIF, CD, ABIF, _
                    PhiSeismic, Service, HotRolledCheck, ColdFormedCheck, WoodCheck, _
                    SteelProductCheck_, WoodProductCheck_, ConcreteCheck, FootingsCheck, Flag1, Flag2, Flag3, Flag4, Flag5, Flag6)
    StrRep = Join(Temp, " ")
    For Each Combo In Combos
        StrRep = StrRep & " " & Join(Combo, " ")
    Next Combo
    StrRep = StrRep & ";"
    
End Function
