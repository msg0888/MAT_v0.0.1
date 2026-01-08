Attribute VB_Name = "CustomFunctions"
Public Function XIF(ParamArray vars() As Variant)
    ' checks for an even number of inputs
    ' the first value of each pair is a condition, the second is the resulting value
    ' for a "catchall" condition, make the last pair TRUE, value

    XIF = CVErr(xlErrNA)
    If ((UBound(vars) - LBound(vars) + 1) Mod 2) <> 0 Then
        Exit Function
    End If
    
    Dim i
    For i = LBound(vars) To UBound(vars) Step 2
        'Debug.Print vars(i), vars(i + 1)
        If vars(i) Then
            XIF = vars(i + 1)
            Exit Function
        End If
    Next i
End Function

Public Function SIND(deg)
Dim Rad
Rad = WorksheetFunction.Radians(deg)
SIND = Sin(Rad)

End Function

Public Function COSD(deg)
Dim Rad
Rad = WorksheetFunction.Radians(deg)
COSD = Cos(Rad)
End Function

Public Function BufferString(str, Length, Optional BufferChar = " ")
    'don't forget to update this
    BufferString = str
    For i = 1 To Length - Len(str)
        BufferString = BufferString & BufferChar
    Next i
End Function


Public Function CaAa(TIA As String, Height As String, Width As String, Depth As String, Shape As String, Side As String, Optional c As String, Optional IceThickness As String)

Dim Ca
Dim Subcritical
Dim Transitional
Dim Supercritical
    
    If TIA = "I" Or TIA = 1 Then
        Subcritical = 32
        Supercritical = 64
    ElseIf TIA = "H" Or TIA = "G" Or TIA = "0" Then
        Subcritical = 39
        Supercritical = 78
    End If
    
    If IceThickness = "" Then
        IceThickness = 0
    ElseIf IceThickness <> "" Then
        IceThickness = IceThickness
    End If

    AspectRatio_Front = (Height + 2 * IceThickness) / (Width + 2 * IceThickness)
    AspectRatio_Side = (Height + 2 * IceThickness) / (Depth + 2 * IceThickness)
    
    If Shape = "Flat" Then
        Shape = 0
        c = 0
    ElseIf Shape = "Round" Then
        Shape = 1
        c = c
    Else
        Shape = Shape
    End If
    
    If Side = 0 Then
        If Shape = 0 Then
            If AspectRatio_Front <= 2.5 Then
                Ca = 1.2
            ElseIf AspectRatio_Front = 7 Then
                Ca = 1.4
            ElseIf AspectRatio_Front >= 25 Then
                Ca = 2
            ElseIf AspectRatio_Front > 2.5 And AspectRatio_Front < 7 Then
                Ca = 1.2 + (AspectRatio_Front - 2.5) * ((1.4 - 1.2) / (7 - 2.5))
            ElseIf AspectRatio_Front > 7 And AspectRatio_Front < 25 Then
                Ca = 1.4 + (AspectRatio_Front - 7) * ((2 - 1.4) / (25 - 7))
            Else
                Ca = 1
            End If
        ElseIf Shape = 1 Then
            If c < Subcritical Then
                If AspectRatio_Front <= 2.5 Then
                    Ca = 0.7
                ElseIf AspectRatio_Front > 2.5 And AspectRatio_Front < 7 Then
                    Ca = 0.7 + (AspectRatio_Front - 2.5) * ((0.8 - 0.7) / (7 - 2.5))
                ElseIf AspectRatio_Front = 7 Then
                    Ca = 0.8
                ElseIf AspectRatio_Front > 7 And AspectRatio_Front < 25 Then
                    Ca = 0.8 + (AspectRatio_Front - 7) * ((1.2 - 0.8) / (25 - 7))
                Else
                    Ca = 1.2
                End If
            ElseIf c >= Subcritical And c <= Supercritical Then
                If AspectRatio_Front <= 2.5 Then
                    Ca = (4.14 / (c ^ 0.485))
                ElseIf AspectRatio_Front > 2.5 And AspectRatio_Front < 7 Then
                    Ca = (4.14 / (c ^ 0.485)) + (AspectRatio_Front - 2.5) * (((3.66 / (c ^ 0.415)) - (4.14 / (c ^ 0.485))) / (7 - 2.5))
                ElseIf AspectRatio_Front = 7 Then
                    Ca = (3.66 / (c ^ 0.415))
                ElseIf AspectRatio_Front > 7 And AspectRatio_Front < 25 Then
                    Ca = (3.66 / (c ^ 0.415)) + (AspectRatio_Front - 7) * (((46.8 / (c ^ 1#)) - (3.66 / (c ^ 0.415))) / (25 - 7))
                Else
                    Ca = (46.8 / (c ^ 1#))
                End If
            ElseIf c > Supercritical Then
                If AspectRatio_Front <= 2.5 Then
                    Ca = 0.5
                ElseIf AspectRatio_Front > 2.5 And AspectRatio_Front < 7 Then
                    Ca = 0.5 + (AspectRatio_Front - 2.5) * ((0.6 - 0.5) / (7 - 2.5))
                ElseIf AspectRatio_Front >= 7 Then
                    Ca = 0.6
                End If
            End If
        End If
    ElseIf Side = 1 Then
        If Shape = 0 Then
            If AspectRatio_Side <= 2.5 Then
                Ca = 1.2
            ElseIf AspectRatio_Side = 7 Then
                Ca = 1.4
            ElseIf AspectRatio_Side >= 25 Then
                Ca = 2
            ElseIf AspectRatio_Side > 2.5 And AspectRatio_Side < 7 Then
                Ca = 1.2 + (AspectRatio_Side - 2.5) * ((1.4 - 1.2) / (7 - 2.5))
            ElseIf AspectRatio_Side > 7 And AspectRatio_Side < 25 Then
                Ca = 1.4 + (AspectRatio_Side - 7) * ((2 - 1.4) / (25 - 7))
            Else
                Ca = 1
            End If
        ElseIf Shape = 1 Then
            If c < Subcritical Then
                If AspectRatio_Side <= 2.5 Then
                    Ca = 0.7
                ElseIf AspectRatio_Side > 2.5 And AspectRatio_Side < 7 Then
                    Ca = 0.7 + (AspectRatio_Side - 2.5) * ((0.8 - 0.7) / (7 - 2.5))
                ElseIf AspectRatio_Side = 7 Then
                    Ca = 0.8
                ElseIf AspectRatio_Side > 7 And AspectRatio_Side < 25 Then
                    Ca = 0.8 + (AspectRatio_Side - 7) * ((1.2 - 0.8) / (25 - 7))
                Else
                    Ca = 1.2
                End If
            ElseIf c >= Subcritical And c <= Supercritical Then
                If AspectRatio_Side <= 2.5 Then
                    Ca = (4.14 / (c ^ 0.485))
                ElseIf AspectRatio_Side > 2.5 And AspectRatio_Side < 7 Then
                    Ca = (4.14 / (c ^ 0.485)) + (AspectRatio_Side - 2.5) * (((3.66 / (c ^ 0.415)) - (4.14 / (c ^ 0.485))) / (7 - 2.5))
                ElseIf AspectRatio_Side = 7 Then
                    Ca = (3.66 / (c ^ 0.415))
                ElseIf AspectRatio_Side > 7 And AspectRatio_Side < 25 Then
                    Ca = (3.66 / (c ^ 0.415)) + (AspectRatio_Side - 7) * (((46.8 / (c ^ 1#)) - (3.66 / (c ^ 0.415))) / (25 - 7))
                Else
                    Ca = (46.8 / (c ^ 1#))
                End If
            ElseIf c > Supercritical Then
                If AspectRatio_Side <= 2.5 Then
                    Ca = 0.5
                ElseIf AspectRatio_Side > 2.5 And AspectRatio_Side < 7 Then
                    Ca = 0.5 + (AspectRatio_Side - 2.5) * ((0.6 - 0.5) / (7 - 2.5))
                ElseIf AspectRatio_Side >= 7 Then
                    Ca = 0.6
                End If
            End If
        End If
    End If

If Side = 0 Then
    CaAa = Ca * (Height + 2 * IceThickness) * (Width + 2 * IceThickness)
Else
    CaAa = Ca * (Height + 2 * IceThickness) * (Depth + 2 * IceThickness)
End If
End Function


Public Function EPA(Length As String, ProjectedWidth As String, ProjectedDepth As String, Shape As String, Side As String, HorizontalOrientation As String, VerticalOrientation As String, Optional windDirection As String, Optional IceThickness As String)

Dim Projection
Dim Ca

    If Shape = "Flat" Then
        Shape = 0
    ElseIf Shape = "Round" Then
        Shape = 1
'    ElseIf Shape = "HSS Flat" Or Shape = 2 Then
    Else
        Shape = Shape
    End If

    If Shape = 0 Then
        Ca = 2
    ElseIf Shape = 1 Then
        Ca = 1.2
'    ElseIf Shape = "HSS Flat" Or Shape = 2 Then
    Else
    End If

    If windDirection = "" Then
        windDirection = 0
    Else
        windDirection = windDirection
    End If
        
    If IceThickness = "" Then
        IceThickness = 0
    Else
        IceThickness = IceThickness
    End If
    
    With WorksheetFunction

        memberFront = Length * Cos(.Pi / 180 * HorizontalOrientation + .Pi / 180 * windDirection) ^ 2
        If VerticalOrientation = 0 Then
            memberTransverse = 0
        Else
            memberTransverse = Length * Sin(.Pi / 180 * VerticalOrientation + .Pi / 180 * windDirection) ^ 2
        End If
        
        If Side = "Top" Or Side = 0 Then
            Projection = ProjectedDepth
        ElseIf Side = "Front" Or Side = 1 Then
            Projection = ProjectedWidth
        End If
        
        EPA = ((((Length + 2 * IceThickness) * (Projection + 2 * IceThickness)) - (Length * Projection)) * 1.2 * Cos(.Pi / 180 * (HorizontalOrientation)) ^ 2) + (memberFront + memberTransverse) * Projection * Ca
    
    End With
End Function



