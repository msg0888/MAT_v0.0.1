Attribute VB_Name = "Loads_Distributed"
Option Explicit

Dim NextRow As Integer, firstRow As Long
Dim lastRow As Long, LastRow_Nodes As Long

Dim a As Double, b As Double                      'Counter to detmerine size of array
Dim c As Double, d As Double                      'Counter to find array of members
Dim r As Double
Dim i As Double, j As Double

Dim W As Double                                     'Projected Width of member
Dim wi As Double                                    'Projected Width of member with ice

'    Dim wl As Double, wil As Double                    'Projected Width of member
Dim EPA As Double, IceEPA As Double                 'Effective Projected Area of member
Dim Front As Double, IceFront As Double             'Front and Ice Front length of member
Dim Transverse As Double, IceTransverse As Double   'Transverse and Ice Transverse length of member
Dim Length As Double, IceLength As Double           'Member lengths
Dim IceWidth As Double
Dim Azimuth As Double, Orientation As Double        'Mount and member orientations
Dim Increment As Variant
Dim WindProjection As Double
Dim IceWindProjection As Double

Dim Dc As Double                                    'Out-to-out dimension for calculating Ice weight (See TIA-222-H, Figure 2-4)
Dim Aiz As Double                                   'Cross-sectional area ice
Dim Ca As Double                                    'Force Coefficient
Dim p As Double, pice As Double                     'Applied Force
Dim DiLoad As Double                                'Distibuted Ice Dead Loads
Dim ConvertLinearForce As Double
Dim ConvertLength As Double
Const Pi As Double = 3.141592654

Dim qz As Range
Dim qiz As Range
Dim tiz As Range
Dim Ka As Range
Dim Gh As Range
Dim Kesw As Range
Dim Kesi As Range

Sub Distributed_Loads()

With shDistributedLoadTables
    If .Range("A4") <> "" Then
    lastRow = .Range("A4").CurrentRegion.Rows.Count
    .Range("A4", "A" & lastRow).EntireRow.ClearContents
    End If
End With

    Call Distributed_Loads_Wind_Z
    Call Distributed_Loads_Wind_X
    Call Distributed_Loads_IceWind_Z
    Call Distributed_Loads_IceWind_X
    
End Sub

Sub Distributed_Loads_Wind_Z()

    Dim pw_0_Z() As Variant, pw_0_X As Variant          '0º distibuted load arrays
    Dim pw_30_Z() As Variant, pw_30_X As Variant        '30º distibuted load arrays
    Dim pw_45_Z() As Variant, pw_45_X As Variant        '45º distibuted load arrays
    Dim pw_60_Z() As Variant, pw_60_X As Variant        '60º distibuted load arrays
    Dim pw_90_Z() As Variant, pw_90_X As Variant        '90º distibuted load arrays
    Dim pw_120_Z() As Variant, pw_120_X As Variant      '120º distibuted load arrays
    Dim pw_135_Z() As Variant, pw_135_X As Variant      '135º distibuted load arrays
    Dim pw_150_Z() As Variant, pw_150_X As Variant      '150º distibuted load arrays
    
    Dim pw_180_Z() As Variant, pw_180_X As Variant      '0º distibuted load arrays
    Dim pw_210_Z() As Variant, pw_210_X As Variant      '30º distibuted load arrays
    Dim pw_225_Z() As Variant, pw_225_X As Variant      '45º distibuted load arrays
    Dim pw_240_Z() As Variant, pw_240_X As Variant      '60º distibuted load arrays
    Dim pw_270_Z() As Variant, pw_270_X As Variant      '90º distibuted load arrays
    Dim pw_300_Z() As Variant, pw_300_X As Variant      '120º distibuted load arrays
    Dim pw_315_Z() As Variant, pw_315_X As Variant      '135º distibuted load arrays
    Dim pw_330_Z() As Variant, pw_330_X As Variant      '150º distibuted load arrays
    
    Dim NextRow As Double, firstRow As Long
    Dim lastRow As Long, LastRow_Nodes As Long
    
    Dim a As Double, b As Double                      'Counter to detmerine size of array
    Dim c As Double, d As Double                      'Counter to find array of members
    Dim r As Double
    
    Dim W As Double                                     'Projected Width of member
    Dim wi As Double                                    'Projected Width of member with ice
    
    '    Dim wl As Double, wil As Double                    'Projected Width of member
    Dim EPA As Double, IceEPA As Double                 'Effective Projected Area of member
    Dim Front As Double, IceFront As Double             'Front and Ice Front length of member
    Dim Transverse As Double, IceTransverse As Double   'Transverse and Ice Transverse length of member
    Dim Length As Double, IceLength As Double           'Member lengths
    Dim Azimuth As Double, Orientation As Double        'Mount and member orientations
    Dim Increment As Variant
    Dim WindProjection As Double
    Dim IceWindProjection As Double
    
    Dim Dc As Double                                    'Out-to-out dimension for calculating Ice weight (See TIA-222-H, Figure 2-4)
    Dim Aiz As Double                                   'Cross-sectional area ice
    Dim Ca As Double                                    'Force Coefficient
    Dim p As Double, pice As Double                     'Applied Force
    Dim DiLoad As Double                                'Distibuted Ice Dead Loads
    Dim ConvertLinearForce As Double
    Dim ConvertLength2in As Double, ConvertLength2ft As Double
    Const Pi As Double = 3.141592654
    
    Dim qz As Range
    Dim qiz As Range
    Dim tiz As Range
    Dim Ka As Range
    Dim Gh As Range
    Dim Kesw As Range
    Dim Kesi As Range


    With shCode
    If .Range("RISA3D.Unit.LinearForce") = "pli" Then
        ConvertLinearForce = 1 / 12
        ElseIf .Range("RISA3D.Unit.LinearForce") = "plf" Then
        ConvertLinearForce = 1
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kli" Then
        ConvertLinearForce = 1 / 1000 * 1 / 12
        ElseIf .Range("RISA3D.Unit.LinearForce") = "klf" Then
        ConvertLinearForce = 1 / 1000
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/m" Then
        ConvertLinearForce = 0.0145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/cm" Then
        ConvertLinearForce = 0.000145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/mm" Then
        ConvertLinearForce = 0.0000145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/m" Then
        ConvertLinearForce = 14.5939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/cm" Then
        ConvertLinearForce = 0.145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/mm" Then
        ConvertLinearForce = 0.0145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/m" Then
        ConvertLinearForce = 0.001489173
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/cm" Then
        ConvertLinearForce = 0.0000148917
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/mm" Then
        ConvertLinearForce = 0.00000148917
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/m" Then
        ConvertLinearForce = 1.489173228
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/cm" Then
        ConvertLinearForce = 0.014891732
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/mm" Then
        ConvertLinearForce = 0.001489173
    End If
    
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength2in = 1
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength2in = 12
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength2in = 0.3048
        ElseIf .Range("RISA3D.Unit.Length") = "cm" Then
        ConvertLength2in = 30.48
        ElseIf .Range("RISA3D.Unit.Length") = "mm" Then
        ConvertLength2in = 304.8
    End If
    
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength2ft = 1 / 12
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength2ft = 1
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength2ft = 3.280839895
        ElseIf .Range("RISA3D.Unit.Length") = "cm" Then
        ConvertLength2ft = 0.032808399
        ElseIf .Range("RISA3D.Unit.Length") = "mm" Then
        ConvertLength2ft = 0.00328084
    End If
    End With
    
    Set qz = shCode.Range("qz")
    Set qiz = shCode.Range("qiz")
    Set tiz = shCode.Range("tiz")
    Set Ka = shCode.Range("Ka")
    Set Gh = shCode.Range("Gh")
    Set Kesw = shCode.Range("Kes.w")
    Set Kesi = shCode.Range("Kes.i")

'    On Error GoTo ErrorHandle
    
   With shGeometry
      lastRow = .Range("B6").CurrentRegion.End(xlDown).row
      LastRow_Nodes = .Range("R6").CurrentRegion.End(xlDown).row
      
      For c = 6 To lastRow                                   'Determine size of array by filtering out "RIGID" members
         If .Range("C" & c).Value <> "RIGID" Then
            a = a + 1
         End If
      Next c
      
      If a >= 1 Then
         'Defining the size of arrays
         ReDim pw_0_Z(1 To a, 1 To 6)
         ReDim pw_30_Z(1 To a, 1 To 6)
         ReDim pw_45_Z(1 To a, 1 To 6)
         ReDim pw_60_Z(1 To a, 1 To 6)
         ReDim pw_90_Z(1 To a, 1 To 6)
         ReDim pw_120_Z(1 To a, 1 To 6)
         ReDim pw_135_Z(1 To a, 1 To 6)
         ReDim pw_150_Z(1 To a, 1 To 6)
         
         ReDim pw_180_Z(1 To a, 1 To 6)
         ReDim pw_210_Z(1 To a, 1 To 6)
         ReDim pw_225_Z(1 To a, 1 To 6)
         ReDim pw_240_Z(1 To a, 1 To 6)
         ReDim pw_270_Z(1 To a, 1 To 6)
         ReDim pw_300_Z(1 To a, 1 To 6)
         ReDim pw_315_Z(1 To a, 1 To 6)
         ReDim pw_330_Z(1 To a, 1 To 6)
         
      Else
         MsgBox "Member information must be entered before distributed loads can be calculated.", vbInformation, "Members do not exist"
         Exit Sub
      End If
      
      With Application
         .ScreenUpdating = False
         .Calculation = xlCalculationManual
         .StatusBar = "Calculating distributed loads"
      End With
      
      'Populates arrays
      a = 1
      For c = 6 To lastRow
         If .Range("C" & c) <> "RIGID" Then       'Only runs loop for non-Rigid members. "If" statement could be removed since "importing_members" filters out RIGID Links, but
                                                  'decided to leave incase user prefers to copy/paste members from RISA 3D
         pw_0_Z(a, 1) = .Cells(c, 2)
         pw_30_Z(a, 1) = .Cells(c, 2)
         pw_45_Z(a, 1) = .Cells(c, 2)
         pw_60_Z(a, 1) = .Cells(c, 2)
         pw_90_Z(a, 1) = .Cells(c, 2)
         pw_120_Z(a, 1) = .Cells(c, 2)
         pw_135_Z(a, 1) = .Cells(c, 2)
         pw_150_Z(a, 1) = .Cells(c, 2)

         pw_180_Z(a, 1) = .Cells(c, 2)
         pw_210_Z(a, 1) = .Cells(c, 2)
         pw_225_Z(a, 1) = .Cells(c, 2)
         pw_240_Z(a, 1) = .Cells(c, 2)
         pw_270_Z(a, 1) = .Cells(c, 2)
         pw_300_Z(a, 1) = .Cells(c, 2)
         pw_315_Z(a, 1) = .Cells(c, 2)
         pw_300_Z(a, 1) = .Cells(c, 2)
         pw_330_Z(a, 1) = .Cells(c, 2)
      
         If WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 4, False) = "Flat" Then       'Determining p coefficient based on Round or Flat Profile
            Ca = 2
         ElseIf WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 4, False) = "HSS Flat" Then
            If WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 6, False) Then
                Ca = WorksheetFunction.Max(1.2 - 2.8 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 0.85)
            ElseIf WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 6, False) >= 25 Then
                Ca = WorksheetFunction.Max(2 - 6 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 1.25)
            Else
                Ca = WorksheetFunction.Max(1.4 - 4 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 0.9)
            End If
         Else
            Ca = 1.2
         End If
                
         
         p = qz * Ka * Kesw * Gh                               'Calculate Forces for NO ICE condition
         pice = qiz * Ka * Kesi * Gh                            'Caculates p WITH ICE condition. Per Code, can assume Ca=1.2 (round) for projected area for ice conditions

         Dc = WorksheetFunction.VLookup(.Cells(c, 3), .Range("rngR3DSectionSets"), 9, False)
         Aiz = Pi * tiz * ((Dc + shCode.Range("tiz")))  'in^2
         DiLoad = Aiz '* (56 / 1000)  'kips


'>>>No Ice<<<'
With shGeometry
         Length = .Cells(c, 15).Value / 12 'in
         WindProjection = WorksheetFunction.VLookup(.Cells(c, 3).Value, .Range("rngR3DSectionSets"), 5, False) / 12 'in
         j = WorksheetFunction.VLookup(shGeometry.Range("B" & c), shGeometry.Range("E6:K99999"), 6, False) * ConvertLength2ft
         Orientation = .Cells(c, 16).Value 'degrees
         
         Azimuth = 0
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_0_Z(a, 2) = "Z"
         pw_0_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_0_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_0_Z(a, 5) = "0.0%"
         pw_0_Z(a, 6) = "100.0%"
         
         Azimuth = 30
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_30_Z(a, 2) = "Z"
         pw_30_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_30_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_30_Z(a, 5) = "0.0%"
         pw_30_Z(a, 6) = "100.0%"
         
         Azimuth = 45
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_45_Z(a, 2) = "Z"
         pw_45_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_45_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_45_Z(a, 5) = "0.0%"
         pw_45_Z(a, 6) = "100.0%"
         
         Azimuth = 60
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_60_Z(a, 2) = "Z"
         pw_60_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_60_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_60_Z(a, 5) = "0.0%"
         pw_60_Z(a, 6) = "100.0%"
         
         Azimuth = 90
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_90_Z(a, 2) = "Z"
         pw_90_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_90_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_90_Z(a, 5) = "0.0%"
         pw_90_Z(a, 6) = "100.0%"
         
         Azimuth = 120
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_120_Z(a, 2) = "Z"
         pw_120_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_120_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_120_Z(a, 5) = "0.0%"
         pw_120_Z(a, 6) = "100.0%"
         
         Azimuth = 135
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_135_Z(a, 2) = "Z"
         pw_135_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_135_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_135_Z(a, 5) = "0.0%"
         pw_135_Z(a, 6) = "100.0%"
         
         Azimuth = 150
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_150_Z(a, 2) = "Z"
         pw_150_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_150_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_150_Z(a, 5) = "0.0%"
         pw_150_Z(a, 6) = "100.0%"
         
         Azimuth = 180
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_180_Z(a, 2) = "Z"
         pw_180_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_180_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_180_Z(a, 5) = "0.0%"
         pw_180_Z(a, 6) = "100.0%"
         
         Azimuth = 210
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_210_Z(a, 2) = "Z"
         pw_210_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_210_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_210_Z(a, 5) = "0.0%"
         pw_210_Z(a, 6) = "100.0%"
         
         Azimuth = 225
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_225_Z(a, 2) = "Z"
         pw_225_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_225_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_225_Z(a, 5) = "0.0%"
         pw_225_Z(a, 6) = "100.0%"
         
         Azimuth = 240
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_240_Z(a, 2) = "Z"
         pw_240_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_240_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_240_Z(a, 5) = "0.0%"
         pw_240_Z(a, 6) = "100.0%"
         
         Azimuth = 270
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_270_Z(a, 2) = "Z"
         pw_270_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_270_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_270_Z(a, 5) = "0.0%"
         pw_270_Z(a, 6) = "100.0%"
         
         Azimuth = 300
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_300_Z(a, 2) = "Z"
         pw_300_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_300_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_300_Z(a, 5) = "0.0%"
         pw_300_Z(a, 6) = "100.0%"
         
         Azimuth = 315
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_315_Z(a, 2) = "Z"
         pw_315_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_315_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_315_Z(a, 5) = "0.0%"
         pw_315_Z(a, 6) = "100.0%"
         
         Azimuth = 330
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_330_Z(a, 2) = "Z"
         pw_330_Z(a, 3) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_330_Z(a, 4) = -ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Cos(Azimuth * Pi / 180), 0)
         pw_330_Z(a, 5) = "0.0%"
         pw_330_Z(a, 6) = "100.0%"
End With

         a = a + 1
         End If
         Next c
   End With       'worksheets("Import Model)

     
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
With Worksheets("Distributed Load Tables")
    If .Range("A4") <> "" Then
        lastRow = .Range("A4").CurrentRegion.Rows.Count
    End If

If a >= 1 Then
NextRow = .Range("A3").CurrentRegion.Rows.Count + 1

      'Prints array to Loads Tables sheet
      .Range("A" & NextRow, "F" & NextRow + a - 2) = pw_0_Z
      .Range("H" & NextRow, "M" & NextRow + a - 2) = pw_30_Z
      .Range("O" & NextRow, "T" & NextRow + a - 2) = pw_45_Z
      .Range("V" & NextRow, "AA" & NextRow + a - 2) = pw_60_Z
      .Range("AC" & NextRow, "AH" & NextRow + a - 2) = pw_90_Z
      .Range("AJ" & NextRow, "AO" & NextRow + a - 2) = pw_120_Z
      .Range("AQ" & NextRow, "AV" & NextRow + a - 2) = pw_135_Z
      .Range("AX" & NextRow, "BC" & NextRow + a - 2) = pw_150_Z

      .Range("BE" & NextRow, "BJ" & NextRow + a - 2) = pw_180_Z
      .Range("BL" & NextRow, "BQ" & NextRow + a - 2) = pw_210_Z
      .Range("BS" & NextRow, "BX" & NextRow + a - 2) = pw_225_Z
      .Range("BZ" & NextRow, "CE" & NextRow + a - 2) = pw_240_Z
      .Range("CG" & NextRow, "CL" & NextRow + a - 2) = pw_270_Z
      .Range("CN" & NextRow, "CS" & NextRow + a - 2) = pw_300_Z
      .Range("CU" & NextRow, "CZ" & NextRow + a - 2) = pw_315_Z
      .Range("DB" & NextRow, "DG" & NextRow + a - 2) = pw_330_Z

End If

      lastRow = .Range("A" & Rows.Count).End(xlUp).row
      .Range("C" & NextRow, "D" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("E" & NextRow, "F" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("J" & NextRow, "K" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("L" & NextRow, "M" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("Q" & NextRow, "R" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("S" & NextRow, "T" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("X" & NextRow, "Y" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("Z" & NextRow, "AA" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("AE" & NextRow, "AF" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("AG" & NextRow, "AH" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("AL" & NextRow, "AM" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("AN" & NextRow, "AO" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("AS" & NextRow, "AT" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("AU" & NextRow, "AV" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("AZ" & NextRow, "BA" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("BB" & NextRow, "BC" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("BG" & NextRow, "BH" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("BI" & NextRow, "BJ" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("BN" & NextRow, "BO" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("BP" & NextRow, "BQ" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("BU" & NextRow, "BV" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("BW" & NextRow, "BX" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("CB" & NextRow, "CC" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("CD" & NextRow, "CE" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("CI" & NextRow, "CJ" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("CK" & NextRow, "CL" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("CP" & NextRow, "CQ" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("CR" & NextRow, "CS" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("CW" & NextRow, "CX" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("CY" & NextRow, "CZ" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("DD" & NextRow, "DE" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("DF" & NextRow, "DG" & NextRow + a - 2).NumberFormat = "0.0%"
   
   End With    'Worksheets("Distributed Load Tables")
        
'    Call Sort_Distributed

      With Application
         .ScreenUpdating = True
         .Calculation = xlCalculationAutomatic
         .StatusBar = "Ready"
      End With
'   Worksheets("Results").Select
'   Range("A3").Select
   Exit Sub

'ErrorHandle:
'MsgBox "There appears to be an error" & vbNewLine & "with the imported RISA-3D model.", , "Error!"

   With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .StatusBar = "Ready"
   End With
End Sub

Sub Distributed_Loads_Wind_X()

    Dim pw_0_Z() As Variant, pw_0_X As Variant          '0º distibuted load arrays
    Dim pw_30_Z() As Variant, pw_30_X As Variant        '30º distibuted load arrays
    Dim pw_45_Z() As Variant, pw_45_X As Variant        '45º distibuted load arrays
    Dim pw_60_Z() As Variant, pw_60_X As Variant        '60º distibuted load arrays
    Dim pw_90_Z() As Variant, pw_90_X As Variant        '90º distibuted load arrays
    Dim pw_120_Z() As Variant, pw_120_X As Variant      '120º distibuted load arrays
    Dim pw_135_Z() As Variant, pw_135_X As Variant      '135º distibuted load arrays
    Dim pw_150_Z() As Variant, pw_150_X As Variant      '150º distibuted load arrays
    
    Dim pw_180_Z() As Variant, pw_180_X As Variant      '0º distibuted load arrays
    Dim pw_210_Z() As Variant, pw_210_X As Variant      '30º distibuted load arrays
    Dim pw_225_Z() As Variant, pw_225_X As Variant      '45º distibuted load arrays
    Dim pw_240_Z() As Variant, pw_240_X As Variant      '60º distibuted load arrays
    Dim pw_270_Z() As Variant, pw_270_X As Variant      '90º distibuted load arrays
    Dim pw_300_Z() As Variant, pw_300_X As Variant      '120º distibuted load arrays
    Dim pw_315_Z() As Variant, pw_315_X As Variant      '135º distibuted load arrays
    Dim pw_330_Z() As Variant, pw_330_X As Variant      '150º distibuted load arrays
    
    Dim NextRow As Double, firstRow As Long
    Dim lastRow As Long, LastRow_Nodes As Long
    
    Dim a As Double, b As Double                      'Counter to detmerine size of array
    Dim c As Double, d As Double                      'Counter to find array of members
    Dim r As Double
    
    Dim W As Double                                     'Projected Width of member
    Dim wi As Double                                    'Projected Width of member with ice
    
    '    Dim wl As Double, wil As Double                    'Projected Width of member
    Dim EPA As Double, IceEPA As Double                 'Effective Projected Area of member
    Dim Front As Double, IceFront As Double             'Front and Ice Front length of member
    Dim Transverse As Double, IceTransverse As Double   'Transverse and Ice Transverse length of member
    Dim Length As Double, IceLength As Double           'Member lengths
    Dim Azimuth As Double, Orientation As Double        'Mount and member orientations
    Dim Increment As Variant
    Dim WindProjection As Double
    Dim IceWindProjection As Double
    
    Dim Dc As Double                                    'Out-to-out dimension for calculating Ice weight (See TIA-222-H, Figure 2-4)
    Dim Aiz As Double                                   'Cross-sectional area ice
    Dim Ca As Double                                    'Force Coefficient
    Dim p As Double, pice As Double                     'Applied Force
    Dim DiLoad As Double                                'Distibuted Ice Dead Loads
    Dim ConvertLinearForce As Double
    Dim ConvertLength2in As Double, ConvertLength2ft As Double
    Const Pi As Double = 3.141592654
    
    Dim qz As Range
    Dim qiz As Range
    Dim tiz As Range
    Dim Ka As Range
    Dim Gh As Range
    Dim Kesw As Range
    Dim Kesi As Range


    With shCode
    If .Range("RISA3D.Unit.LinearForce") = "pli" Then
        ConvertLinearForce = 1 / 12
        ElseIf .Range("RISA3D.Unit.LinearForce") = "plf" Then
        ConvertLinearForce = 1
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kli" Then
        ConvertLinearForce = 1 / 1000 * 1 / 12
        ElseIf .Range("RISA3D.Unit.LinearForce") = "klf" Then
        ConvertLinearForce = 1 / 1000
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/m" Then
        ConvertLinearForce = 0.0145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/cm" Then
        ConvertLinearForce = 0.000145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/mm" Then
        ConvertLinearForce = 0.0000145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/m" Then
        ConvertLinearForce = 14.5939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/cm" Then
        ConvertLinearForce = 0.145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/mm" Then
        ConvertLinearForce = 0.0145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/m" Then
        ConvertLinearForce = 0.001489173
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/cm" Then
        ConvertLinearForce = 0.0000148917
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/mm" Then
        ConvertLinearForce = 0.00000148917
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/m" Then
        ConvertLinearForce = 1.489173228
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/cm" Then
        ConvertLinearForce = 0.014891732
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/mm" Then
        ConvertLinearForce = 0.001489173
    End If
    
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength2in = 1
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength2in = 12
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength2in = 0.3048
        ElseIf .Range("RISA3D.Unit.Length") = "cm" Then
        ConvertLength2in = 30.48
        ElseIf .Range("RISA3D.Unit.Length") = "mm" Then
        ConvertLength2in = 304.8
    End If
    
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength2ft = 1 / 12
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength2ft = 1
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength2ft = 3.280839895
        ElseIf .Range("RISA3D.Unit.Length") = "cm" Then
        ConvertLength2ft = 0.032808399
        ElseIf .Range("RISA3D.Unit.Length") = "mm" Then
        ConvertLength2ft = 0.00328084
    End If
    End With
    
    Set qz = shCode.Range("qz")
    Set qiz = shCode.Range("qiz")
    Set tiz = shCode.Range("tiz")
    Set Ka = shCode.Range("Ka")
    Set Gh = shCode.Range("Gh")
    Set Kesw = shCode.Range("Kes.w")
    Set Kesi = shCode.Range("Kes.i")

'    On Error GoTo ErrorHandle
    
   With shGeometry
      lastRow = .Range("B6").CurrentRegion.End(xlDown).row
      LastRow_Nodes = .Range("R6").CurrentRegion.End(xlDown).row
      
      For c = 6 To lastRow                                   'Determine size of array by filtering out "RIGID" members
         If .Range("C" & c).Value <> "RIGID" Then
            a = a + 1
         End If
      Next c
      
      If a >= 1 Then
         'Defining the size of arrays
         ReDim pw_0_X(1 To a, 1 To 6)
         ReDim pw_30_X(1 To a, 1 To 6)
         ReDim pw_45_X(1 To a, 1 To 6)
         ReDim pw_60_X(1 To a, 1 To 6)
         ReDim pw_90_X(1 To a, 1 To 6)
         ReDim pw_120_X(1 To a, 1 To 6)
         ReDim pw_135_X(1 To a, 1 To 6)
         ReDim pw_150_X(1 To a, 1 To 6)
         
         ReDim pw_180_X(1 To a, 1 To 6)
         ReDim pw_210_X(1 To a, 1 To 6)
         ReDim pw_225_X(1 To a, 1 To 6)
         ReDim pw_240_X(1 To a, 1 To 6)
         ReDim pw_270_X(1 To a, 1 To 6)
         ReDim pw_300_X(1 To a, 1 To 6)
         ReDim pw_315_X(1 To a, 1 To 6)
         ReDim pw_330_X(1 To a, 1 To 6)
      Else
         MsgBox "Member information must be entered before distributed loads can be calculated.", vbInformation, "Members do not exist"
         Exit Sub
      End If
      
      With Application
         .ScreenUpdating = False
         .Calculation = xlCalculationManual
         .StatusBar = "Calculating distributed loads"
      End With
      
      'Populates arrays
      a = 1
      For c = 6 To lastRow
         If .Range("C" & c) <> "RIGID" Then       'Only runs loop for non-Rigid members. "If" statement could be removed since "importing_members" filters out RIGID Links, but
                                                  'decided to leave incase user prefers to copy/paste members from RISA 3D
         pw_0_X(a, 1) = .Cells(c, 2)
         pw_30_X(a, 1) = .Cells(c, 2)
         pw_45_X(a, 1) = .Cells(c, 2)
         pw_60_X(a, 1) = .Cells(c, 2)
         pw_90_X(a, 1) = .Cells(c, 2)
         pw_120_X(a, 1) = .Cells(c, 2)
         pw_135_X(a, 1) = .Cells(c, 2)
         pw_150_X(a, 1) = .Cells(c, 2)

         pw_180_X(a, 1) = .Cells(c, 2)
         pw_210_X(a, 1) = .Cells(c, 2)
         pw_225_X(a, 1) = .Cells(c, 2)
         pw_240_X(a, 1) = .Cells(c, 2)
         pw_270_X(a, 1) = .Cells(c, 2)
         pw_300_X(a, 1) = .Cells(c, 2)
         pw_315_X(a, 1) = .Cells(c, 2)
         pw_300_X(a, 1) = .Cells(c, 2)
         pw_330_X(a, 1) = .Cells(c, 2)
      
         If WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 4, False) = "Flat" Then       'Determining p coefficient based on Round or Flat Profile
            Ca = 2
         ElseIf WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 4, False) = "HSS Flat" Then
            If WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 6, False) Then
                Ca = WorksheetFunction.Max(1.2 - 2.8 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 0.85)
            ElseIf WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 6, False) >= 25 Then
                Ca = WorksheetFunction.Max(2 - 6 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 1.25)
            Else
                Ca = WorksheetFunction.Max(1.4 - 4 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 0.9)
            End If
         Else
            Ca = 1.2
         End If
                
         
         p = qz * Ka * Kesw * Gh                               'Calculate Forces for NO ICE condition
         pice = qiz * Ka * Kesi * Gh                            'Caculates p WITH ICE condition. Per Code, can assume Ca=1.2 (round) for projected area for ice conditions

         Dc = WorksheetFunction.VLookup(.Cells(c, 3), .Range("rngR3DSectionSets"), 9, False)
         Aiz = Pi * tiz * ((Dc + shCode.Range("tiz")))  'ft^2
         DiLoad = Aiz '* (56 / 1000)  'kips


'>>>No Ice<<<
With shGeometry
         Length = .Cells(c, 15).Value / 12 'in
         WindProjection = WorksheetFunction.VLookup(.Cells(c, 3).Value, .Range("rngR3DSectionSets"), 5, False) / 12
         j = WorksheetFunction.VLookup(shGeometry.Range("B" & c), shGeometry.Range("E6:K99999"), 6, False) * ConvertLength2ft
         Orientation = .Cells(c, 16).Value
         
         Azimuth = 0
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         
         pw_0_X(a, 2) = "X"
         pw_0_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_0_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_0_X(a, 5) = "0.0%"
         pw_0_X(a, 6) = "100.0%"
         
         Azimuth = 30
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_30_X(a, 2) = "X"
         pw_30_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_30_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_30_X(a, 5) = "0.0%"
         pw_30_X(a, 6) = "100.0%"
         
         Azimuth = 45
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_45_X(a, 2) = "X"
         pw_45_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_45_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_45_X(a, 5) = "0.0%"
         pw_45_X(a, 6) = "100.0%"
         
         Azimuth = 60
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_60_X(a, 2) = "X"
         pw_60_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_60_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_60_X(a, 5) = "0.0%"
         pw_60_X(a, 6) = "100.0%"
         
         Azimuth = 90
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_90_X(a, 2) = "X"
         pw_90_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_90_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_90_X(a, 5) = "0.0%"
         pw_90_X(a, 6) = "100.0%"
         
         Azimuth = 120
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_120_X(a, 2) = "X"
         pw_120_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_120_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_120_X(a, 5) = "0.0%"
         pw_120_X(a, 6) = "100.0%"
         
         Azimuth = 135
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_135_X(a, 2) = "X"
         pw_135_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_135_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_135_X(a, 5) = "0.0%"
         pw_135_X(a, 6) = "100.0%"
         
         Azimuth = 150
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_150_X(a, 2) = "X"
         pw_150_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_150_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_150_X(a, 5) = "0.0%"
         pw_150_X(a, 6) = "100.0%"
         
         Azimuth = 180
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_180_X(a, 2) = "X"
         pw_180_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_180_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_180_X(a, 5) = "0.0%"
         pw_180_X(a, 6) = "100.0%"
         
         Azimuth = 210
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_210_X(a, 2) = "X"
         pw_210_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_210_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_210_X(a, 5) = "0.0%"
         pw_210_X(a, 6) = "100.0%"
         
         Azimuth = 225
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_225_X(a, 2) = "X"
         pw_225_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_225_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_225_X(a, 5) = "0.0%"
         pw_225_X(a, 6) = "100.0%"
         
         Azimuth = 240
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_240_X(a, 2) = "X"
         pw_240_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_240_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_240_X(a, 5) = "0.0%"
         pw_240_X(a, 6) = "100.0%"
         
         Azimuth = 270
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_270_X(a, 2) = "X"
         pw_270_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_270_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_270_X(a, 5) = "0.0%"
         pw_270_X(a, 6) = "100.0%"
         
         Azimuth = 300
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_300_X(a, 2) = "X"
         pw_300_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_300_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_300_X(a, 5) = "0.0%"
         pw_300_X(a, 6) = "100.0%"
         
         Azimuth = 315
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_315_X(a, 2) = "X"
         pw_315_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_315_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_315_X(a, 5) = "0.0%"
         pw_315_X(a, 6) = "100.0%"
         
         Azimuth = 330
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         
         pw_330_X(a, 2) = "X"
         pw_330_X(a, 3) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_330_X(a, 4) = ConvertLinearForce * WorksheetFunction.IfError(p * EPA / Length * Sin(Azimuth * Pi / 180), 0)
         pw_330_X(a, 5) = "0.0%"
         pw_330_X(a, 6) = "100.0%"
End With
        a = a + 1
         End If
         Next c
   End With       'worksheets("Import Model)

     
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
With Worksheets("Distributed Load Tables")
    If .Range("A4") <> "" Then
    lastRow = .Range("A4").CurrentRegion.Rows.Count
End If


If a >= 1 Then
NextRow = .Range("A3").CurrentRegion.Rows.Count + 1

      'Prints array to Loads Tables sheet
      .Range("A" & NextRow, "F" & NextRow + a - 2) = pw_0_X
      .Range("H" & NextRow, "M" & NextRow + a - 2) = pw_30_X
      .Range("O" & NextRow, "T" & NextRow + a - 2) = pw_45_X
      .Range("V" & NextRow, "AA" & NextRow + a - 2) = pw_60_X
      .Range("AC" & NextRow, "AH" & NextRow + a - 2) = pw_90_X
      .Range("AJ" & NextRow, "AO" & NextRow + a - 2) = pw_120_X
      .Range("AQ" & NextRow, "AV" & NextRow + a - 2) = pw_135_X
      .Range("AX" & NextRow, "BC" & NextRow + a - 2) = pw_150_X

      .Range("BE" & NextRow, "BJ" & NextRow + a - 2) = pw_180_X
      .Range("BL" & NextRow, "BQ" & NextRow + a - 2) = pw_210_X
      .Range("BS" & NextRow, "BX" & NextRow + a - 2) = pw_225_X
      .Range("BZ" & NextRow, "CE" & NextRow + a - 2) = pw_240_X
      .Range("CG" & NextRow, "CL" & NextRow + a - 2) = pw_270_X
      .Range("CN" & NextRow, "CS" & NextRow + a - 2) = pw_300_X
      .Range("CU" & NextRow, "CZ" & NextRow + a - 2) = pw_315_X
      .Range("DB" & NextRow, "DG" & NextRow + a - 2) = pw_330_X

End If


      'Formats cell values
      lastRow = .Range("A" & Rows.Count).End(xlUp).row
      
      .Range("C" & NextRow, "D" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("E" & NextRow, "F" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("J" & NextRow, "K" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("L" & NextRow, "M" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("Q" & NextRow, "R" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("S" & NextRow, "T" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("X" & NextRow, "Y" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("Z" & NextRow, "AA" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("AE" & NextRow, "AF" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("AG" & NextRow, "AH" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("AL" & NextRow, "AM" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("AN" & NextRow, "AO" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("AS" & NextRow, "AT" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("AU" & NextRow, "AV" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("AZ" & NextRow, "BA" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("BB" & NextRow, "BC" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("BG" & NextRow, "BH" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("BI" & NextRow, "BJ" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("BN" & NextRow, "BO" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("BP" & NextRow, "BQ" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("BU" & NextRow, "BV" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("BW" & NextRow, "BX" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("CB" & NextRow, "CC" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("CD" & NextRow, "CE" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("CI" & NextRow, "CJ" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("CK" & NextRow, "CL" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("CP" & NextRow, "CQ" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("CR" & NextRow, "CS" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("CW" & NextRow, "CX" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("CY" & NextRow, "CZ" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("DD" & NextRow, "DE" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("DF" & NextRow, "DG" & NextRow + a - 2).NumberFormat = "0.0%"
   
   End With    'Worksheets("Distributed Load Tables")
        
'    Call Sort_Distributed

      With Application
         .ScreenUpdating = True
         .Calculation = xlCalculationAutomatic
         .StatusBar = "Ready"
      End With
'   Worksheets("Results").Select
'   Range("A3").Select
   Exit Sub

'ErrorHandle:
'MsgBox "There appears to be an error" & vbNewLine & "with the imported RISA-3D model.", , "Error!"

   With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .StatusBar = "Ready"
   End With
End Sub

Sub Distributed_Loads_IceWind_Z()

    Dim piw_0_Z() As Variant, piw_0_X As Variant        '0º ice distibuted load arrays
    Dim piw_30_Z() As Variant, piw_30_X As Variant      '30º ice distibuted load arrays
    Dim piw_45_Z() As Variant, piw_45_X As Variant      '30º ice distibuted load arrays
    Dim piw_60_Z() As Variant, piw_60_X As Variant      '60º ice distibuted load arrays
    Dim piw_90_Z() As Variant, piw_90_X As Variant      '90º ice distibuted load arrays
    Dim piw_120_Z() As Variant, piw_120_X As Variant    '120º ice distibuted load arrays
    Dim piw_135_Z() As Variant, piw_135_X As Variant    '120º ice distibuted load arrays
    Dim piw_150_Z() As Variant, piw_150_X As Variant    '150º ice distibuted load arrays
    
    Dim piw_180_Z() As Variant, piw_180_X As Variant    '0º ice distibuted load arrays
    Dim piw_210_Z() As Variant, piw_210_X As Variant    '30º ice distibuted load arrays
    Dim piw_225_Z() As Variant, piw_225_X As Variant    '30º ice distibuted load arrays
    Dim piw_240_Z() As Variant, piw_240_X As Variant    '60º ice distibuted load arrays
    Dim piw_270_Z() As Variant, piw_270_X As Variant    '90º ice distibuted load arrays
    Dim piw_300_Z() As Variant, piw_300_X As Variant    '120º ice distibuted load arrays
    Dim piw_315_Z() As Variant, piw_315_X As Variant    '120º ice distibuted load arrays
    Dim piw_330_Z() As Variant, piw_330_X As Variant    '150º ice distibuted load arrays
    
    Dim pid() As Variant                                'Ice dead load array
    
    Dim NextRow As Integer, firstRow As Long
    Dim lastRow As Long, LastRow_Nodes As Long
    
    Dim a As Integer, b As Integer                      'Counter to detmerine size of array
    Dim c As Integer, d As Integer                      'Counter to find array of members
    Dim r As Integer
    
    Dim W As Double                                     'Projected Width of member
    Dim wi As Double                                    'Projected Width of member with ice
    
'    Dim wl As Double, wil As Double                    'Projected Width of member
    Dim EPA As Double, IceEPA As Double                 'Effective Projected Area of member
    Dim Front As Double, IceFront As Double             'Front and Ice Front length of member
    Dim Transverse As Double, IceTransverse As Double   'Transverse and Ice Transverse length of member
    Dim Length As Double, IceLength As Double           'Member lengths
    Dim Azimuth As Double, Orientation As Double        'Mount and member orientations
    Dim Increment As Variant
    Dim WindProjection As Double
    Dim IceWindProjection As Double
    
    Dim Dc As Double                                    'Out-to-out dimension for calculating Ice weight (See TIA-222-H, Figure 2-4)
    Dim Aiz As Double                                   'Cross-sectional area ice
    Dim Ca As Double                                    'Force Coefficient
    Dim p As Double, pice As Double                     'Applied Force
    Dim DiLoad As Double                                'Distibuted Ice Dead Loads
    Dim ConvertLinearForce As Double
    Dim ConvertLength2in As Double, ConvertLength2ft As Double
    Const Pi As Double = 3.141592654
    
    Dim qz As Range
    Dim qiz As Range
    Dim tiz As Range
    Dim Ka As Range
    Dim Gh As Range
    Dim Kesw As Range
    Dim Kesi As Range


    With shCode
    If .Range("RISA3D.Unit.LinearForce") = "pli" Then
        ConvertLinearForce = 1 / 12
        ElseIf .Range("RISA3D.Unit.LinearForce") = "plf" Then
        ConvertLinearForce = 1
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kli" Then
        ConvertLinearForce = 1 / 1000 * 1 / 12
        ElseIf .Range("RISA3D.Unit.LinearForce") = "klf" Then
        ConvertLinearForce = 1 / 1000
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/m" Then
        ConvertLinearForce = 0.0145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/cm" Then
        ConvertLinearForce = 0.000145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/mm" Then
        ConvertLinearForce = 0.0000145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/m" Then
        ConvertLinearForce = 14.5939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/cm" Then
        ConvertLinearForce = 0.145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/mm" Then
        ConvertLinearForce = 0.0145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/m" Then
        ConvertLinearForce = 0.001489173
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/cm" Then
        ConvertLinearForce = 0.0000148917
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/mm" Then
        ConvertLinearForce = 0.00000148917
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/m" Then
        ConvertLinearForce = 1.489173228
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/cm" Then
        ConvertLinearForce = 0.014891732
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/mm" Then
        ConvertLinearForce = 0.001489173
    End If
    
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength2in = 1
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength2in = 12
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength2in = 0.3048
        ElseIf .Range("RISA3D.Unit.Length") = "cm" Then
        ConvertLength2in = 30.48
        ElseIf .Range("RISA3D.Unit.Length") = "mm" Then
        ConvertLength2in = 304.8
    End If
    
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength2ft = 1 / 12
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength2ft = 1
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength2ft = 3.280839895
        ElseIf .Range("RISA3D.Unit.Length") = "cm" Then
        ConvertLength2ft = 0.032808399
        ElseIf .Range("RISA3D.Unit.Length") = "mm" Then
        ConvertLength2ft = 0.00328084
    End If
    End With
    
    Set qz = shCode.Range("qz")
    Set qiz = shCode.Range("qiz")
    Set tiz = shCode.Range("tiz")
    Set Ka = shCode.Range("Ka")
    Set Gh = shCode.Range("Gh")
    Set Kesw = shCode.Range("Kes.w")
    Set Kesi = shCode.Range("Kes.i")

'    On Error GoTo ErrorHandle
    
   With shGeometry
      lastRow = .Range("B6").CurrentRegion.End(xlDown).row
      LastRow_Nodes = .Range("R6").CurrentRegion.End(xlDown).row
      
      For c = 6 To lastRow                                   'Determine size of array by filtering out "RIGID" members
         If .Range("C" & c).Value <> "RIGID" Then
            a = a + 1
         End If
      Next c
      
      If a >= 1 Then
         'Defining the size of arrays
         ReDim piw_0_Z(1 To a, 1 To 6)
         ReDim piw_30_Z(1 To a, 1 To 6)
         ReDim piw_45_Z(1 To a, 1 To 6)
         ReDim piw_60_Z(1 To a, 1 To 6)
         ReDim piw_90_Z(1 To a, 1 To 6)
         ReDim piw_120_Z(1 To a, 1 To 6)
         ReDim piw_135_Z(1 To a, 1 To 6)
         ReDim piw_150_Z(1 To a, 1 To 6)

         ReDim piw_180_Z(1 To a, 1 To 6)
         ReDim piw_210_Z(1 To a, 1 To 6)
         ReDim piw_225_Z(1 To a, 1 To 6)
         ReDim piw_240_Z(1 To a, 1 To 6)
         ReDim piw_270_Z(1 To a, 1 To 6)
         ReDim piw_300_Z(1 To a, 1 To 6)
         ReDim piw_315_Z(1 To a, 1 To 6)
         ReDim piw_330_Z(1 To a, 1 To 6)
         
         
         ReDim pid(1 To a, 1 To 6)
      Else
         MsgBox "Member information must be entered before distributed loads can be calculated.", vbInformation, "Members do not exist"
         Exit Sub
      End If
      
      With Application
         .ScreenUpdating = False
         .Calculation = xlCalculationManual
         .StatusBar = "Calculating distributed loads"
      End With
      
      'Populates arrays
      a = 1
      For c = 6 To lastRow
         If .Range("C" & c) <> "RIGID" Then       'Only runs loop for non-Rigid members. "If" statement could be removed since "importing_members" filters out RIGID Links, but
                                                  'decided to leave incase user prefers to copy/paste members from RISA 3D
         piw_0_Z(a, 1) = .Cells(c, 2)
         piw_30_Z(a, 1) = .Cells(c, 2)
         piw_45_Z(a, 1) = .Cells(c, 2)
         piw_60_Z(a, 1) = .Cells(c, 2)
         piw_90_Z(a, 1) = .Cells(c, 2)
         piw_120_Z(a, 1) = .Cells(c, 2)
         piw_135_Z(a, 1) = .Cells(c, 2)
         piw_150_Z(a, 1) = .Cells(c, 2)

         piw_180_Z(a, 1) = .Cells(c, 2)
         piw_210_Z(a, 1) = .Cells(c, 2)
         piw_225_Z(a, 1) = .Cells(c, 2)
         piw_240_Z(a, 1) = .Cells(c, 2)
         piw_240_Z(a, 1) = .Cells(c, 2)
         piw_270_Z(a, 1) = .Cells(c, 2)
         piw_315_Z(a, 1) = .Cells(c, 2)
         piw_300_Z(a, 1) = .Cells(c, 2)
         piw_330_Z(a, 1) = .Cells(c, 2)

         pid(a, 1) = .Cells(c, 2)
      
         If WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 4, False) = "Flat" Then       'Determining p coefficient based on Round or Flat Profile
            Ca = 2
         ElseIf WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 4, False) = "HSS Flat" Then
            If WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 6, False) Then
                Ca = WorksheetFunction.Max(1.2 - 2.8 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 0.85)
            ElseIf WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 6, False) >= 25 Then
                Ca = WorksheetFunction.Max(2 - 6 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 1.25)
            Else
                Ca = WorksheetFunction.Max(1.4 - 4 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 0.9)
            End If
         Else
            Ca = 1.2
         End If
                
         
         p = qz * Ka * Kesw * Gh      'psf                     'Calculate Forces for NO ICE condition
         pice = qiz * Ka * Kesw * Gh  'psf                     'Caculates p WITH ICE condition. Per Code, can assume Ca=1.2 (round) for projected area for ice conditions

         Dc = WorksheetFunction.VLookup(.Cells(c, 3), .Range("rngR3DSectionSets"), 9, False)
         Aiz = Pi * (tiz * Kesi) * ((Dc + (tiz * Kesi))) / 144 'in^2
         DiLoad = Aiz * 56


'>>>Ice<<<
With shGeometry
         Length = .Cells(c, 15).Value / 12 'in
         j = WorksheetFunction.VLookup(shGeometry.Range("B" & c), shGeometry.Range("E6:K99999"), 6, False) * ConvertLength2ft
         WindProjection = WorksheetFunction.VLookup(.Cells(c, 3).Value, .Range("rngR3DSectionSets"), 5, False) / 12
         Orientation = .Cells(c, 16).Value
         IceWindProjection = WindProjection + 2 * tiz / 12 * Kesi 'ft
         IceLength = Length + 2 * tiz / 12 * Kesi 'ft
         
         Azimuth = 0
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_0_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_0_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_0_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_0_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_0_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_0_Z(a, 5) = "0.0%"
         piw_0_Z(a, 6) = "100.0%"
         
         
         Azimuth = 30
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_30_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_30_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_30_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_30_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_30_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_30_Z(a, 5) = "0.0%"
         piw_30_Z(a, 6) = "100.0%"
         
         
         Azimuth = 45
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_45_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_45_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_45_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_45_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_45_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_45_Z(a, 5) = "0.0%"
         piw_45_Z(a, 6) = "100.0%"

         
         Azimuth = 60
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_60_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_60_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_60_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_60_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_60_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_60_Z(a, 5) = "0.0%"
         piw_60_Z(a, 6) = "100.0%"


         Azimuth = 90
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_90_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_90_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_90_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_90_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_90_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_90_Z(a, 5) = "0.0%"
         piw_90_Z(a, 6) = "100.0%"


         Azimuth = 120
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_120_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_120_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_120_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_120_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_120_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_120_Z(a, 5) = "0.0%"
         piw_120_Z(a, 6) = "100.0%"

         Azimuth = 135
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_135_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_135_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_135_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_135_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_135_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_135_Z(a, 5) = "0.0%"
         piw_135_Z(a, 6) = "100.0%"

         Azimuth = 150
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_150_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_150_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_150_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_150_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_150_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_150_Z(a, 5) = "0.0%"
         piw_150_Z(a, 6) = "100.0%"


         Azimuth = 180
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_180_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_180_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_180_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_180_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_180_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_180_Z(a, 5) = "0.0%"
         piw_180_Z(a, 6) = "100.0%"


         Azimuth = 210
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_210_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_210_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_210_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_210_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_210_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_210_Z(a, 5) = "0.0%"
         piw_210_Z(a, 6) = "100.0%"


         Azimuth = 225
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_225_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_225_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_225_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_225_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_225_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_225_Z(a, 5) = "0.0%"
         piw_225_Z(a, 6) = "100.0%"


         Azimuth = 240
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_240_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_240_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_240_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_240_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_240_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_240_Z(a, 5) = "0.0%"
         piw_240_Z(a, 6) = "100.0%"


         Azimuth = 270
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_270_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_270_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_270_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_270_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_270_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_270_Z(a, 5) = "0.0%"
         piw_270_Z(a, 6) = "100.0%"


         Azimuth = 300
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_300_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_300_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_300_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_300_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_300_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_300_Z(a, 5) = "0.0%"
         piw_300_Z(a, 6) = "100.0%"


         Azimuth = 315
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_315_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_315_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_315_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_315_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_315_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_315_Z(a, 5) = "0.0%"
         piw_315_Z(a, 6) = "100.0%"


         Azimuth = 330
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_330_Z(a, 2) = "Z"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_330_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_330_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_330_Z(a, 3) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_330_Z(a, 4) = -ConvertLinearForce * pice * Cos(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_330_Z(a, 5) = "0.0%"
         piw_330_Z(a, 6) = "100.0%"
End With

'>>>Ice Dead<<<
         pid(a, 2) = "Y"
         pid(a, 3) = -ConvertLinearForce * DiLoad
         pid(a, 4) = -ConvertLinearForce * DiLoad
         pid(a, 5) = "0.0%"
         pid(a, 6) = "100.0%"
         a = a + 1
         End If
         Next c
   End With       'worksheets("Import Model)

     
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
With Worksheets("Distributed Load Tables")
    If .Range("DI4") <> "" Then
    lastRow = .Range("DI4").CurrentRegion.Rows.Count
End If

If a >= 1 Then
NextRow = .Range("DI3").CurrentRegion.Rows.Count + 2

      'Prints array to Loads Tables sheet
      .Range("DI" & NextRow, "DN" & NextRow + a - 2) = piw_0_Z
      .Range("DP" & NextRow, "DU" & NextRow + a - 2) = piw_30_Z
      .Range("DW" & NextRow, "EB" & NextRow + a - 2) = piw_45_Z
      .Range("ED" & NextRow, "EI" & NextRow + a - 2) = piw_60_Z
      .Range("EK" & NextRow, "EP" & NextRow + a - 2) = piw_90_Z
      .Range("ER" & NextRow, "EW" & NextRow + a - 2) = piw_120_Z
      .Range("EY" & NextRow, "FD" & NextRow + a - 2) = piw_135_Z
      .Range("FF" & NextRow, "FK" & NextRow + a - 2) = piw_150_Z

      .Range("FM" & NextRow, "FR" & NextRow + a - 2) = piw_180_Z
      .Range("FT" & NextRow, "FY" & NextRow + a - 2) = piw_210_Z
      .Range("GA" & NextRow, "GF" & NextRow + a - 2) = piw_225_Z
      .Range("GH" & NextRow, "GM" & NextRow + a - 2) = piw_240_Z
      .Range("GO" & NextRow, "GT" & NextRow + a - 2) = piw_270_Z
      .Range("GV" & NextRow, "HA" & NextRow + a - 2) = piw_300_Z
      .Range("HC" & NextRow, "HH" & NextRow + a - 2) = piw_315_Z
      .Range("HJ" & NextRow, "HO" & NextRow + a - 2) = piw_330_Z


      .Range("HQ" & NextRow, "HV" & NextRow + a - 2) = pid

End If


      'Formats cell values
      lastRow = .Range("A" & Rows.Count).End(xlUp).row
      
      .Range("DK" & NextRow, "DL" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("DM" & NextRow, "DN" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("DR" & NextRow, "DS" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("DT" & NextRow, "DU" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("DY" & NextRow, "DZ" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("EA" & NextRow, "EB" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("EF" & NextRow, "EG" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("EH" & NextRow, "EI" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("EM" & NextRow, "EN" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("EO" & NextRow, "EP" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("ET" & NextRow, "EU" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("EV" & NextRow, "EW" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("FA" & NextRow, "FB" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("FC" & NextRow, "FD" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("FH" & NextRow, "FI" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("FJ" & NextRow, "FK" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("FO" & NextRow, "FP" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("FQ" & NextRow, "FR" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("FV" & NextRow, "FW" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("FX" & NextRow, "FY" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("GC" & NextRow, "GD" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("GE" & NextRow, "GF" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("GJ" & NextRow, "GK" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("GL" & NextRow, "GM" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("GQ" & NextRow, "GR" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("GS" & NextRow, "GT" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("GX" & NextRow, "GY" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("GZ" & NextRow, "HA" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("HE" & NextRow, "HF" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("HG" & NextRow, "HH" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("HL" & NextRow, "HM" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("HN" & NextRow, "HO" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("HS" & NextRow, "HT" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("HU" & NextRow, "HV" & NextRow + a - 2).NumberFormat = "0.0%"


   End With    'Worksheets("Distributed Load Tables")
        
'    Call Sort_Distributed

      With Application
         .ScreenUpdating = True
         .Calculation = xlCalculationAutomatic
         .StatusBar = "Ready"
      End With
'   Worksheets("Results").Select
'   Range("A3").Select
   Exit Sub

'ErrorHandle:
'MsgBox "There appears to be an error" & vbNewLine & "with the imported RISA-3D model.", , "Error!"

   With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .StatusBar = "Ready"
   End With
End Sub

Sub Distributed_Loads_IceWind_X()
    
    Dim piw_0_Z() As Variant, piw_0_X As Variant        '0º ice distibuted load arrays
    Dim piw_30_Z() As Variant, piw_30_X As Variant      '30º ice distibuted load arrays
    Dim piw_45_Z() As Variant, piw_45_X As Variant      '30º ice distibuted load arrays
    Dim piw_60_Z() As Variant, piw_60_X As Variant      '60º ice distibuted load arrays
    Dim piw_90_Z() As Variant, piw_90_X As Variant      '90º ice distibuted load arrays
    Dim piw_120_Z() As Variant, piw_120_X As Variant    '120º ice distibuted load arrays
    Dim piw_135_Z() As Variant, piw_135_X As Variant    '120º ice distibuted load arrays
    Dim piw_150_Z() As Variant, piw_150_X As Variant    '150º ice distibuted load arrays
    
    Dim piw_180_Z() As Variant, piw_180_X As Variant    '0º ice distibuted load arrays
    Dim piw_210_Z() As Variant, piw_210_X As Variant    '30º ice distibuted load arrays
    Dim piw_225_Z() As Variant, piw_225_X As Variant    '30º ice distibuted load arrays
    Dim piw_240_Z() As Variant, piw_240_X As Variant    '60º ice distibuted load arrays
    Dim piw_270_Z() As Variant, piw_270_X As Variant    '90º ice distibuted load arrays
    Dim piw_300_Z() As Variant, piw_300_X As Variant    '120º ice distibuted load arrays
    Dim piw_315_Z() As Variant, piw_315_X As Variant    '120º ice distibuted load arrays
    Dim piw_330_Z() As Variant, piw_330_X As Variant    '150º ice distibuted load arrays
    
    Dim pid() As Variant                                'Ice dead load array
    
    Dim NextRow As Integer, firstRow As Long
    Dim lastRow As Long, LastRow_Nodes As Long
    
    Dim a As Integer, b As Integer                      'Counter to detmerine size of array
    Dim c As Integer, d As Integer                      'Counter to find array of members
    Dim r As Integer
    
    Dim W As Double                                     'Projected Width of member
    Dim wi As Double                                    'Projected Width of member with ice
    
'    Dim wl As Double, wil As Double                    'Projected Width of member
    Dim EPA As Double, IceEPA As Double                 'Effective Projected Area of member
    Dim Front As Double, IceFront As Double             'Front and Ice Front length of member
    Dim Transverse As Double, IceTransverse As Double   'Transverse and Ice Transverse length of member
    Dim Length As Double, IceLength As Double           'Member lengths
    Dim Azimuth As Double, Orientation As Double        'Mount and member orientations
    Dim Increment As Variant
    Dim WindProjection As Double
    Dim IceWindProjection As Double
    
    Dim Dc As Double                                    'Out-to-out dimension for calculating Ice weight (See TIA-222-H, Figure 2-4)
    Dim Aiz As Double                                   'Cross-sectional area ice
    Dim Ca As Double                                    'Force Coefficient
    Dim p As Double, pice As Double                     'Applied Force
    Dim DiLoad As Double                                'Distibuted Ice Dead Loads
    Dim ConvertLinearForce As Double
    Dim ConvertLength2in As Double, ConvertLength2ft As Double
    Const Pi As Double = 3.141592654
    
    Dim qz As Range
    Dim qiz As Range
    Dim tiz As Range
    Dim Ka As Range
    Dim Gh As Range
    Dim Kesw As Range
    Dim Kesi As Range


    With shCode
    If .Range("RISA3D.Unit.LinearForce") = "pli" Then
        ConvertLinearForce = 1 / 12
        ElseIf .Range("RISA3D.Unit.LinearForce") = "plf" Then
        ConvertLinearForce = 1
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kli" Then
        ConvertLinearForce = 1 / 1000 * 1 / 12
        ElseIf .Range("RISA3D.Unit.LinearForce") = "klf" Then
        ConvertLinearForce = 1 / 1000
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/m" Then
        ConvertLinearForce = 0.0145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/cm" Then
        ConvertLinearForce = 0.000145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kN/mm" Then
        ConvertLinearForce = 0.0000145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/m" Then
        ConvertLinearForce = 14.5939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/cm" Then
        ConvertLinearForce = 0.145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "N/mm" Then
        ConvertLinearForce = 0.0145939042
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/m" Then
        ConvertLinearForce = 0.001489173
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/cm" Then
        ConvertLinearForce = 0.0000148917
        ElseIf .Range("RISA3D.Unit.LinearForce") = "mt/mm" Then
        ConvertLinearForce = 0.00000148917
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/m" Then
        ConvertLinearForce = 1.489173228
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/cm" Then
        ConvertLinearForce = 0.014891732
        ElseIf .Range("RISA3D.Unit.LinearForce") = "kg/mm" Then
        ConvertLinearForce = 0.001489173
    End If
    
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength2in = 1
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength2in = 12
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength2in = 0.3048
        ElseIf .Range("RISA3D.Unit.Length") = "cm" Then
        ConvertLength2in = 30.48
        ElseIf .Range("RISA3D.Unit.Length") = "mm" Then
        ConvertLength2in = 304.8
    End If
    
    If .Range("RISA3D.Unit.Length") = "in" Then
        ConvertLength2ft = 1 / 12
        ElseIf .Range("RISA3D.Unit.Length") = "ft" Then
        ConvertLength2ft = 1
        ElseIf .Range("RISA3D.Unit.Length") = "m" Then
        ConvertLength2ft = 3.280839895
        ElseIf .Range("RISA3D.Unit.Length") = "cm" Then
        ConvertLength2ft = 0.032808399
        ElseIf .Range("RISA3D.Unit.Length") = "mm" Then
        ConvertLength2ft = 0.00328084
    End If
    End With
    
    Set qz = shCode.Range("qz")
    Set qiz = shCode.Range("qiz")
    Set tiz = shCode.Range("tiz")
    Set Ka = shCode.Range("Ka")
    Set Gh = shCode.Range("Gh")
    Set Kesw = shCode.Range("Kes.w")
    Set Kesi = shCode.Range("Kes.i")

'    On Error GoTo ErrorHandle
    
   With shGeometry
      lastRow = .Range("B6").CurrentRegion.End(xlDown).row
      LastRow_Nodes = .Range("R6").CurrentRegion.End(xlDown).row
      
      For c = 6 To lastRow                                   'Determine size of array by filtering out "RIGID" members
         If .Range("C" & c).Value <> "RIGID" Then
            a = a + 1
         End If
      Next c
      
      If a >= 1 Then
         'Defining the size of arrays
         ReDim piw_0_X(1 To a, 1 To 6)
         ReDim piw_30_X(1 To a, 1 To 6)
         ReDim piw_45_X(1 To a, 1 To 6)
         ReDim piw_60_X(1 To a, 1 To 6)
         ReDim piw_90_X(1 To a, 1 To 6)
         ReDim piw_120_X(1 To a, 1 To 6)
         ReDim piw_135_X(1 To a, 1 To 6)
         ReDim piw_150_X(1 To a, 1 To 6)

         ReDim piw_180_X(1 To a, 1 To 6)
         ReDim piw_210_X(1 To a, 1 To 6)
         ReDim piw_225_X(1 To a, 1 To 6)
         ReDim piw_240_X(1 To a, 1 To 6)
         ReDim piw_270_X(1 To a, 1 To 6)
         ReDim piw_300_X(1 To a, 1 To 6)
         ReDim piw_315_X(1 To a, 1 To 6)
         ReDim piw_330_X(1 To a, 1 To 6)
         
         
         ReDim pid(1 To a, 1 To 6)
      Else
         MsgBox "Member information must be entered before distributed loads can be calculated.", vbInformation, "Members do not exist"
         Exit Sub
      End If
      
      With Application
         .ScreenUpdating = False
         .Calculation = xlCalculationManual
         .StatusBar = "Calculating distributed loads"
      End With
      
      'Populates arrays
      a = 1
      For c = 6 To lastRow
         If .Range("C" & c) <> "RIGID" Then       'Only runs loop for non-Rigid members. "If" statement could be removed since "importing_members" filters out RIGID Links, but
                                                  'decided to leave incase user prefers to copy/paste members from RISA 3D
         piw_0_X(a, 1) = .Cells(c, 2)
         piw_30_X(a, 1) = .Cells(c, 2)
         piw_45_X(a, 1) = .Cells(c, 2)
         piw_60_X(a, 1) = .Cells(c, 2)
         piw_90_X(a, 1) = .Cells(c, 2)
         piw_120_X(a, 1) = .Cells(c, 2)
         piw_135_X(a, 1) = .Cells(c, 2)
         piw_150_X(a, 1) = .Cells(c, 2)

         piw_180_X(a, 1) = .Cells(c, 2)
         piw_210_X(a, 1) = .Cells(c, 2)
         piw_225_X(a, 1) = .Cells(c, 2)
         piw_240_X(a, 1) = .Cells(c, 2)
         piw_240_X(a, 1) = .Cells(c, 2)
         piw_270_X(a, 1) = .Cells(c, 2)
         piw_315_X(a, 1) = .Cells(c, 2)
         piw_300_X(a, 1) = .Cells(c, 2)
         piw_330_X(a, 1) = .Cells(c, 2)

         pid(a, 1) = .Cells(c, 2)
      
         If WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 4, False) = "Flat" Then       'Determining p coefficient based on Round or Flat Profile
            Ca = 2
         ElseIf WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 4, False) = "HSS Flat" Then
            If WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 6, False) Then
                Ca = WorksheetFunction.Max(1.2 - 2.8 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 0.85)
            ElseIf WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 6, False) >= 25 Then
                Ca = WorksheetFunction.Max(2 - 6 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 1.25)
            Else
                Ca = WorksheetFunction.Max(1.4 - 4 * 1.5 * WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 7, False) / WorksheetFunction.VLookup(.Range("C" & c), .Range("rngR3DSectionSets"), 5, False), 0.9)
            End If
         Else
            Ca = 1.2
         End If
                
         
         p = qz * Ka * Kesw * Gh      'psf                     'Calculate Forces for NO ICE condition
         pice = qiz * Ka * Kesw * Gh  'psf                     'Caculates p WITH ICE condition. Per Code, can assume Ca=1.2 (round) for projected area for ice conditions

         Dc = WorksheetFunction.VLookup(.Cells(c, 3), .Range("rngR3DSectionSets"), 9, False)
         Aiz = Pi * (tiz * Kesi) * ((Dc + (tiz * Kesi))) / 144 'in^2
         DiLoad = Aiz * 56


'>>>Ice<<<
With shGeometry
         Length = .Cells(c, 15).Value / 12 'in
         j = WorksheetFunction.VLookup(shGeometry.Range("B" & c), shGeometry.Range("E6:K99999"), 6, False) * ConvertLength2ft
         WindProjection = WorksheetFunction.VLookup(.Cells(c, 3).Value, .Range("rngR3DSectionSets"), 5, False) / 12
         Orientation = .Cells(c, 16).Value
         IceWindProjection = WindProjection + 2 * tiz / 12 * Kesi 'ft
         IceLength = Length + 2 * tiz / 12 * Kesi 'ft
         
         Azimuth = 0
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_0_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_0_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_0_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_0_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_0_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_0_X(a, 5) = "0.0%"
         piw_0_X(a, 6) = "100.0%"
         
         
         Azimuth = 30
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_30_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_30_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_30_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_30_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_30_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_30_X(a, 5) = "0.0%"
         piw_30_X(a, 6) = "100.0%"
         
         
         Azimuth = 45
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_45_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_45_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_45_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_45_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_45_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_45_X(a, 5) = "0.0%"
         piw_45_X(a, 6) = "100.0%"

         
         Azimuth = 60
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_60_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_60_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_60_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_60_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_60_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_60_X(a, 5) = "0.0%"
         piw_60_X(a, 6) = "100.0%"


         Azimuth = 90
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_90_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_90_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_90_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_90_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_90_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_90_X(a, 5) = "0.0%"
         piw_90_X(a, 6) = "100.0%"


         Azimuth = 120
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_120_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_120_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_120_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_120_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_120_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_120_X(a, 5) = "0.0%"
         piw_120_X(a, 6) = "100.0%"

         Azimuth = 135
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_135_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_135_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_135_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_135_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_135_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_135_X(a, 5) = "0.0%"
         piw_135_X(a, 6) = "100.0%"

         Azimuth = 150
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_150_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_150_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_150_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_150_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_150_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_150_X(a, 5) = "0.0%"
         piw_150_X(a, 6) = "100.0%"


         Azimuth = 180
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_180_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_180_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_180_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_180_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_180_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_180_X(a, 5) = "0.0%"
         piw_180_X(a, 6) = "100.0%"


         Azimuth = 210
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_210_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_210_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_210_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_210_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_210_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_210_X(a, 5) = "0.0%"
         piw_210_X(a, 6) = "100.0%"


         Azimuth = 225
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_225_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_225_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_225_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_225_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_225_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_225_X(a, 5) = "0.0%"
         piw_225_X(a, 6) = "100.0%"


         Azimuth = 240
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_240_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_240_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_240_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_240_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_240_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_240_X(a, 5) = "0.0%"
         piw_240_X(a, 6) = "100.0%"


         Azimuth = 270
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_270_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_270_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_270_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_270_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_270_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_270_X(a, 5) = "0.0%"
         piw_270_X(a, 6) = "100.0%"


         Azimuth = 300
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_300_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_300_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_300_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_300_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_300_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_300_X(a, 5) = "0.0%"
         piw_300_X(a, 6) = "100.0%"


         Azimuth = 315
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_315_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_315_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_315_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_315_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_315_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_315_X(a, 5) = "0.0%"
         piw_315_X(a, 6) = "100.0%"


         Azimuth = 330
         Front = WorksheetFunction.IfError(Length * Cos((Azimuth + Orientation) * Pi / 180) ^ 2, 0)
         If j = 0 Then
            Transverse = 0
         ElseIf Abs(j) > 0 And Orientation = 0 Or Orientation > 0 Then
            Transverse = (Abs(j) * Sin(Pi / 180 * (Orientation + Azimuth)) ^ 2)
         Else
            Transverse = 0
         End If
         IceFront = Front + 2 * tiz / 12 * Kesi
         IceTransverse = Transverse + 2 * tiz / 12 * Kesi
         EPA = WorksheetFunction.IfError((Front + Transverse) * WindProjection * Ca, 0)
         IceEPA = WorksheetFunction.IfError((IceLength * IceWindProjection - Front * WindProjection) * 1.2, 0)

         piw_330_X(a, 2) = "X"
         If WorksheetFunction.VLookup(.Range("M" & c), .Range("S6:V99999"), 3, False) = WorksheetFunction.VLookup(.Range("N" & c), .Range("S6:V99999"), 3, False) Then
            piw_330_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
            piw_330_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * (((((IceLength * IceWindProjection) - (Length * WindProjection)) * 1.2) * Cos(Pi / 180 * Orientation) ^ 2) + EPA) / Length
         Else
            piw_330_X(a, 3) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
            piw_330_X(a, 4) = ConvertLinearForce * pice * Sin(Pi / 180 * Azimuth) * ((((IceFront * IceWindProjection - Front * WindProjection) + (IceTransverse * IceWindProjection - Transverse * WindProjection)) * 1.2) + EPA) / Length
         End If
         piw_330_X(a, 5) = "0.0%"
         piw_330_X(a, 6) = "100.0%"
End With

         a = a + 1
         End If
         Next c
   End With       'worksheets("Import Model)

     
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
With Worksheets("Distributed Load Tables")
    If .Range("DI4") <> "" Then
    lastRow = .Range("DI4").CurrentRegion.Rows.Count
'    .Range("DI4", "DI" & Lastrow).EntireRow.Delete
    End If

If a >= 1 Then
NextRow = .Range("DI3").CurrentRegion.Rows.Count + 2

      'Prints array to Loads Tables sheet
      .Range("DI" & NextRow, "DN" & NextRow + a - 2) = piw_0_X
      .Range("DP" & NextRow, "DU" & NextRow + a - 2) = piw_30_X
      .Range("DW" & NextRow, "EB" & NextRow + a - 2) = piw_45_X
      .Range("ED" & NextRow, "EI" & NextRow + a - 2) = piw_60_X
      .Range("EK" & NextRow, "EP" & NextRow + a - 2) = piw_90_X
      .Range("ER" & NextRow, "EW" & NextRow + a - 2) = piw_120_X
      .Range("EY" & NextRow, "FD" & NextRow + a - 2) = piw_135_X
      .Range("FF" & NextRow, "FK" & NextRow + a - 2) = piw_150_X

      .Range("FM" & NextRow, "FR" & NextRow + a - 2) = piw_180_X
      .Range("FT" & NextRow, "FY" & NextRow + a - 2) = piw_210_X
      .Range("GA" & NextRow, "GF" & NextRow + a - 2) = piw_225_X
      .Range("GH" & NextRow, "GM" & NextRow + a - 2) = piw_240_X
      .Range("GO" & NextRow, "GT" & NextRow + a - 2) = piw_270_X
      .Range("GV" & NextRow, "HA" & NextRow + a - 2) = piw_300_X
      .Range("HC" & NextRow, "HH" & NextRow + a - 2) = piw_315_X
      .Range("HJ" & NextRow, "HO" & NextRow + a - 2) = piw_330_X

End If


      'Formats cell values
      lastRow = .Range("A" & Rows.Count).End(xlUp).row
      
      .Range("DK" & NextRow, "DL" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("DM" & NextRow, "DN" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("DR" & NextRow, "DS" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("DT" & NextRow, "DU" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("DY" & NextRow, "DZ" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("EA" & NextRow, "EB" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("EF" & NextRow, "EG" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("EH" & NextRow, "EI" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("EM" & NextRow, "EN" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("EO" & NextRow, "EP" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("ET" & NextRow, "EU" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("EV" & NextRow, "EW" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("FA" & NextRow, "FB" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("FC" & NextRow, "FD" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("FH" & NextRow, "FI" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("FJ" & NextRow, "FK" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("FO" & NextRow, "FP" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("FQ" & NextRow, "FR" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("FV" & NextRow, "FW" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("FX" & NextRow, "FY" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("GC" & NextRow, "GD" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("GE" & NextRow, "GF" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("GJ" & NextRow, "GK" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("GL" & NextRow, "GM" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("GQ" & NextRow, "GR" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("GS" & NextRow, "GT" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("GX" & NextRow, "GY" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("GZ" & NextRow, "HA" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("HE" & NextRow, "HF" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("HG" & NextRow, "HH" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("HL" & NextRow, "HM" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("HN" & NextRow, "HO" & NextRow + a - 2).NumberFormat = "0.0%"

      .Range("HS" & NextRow, "HT" & NextRow + a - 2).NumberFormat = "0.000"
      .Range("HU" & NextRow, "HV" & NextRow + a - 2).NumberFormat = "0.0%"


   End With    'Worksheets("Distributed Load Tables")
        
'    Call Sort_Distributed

      With Application
         .ScreenUpdating = True
         .Calculation = xlCalculationAutomatic
         .StatusBar = "Ready"
      End With
'   Worksheets("Results").Select
'   Range("A3").Select
   Exit Sub

'ErrorHandle:
'MsgBox "There appears to be an error" & vbNewLine & "with the imported RISA-3D model.", , "Error!"

   With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .StatusBar = "Ready"
   End With
End Sub

Sub Goto_Mount()
   Worksheets("Results").Select
   Application.GoTo shGeometry.Range("A91"), True
End Sub

Sub Sort_Distributed()

    ActiveWorkbook.Worksheets("Distributed Load Tables").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Distributed Load Tables").Sort.SortFields.Add2 key:= _
        Range("A4:A1048576"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Distributed Load Tables").Sort
        .SetRange Range("A4:CE1048576")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ActiveWorkbook.Worksheets("Distributed Load Tables").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Distributed Load Tables").Sort.SortFields.Add2 key:= _
        Range("CG4:CG1048576"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Distributed Load Tables").Sort
        .SetRange Range("CG4:CL1048576")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

