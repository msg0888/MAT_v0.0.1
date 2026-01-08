Attribute VB_Name = "Loads_Area"
Option Explicit
Sub Area_Loads()
    
    Dim TotalRows As Integer                                    'Total Rows to loop though to determine how many pieces of equipment exist
    Dim lastRow As Integer, NextRow As Integer
    Dim Area() As Variant, iArea() As Variant              'Front Antenna Top & Non-Antenna arrays
    Dim a As Long, b As Long, r As Long     'Various counter and looping variables
    Dim X As Long, Y As Long, Z As Long
    Dim h As Long, t As Long, k As Long, d As Long
    Dim Start As Double
    
     With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .StatusBar = "Calculating..."
     End With

'    On Error GoTo Handle
    
    Start = Timer
    
      With shMaintenance_Loads
         TotalRows = .Range("L" & Rows.Count).End(xlUp).row
         
         For t = 7 To 14
               a = a + 1
         Next t
         If a >= 1 Then
               ReDim Area(1 To a, 1 To 7)
               ReDim iArea(1 To a, 1 To 7)
         Else
            MsgBox "No maintenance loads have been added.", vbInformation, "Please update maintenance loads tables in the Import Model tab"
            With Application
               .ScreenUpdating = True
               .Calculation = xlCalculationAutomatic
               .StatusBar = "Ready"
            End With
            Exit Sub
         End If
                 
         If shCode.Range("QtySectors") = 3 Then
         X = 1
         For r = 7 To 9
            If .Cells(r, 12) <> "" Then
            
               Area(X, 1) = .Cells(r, 12)
               Area(X, 2) = .Cells(r, 13)
               Area(X, 3) = .Cells(r, 14)
               Area(X, 4) = .Cells(r, 15)
               Area(X, 5) = .Cells(r, 16)
               Area(X, 6) = .Cells(r, 17)
               Area(X, 7) = .Cells(r, 18).Value
               
               X = X + 1
            End If
         Next r
         Else
         X = 1
         For r = 7 To 10
            If .Cells(r, 13) <> "" Then
            
               Area(X, 1) = .Cells(r, 12)
               Area(X, 2) = .Cells(r, 13)
               Area(X, 3) = .Cells(r, 14)
               Area(X, 4) = .Cells(r, 15)
               Area(X, 5) = .Cells(r, 16)
               Area(X, 6) = .Cells(r, 17)
               Area(X, 7) = .Cells(r, 18).Value
               
               X = X + 1
            End If
         Next r
         End If
         
         If shCode.Range("QtySectors") = 3 Then
         X = 1
         For r = 11 To 13
            If .Cells(r, 13) <> "" Then

               iArea(X, 1) = .Cells(r, 12)
               iArea(X, 2) = .Cells(r, 13)
               iArea(X, 3) = .Cells(r, 14)
               iArea(X, 4) = .Cells(r, 15)
               iArea(X, 5) = .Cells(r, 16)
               iArea(X, 6) = .Cells(r, 17)
               iArea(X, 7) = .Cells(r, 18).Value

               X = X + 1
            End If
         Next r
         Else
         X = 1
         For r = 11 To 14
            If .Cells(r, 13) <> "" Then

               iArea(X, 1) = .Cells(r, 12)
               iArea(X, 2) = .Cells(r, 13)
               iArea(X, 3) = .Cells(r, 14)
               iArea(X, 4) = .Cells(r, 15)
               iArea(X, 5) = .Cells(r, 16)
               iArea(X, 6) = .Cells(r, 17)
               iArea(X, 7) = .Cells(r, 18).Value

               X = X + 1
            End If
         Next r
         End If
         

      End With
      
        'Deletes rows if already populated before copying new arrays
      With shAreaLoadTables
         If .Range("A4") <> "" Then
            lastRow = .Range("A4").CurrentRegion.Rows.Count
            .Range("A4", "A" & lastRow).EntireRow.ClearContents
         End If
         
         
         'Prints array to "Loads Tables" sheet
         If a >= 1 Then
            NextRow = .Range("A3").CurrentRegion.Rows.Count + 1
            .Range("A" & NextRow, "G" & NextRow + a - 1).Value = Area
            .Range("I" & NextRow, "O" & NextRow + a - 1).Value = iArea

         End If
         
         
         'Formats cell values
         lastRow = .Range("A3").CurrentRegion.Rows.Count
         .Range("G4", "G" & lastRow).NumberFormat = "0.0000"
         .Range("O4", "O" & lastRow).NumberFormat = "0.0000"

      End With   'Worksheets("Area Load Tables")
    
    With Application
        .ScreenUpdating = True
        .StatusBar = "Ready"
        .Calculation = xlCalculationAutomatic
    End With
        
'    Worksheets("Area Load Tables").Activate
    
    Start = Timer - Start
    Debug.Print Start
    
Exit Sub
   
'Handle:
'   AppActivate Application.Caption
'    MsgBox "An error has occured. It appears that one of the tables has " _
'    & "incomplete information. Please ensure that all Mount Pipes and " _
'    & "Equipment information is entered correctly and try again.", vbCritical, Title:="Error!"
'    Application.Calculation = xlCalculationAutomatic

End Sub
