Attribute VB_Name = "frm_Support"
Attribute VB_Base = "0{2106B0F3-BDFC-4EEC-B6B9-E408E236F1D8}{D7F44E77-295B-4C4E-A921-A60A251CDC53}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CommandButton1_Click()
'    Call Macro_Send_Email
Dim eApp As Outlook.Application
Dim eSource As String
Set eApp = New Outlook.Application
Dim eItem As Outlook.MailItem
Set eItem = eApp.CreateItem(olMailItem)
eItem.To = "mgirgis@nbcllc.com"
'These items are optional
'eItem.CC = "glucas@nbcllc.com; mshufflebarger@nbcllc.com; jmarshalek@nbcllc.com"
'etem.BCC = "y@gmail.com"
eItem.Subject = "NB+C MAT - Support"
eItem.Body = "Please save the tool and attach to this email." & vbNewLine & vbNewLine & "Do not leave fields blank." _
& vbNewLine & vbNewLine & "Error Type:  " & ErrorType & vbNewLine & "Sheets:  " & Sheets & vbNewLine & "Buttons:  " & Buttons & _
vbNewLine & vbNewLine & "Description of Error:  " & Description
'If you want to attach this workbook, then uncomment these two lines from below
Source = ThisWorkbook.FullName
ThisWorkbook.Save
eItem.Attachments.Add Source
eItem.Display 'can use .Send

End Sub

Private Sub CommandButton2_Click()
    Exit Sub
End Sub

Private Sub ErrorType_DropbuttonClick()

If Me.ErrorType.ListCount = 0 Then

    With Me.ErrorType
        .AddItem "VBA Error Code (states VBA in the header)"
        .AddItem "Mount Analysis Tool Error"
        .AddItem "RISA-3D Error"
        .AddItem "Formula / Typographical / Formatting"
        .AddItem "Other"
        .AddItem "Unknown"
        .AddItem "N/A"
    End With

End If
End Sub

Private Sub Buttons_DropbuttonClick()

If Me.Buttons.ListCount = 0 Then

    With Me.Buttons
        .AddItem "Yes"
        .AddItem "No"
        .AddItem "Unknown"
        .AddItem "N/A"
    End With
    
End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Sheets_DropbuttonClick()

If Me.Sheets.ListCount = 0 Then
    
    With Sheets
        .AddItem "Code"
        .AddItem "Geometry"
        .AddItem "Discrete Loads"
        .AddItem "Dish Loads"
        .AddItem "Maintenance Loads"
        .AddItem "RISA-3D"
        .AddItem "CodeCheck_Batch"
        .AddItem "MemberEndRxns_Batch"
        .AddItem "NodeReactions_Batch"
        .AddItem "Placement Diagrams"
        .AddItem "Unknown"
        .AddItem "N/A"
    End With

End If
End Sub

Private Sub Troubleshoot_NoFile_Click()
Dim eApp As Outlook.Application
Dim eSource As String
Set eApp = New Outlook.Application
Dim eItem As Outlook.MailItem
Set eItem = eApp.CreateItem(olMailItem)
eItem.To = "mgirgis@nbcllc.com"
'These items are optional
'eItem.CC = "glucas@nbcllc.com; mshufflebarger@nbcllc.com; jmarshalek@nbcllc.com"
'etem.BCC = "y@gmail.com"
eItem.Subject = "NB+C MAT v3 - Support"
eItem.Body = "Please save the tool and attach to this email." & vbNewLine & vbNewLine & "Do not leave fields blank." _
& vbNewLine & vbNewLine & "Error Type:  " & ErrorType & vbNewLine & "Sheets:  " & Sheets & vbNewLine & "Buttons:  " & Buttons & _
vbNewLine & vbNewLine & "Description of Error:  " & Description
'If you want to attach this workbook, then uncomment these two lines from below
Source = ThisWorkbook.FullName
'ThisWorkbook.Save
'eItem.Attachments.Add Source
eItem.Display 'can use .Send
End Sub

Sub UserForm_Initialize()
    
    
End Sub

