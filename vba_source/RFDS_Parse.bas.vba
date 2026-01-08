Attribute VB_Name = "RFDS_Parse"
Option Explicit

'=========================================================
' ENDPOINTS (NO &filename=... here — VBA will append it)
'=========================================================
Private Const ENDPOINT_VZW As String = _
"https://rfds-parser-vzw-fzf2gja2cuc4gqcw.eastus2-01.azurewebsites.net/api/process_pdf?code=REDACTEDXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

Private Const ENDPOINT_TMO As String = _
"https://rfds-parser-vzw-fzf2gja2cuc4gqcw.eastus2-01.azurewebsites.net/api/process_pdf_tmob?code=REDACTEDXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

Private Const ENDPOINT_ATT As String = _
"https://rfds-parser-vzw-fzf2gja2cuc4gqcw.eastus2-01.azurewebsites.net/api/process_pdf_att?code=REDACTEDXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

'=========================================================
' APP STATE for safe fast-mode toggling
'=========================================================
Private Type AppState
    ScreenUpdating As Boolean
    Calculation As XlCalculation
    EnableEvents As Boolean
    DisplayStatusBar As Boolean
End Type

'=========================================================
' PUBLIC ENTRY POINT (call this with "VZW" / "TMO" / "ATT")
'=========================================================
Public Sub RFDS_RunForCarrier(ByVal carrier As String)
    Dim st As AppState
    BeginFastMode st

    On Error GoTo CleanFail

    carrier = UCase$(Trim$(carrier))

    Dim pdfPath As String
    pdfPath = PickPdfFile()
    If Len(pdfPath) = 0 Then GoTo CleanExit

    Dim endpoint As String
    endpoint = GetEndpointForCarrier(carrier)
    If Len(endpoint) = 0 Then
        Err.Raise vbObjectError + 100, "RFDS_RunForCarrier", "No endpoint configured for carrier: " & carrier
    End If

    Dim fileName As String
    fileName = GetFileNameFromPath(pdfPath)

    Dim url As String
    url = BuildUrlWithFilename(endpoint, fileName)

    Dim pdfBytes() As Byte
    pdfBytes = ReadAllBytes(pdfPath)

    Dim xlsxBytes() As Byte
    xlsxBytes = HttpPostPdf_ReturnBytes(url, pdfBytes)

    Dim outPath As String
    outPath = Replace(pdfPath, ".pdf", "", 1, -1, vbTextCompare) & "_" & carrier & "_Report.xlsx"
    WriteAllBytes outPath, xlsxBytes

    ' Open generated workbook (ActiveWorkbook becomes the output file)
    Workbooks.Open outPath

    ' Run your existing carrier-specific import macros
    RunCarrierImport carrier

CleanExit:
    EndFastMode st
    Exit Sub

CleanFail:
    EndFastMode st
    MsgBox "RFDS failed:" & vbCrLf & Err.Description, vbExclamation
End Sub

'=========================================================
' ROUTING: Carrier -> Endpoint
'=========================================================
Private Function GetEndpointForCarrier(ByVal carrier As String) As String
    Select Case UCase$(carrier)
        Case "VZW"
            GetEndpointForCarrier = ENDPOINT_VZW
        Case "TMO"
            GetEndpointForCarrier = ENDPOINT_TMO
        Case "ATT"
            GetEndpointForCarrier = ENDPOINT_ATT
        Case Else
            GetEndpointForCarrier = ""
    End Select
End Function

'=========================================================
' ROUTING: Carrier -> Your import macro
'=========================================================
Private Sub RunCarrierImport(ByVal carrier As String)
    Select Case UCase$(carrier)
        Case "TMO"
            ImportRFDS_TMO
        Case "ATT"
            ImportRFDS_ATT
        Case "VZW"
            ImportRFDS_VZW
        Case Else
            Err.Raise vbObjectError + 101, "RunCarrierImport", "Unknown carrier: " & carrier
    End Select
End Sub

'=========================================================
' URL builder: safely appends filename (and strips any existing filename=)
'=========================================================
Private Function BuildUrlWithFilename(ByVal endpoint As String, ByVal fileName As String) As String
    Dim url As String
    url = Trim$(endpoint)

    ' Remove accidental wrapping quotes
    If Left$(url, 1) = """" And Right$(url, 1) = """" Then
        url = Mid$(url, 2, Len(url) - 2)
    End If

    ' If someone pasted an endpoint that already has filename=..., strip it
    Dim p As Long
    p = InStr(1, url, "&filename=", vbTextCompare)
    If p > 0 Then url = Left$(url, p - 1)

    p = InStr(1, url, "?filename=", vbTextCompare)
    If p > 0 Then url = Left$(url, p - 1)

    ' Validate protocol
    If LCase$(Left$(url, 8)) <> "https://" And LCase$(Left$(url, 7)) <> "http://" Then
        Err.Raise vbObjectError + 102, "BuildUrlWithFilename", "Endpoint must start with http:// or https://"
    End If

    If InStr(1, url, "?", vbTextCompare) > 0 Then
        url = url & "&filename=" & UrlEncode(fileName)
    Else
        url = url & "?filename=" & UrlEncode(fileName)
    End If

    BuildUrlWithFilename = url
End Function

'=========================================================
' UI: Pick a PDF
'=========================================================
Private Function PickPdfFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .title = "Select RFDS PDF"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "PDF Files", "*.pdf"
        If .Show <> -1 Then
            PickPdfFile = ""
        Else
            PickPdfFile = .SelectedItems(1)
        End If
    End With
End Function

'=========================================================
' HTTP: POST raw PDF bytes -> returns XLSX bytes
'=========================================================
Private Function HttpPostPdf_ReturnBytes(ByVal url As String, ByRef bodyBytes() As Byte) As Byte()
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' resolve, connect, send, receive (ms)
    http.SetTimeouts 30000, 30000, 30000, 300000

    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/pdf"
    http.Send bodyBytes

    If http.Status < 200 Or http.Status > 299 Then
        Dim errText As String
        On Error Resume Next
        errText = http.ResponseText
        On Error GoTo 0

        Err.Raise vbObjectError + 103, "HttpPostPdf_ReturnBytes", _
            "HTTP " & http.Status & " " & http.StatusText & vbCrLf & errText
    End If

    HttpPostPdf_ReturnBytes = http.ResponseBody
End Function

'=========================================================
' Fast-mode toggles (restores your prior state)
'=========================================================
Private Sub BeginFastMode(ByRef st As AppState)
    With Application
        st.ScreenUpdating = .ScreenUpdating
        st.Calculation = .Calculation
        st.EnableEvents = .EnableEvents
        st.DisplayStatusBar = .DisplayStatusBar

        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayStatusBar = False
    End With
End Sub

Private Sub EndFastMode(ByRef st As AppState)
    With Application
        .ScreenUpdating = st.ScreenUpdating
        .Calculation = st.Calculation
        .EnableEvents = st.EnableEvents
        .DisplayStatusBar = st.DisplayStatusBar
    End With
End Sub

'=========================================================
' Byte IO
'=========================================================
Private Function ReadAllBytes(ByVal path As String) As Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.LoadFromFile path
    ReadAllBytes = stm.Read
    stm.Close
End Function

Private Sub WriteAllBytes(ByVal path As String, ByRef bytes() As Byte)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.Write bytes
    stm.SaveToFile path, 2 ' overwrite
    stm.Close
End Sub

'=========================================================
' Utilities
'=========================================================
Private Function GetFileNameFromPath(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        GetFileNameFromPath = Mid$(fullPath, p + 1)
    Else
        GetFileNameFromPath = fullPath
    End If
End Function

Private Function UrlEncode(ByVal s As String) As String
    Dim i As Long, ch As Integer, out As String
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        Select Case ch
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126 '0-9 A-Z a-z - . _ ~
                out = out & ChrW(ch)
            Case 32
                out = out & "%20"
            Case Else
                out = out & "%" & Right$("0" & Hex$(ch), 2)
        End Select
    Next i
    UrlEncode = out
End Function

