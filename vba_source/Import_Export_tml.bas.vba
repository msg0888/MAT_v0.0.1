Attribute VB_Name = "Import_Export_tml"
Sub Show_ExportForm()
    frm_ExportEPA.Show
End Sub

Sub Export_Loading_tnxTower()
    Dim savePath As Variant
    Dim xmlText  As String
    Dim Answer As String
    
    ' 1) Ask where to save
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=Range("Carrier_SiteNumber") & "_" & CleanFileName(Range("Carrier_SiteName")) & "_" & Range("Carrier") & "_" & Format(Date, "YYYY-MM-DD") & "_" & Format(Now, "hh-mm-ss") & "_" & "Loading Profile.tml", _
        FileFilter:="TML Files (*.tml), *.tml")
    If savePath = False Then Exit Sub

    ' 2) Build the XML declaration and root open
    xmlText = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
    xmlText = xmlText & _
       "<ccisitesRisatowerDataTransfer " & _
       "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
       "xsi:noNamespaceSchemaLocation=""CCISites_RISATower_Schema_v1.xsd"">" & vbCrLf

    ' 3) Insert your static header block
    xmlText = xmlText & BuildStaticHeader()

    ' 4) Append your data sections
    'xmlText = xmlText & BuildSectionXML("FeedlineData", "feedlineData", "feedline")
    If Sheets("DiscreteLoadData").Range("A2") <> 0 Then
        xmlText = xmlText & BuildSectionXML("DiscreteLoadData", "discreteLoadData", "discreteLoad")
    End If
    If Sheets("DishData").Range("A2") <> 0 Then
        xmlText = xmlText & BuildSectionXML("DishData", "dishData", "dish")
    End If

    ' 5) Close the root element
    xmlText = xmlText & "</ccisitesRisatowerDataTransfer>" & vbCrLf

    ' 6) Write it out as true UTF-8
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2                 ' adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText xmlText
        .SaveToFile savePath, 2    ' adSaveCreateOverWrite
        .Close
    End With

    MsgBox "tnxTower Loading File (Schema) exported to " & savePath, vbInformation         'MsgBox "TML exported in UTF-8 to " & savePath, vbInformation
End Sub

'------------------------------------------------------------------------------
' Returns your static header block (timestamp, unitSystem, etc.)
Function BuildStaticHeader() As String
    Dim out As String
    out = "    <timeStamp>2019-06-24T13:31:56.766-04:00</timeStamp>" & vbCrLf
    out = out & "    <unitSystem>" & vbCrLf
    out = out & "        <length>ft</length>" & vbCrLf
    out = out & "        <properties>in</properties>" & vbCrLf
    out = out & "        <speed>mph</speed>" & vbCrLf
    out = out & "        <angle>deg</angle>" & vbCrLf
    out = out & "        <weight>lbs</weight>" & vbCrLf
    out = out & "        <linearWeight>lbs/ft</linearWeight>" & vbCrLf
    out = out & "        <density>lbs/ft^3</density>" & vbCrLf
    out = out & "        <temperature>F</temperature>" & vbCrLf
    out = out & "        <strength>psi</strength>" & vbCrLf
    out = out & "    </unitSystem>" & vbCrLf
    out = out & "    <isLoadCase>false</isLoadCase>" & vbCrLf
    out = out & "    <towerMetadata>" & vbCrLf
    out = out & "        <cciInfo>" & vbCrLf
    out = out & "            <recipientName>NB+C ES</recipientName>" & vbCrLf
    out = out & "            <address>" & vbCrLf
    out = out & "                <streetAddress>8601 Six Forks Road&#xD;Suite 540</streetAddress>" & vbCrLf
    out = out & "                <city>Raleigh</city>" & vbCrLf
    out = out & "                <state>NC</state>" & vbCrLf
    out = out & "                <zip>27615</zip>" & vbCrLf
    out = out & "            </address>" & vbCrLf
    out = out & "            <cciBuNum>800000</cciBuNum>" & vbCrLf
    out = out & "            <cciSiteName>SITENAME</cciSiteName>" & vbCrLf
    out = out & "            <cciJdeJobNum>123456</cciJdeJobNum>" & vbCrLf
    out = out & "            <cciWorkOrderNum>654321</cciWorkOrderNum>" & vbCrLf
    out = out & "            <purchaseOrderNum>987654</purchaseOrderNum>" & vbCrLf
    out = out & "            <applicationNum>654987</applicationNum>" & vbCrLf
    out = out & "            <revisionNum>0</revisionNum>" & vbCrLf
    out = out & "        </cciInfo>" & vbCrLf
    out = out & "        <additionalProjectInfo>" & vbCrLf
    out = out & "            <projectNum>1759945</projectNum>" & vbCrLf
    out = out & "            <jobType>Analysis</jobType>" & vbCrLf
    out = out & "            <carrierName>T-Mobile</carrierName>" & vbCrLf
    out = out & "            <carrierSiteNum>9BH0198A</carrierSiteNum>" & vbCrLf
    out = out & "            <carrierSiteName>Mount Olive - US 2</carrierSiteName>" & vbCrLf
    out = out & "            <address>" & vbCrLf
    out = out & "                <streetAddress>12353 US Hwy 280</streetAddress>" & vbCrLf
    out = out & "                <city>Sylacauga</city>" & vbCrLf
    out = out & "                <county>Coosa County</county>" & vbCrLf
    out = out & "                <state>AL</state>" & vbCrLf
    out = out & "                <zip>35150</zip>" & vbCrLf
    out = out & "            </address>" & vbCrLf
    out = out & "            <latitude>33.073055555555555</latitude>" & vbCrLf
    out = out & "            <longtitude>-86.16833333333334</longtitude>" & vbCrLf
    out = out & "            <towerManufacturer>ROHN</towerManufacturer>" & vbCrLf
    out = out & "            <originalDesignSpeed>0.0</originalDesignSpeed>" & vbCrLf
    out = out & "            <groutInstalledIndicator>false</groutInstalledIndicator>" & vbCrLf
    out = out & "        </additionalProjectInfo>" & vbCrLf
    out = out & "        <documentData>" & vbCrLf
    out = out & "            <doc>" & vbCrLf
    out = out & "                <type>4-GEOTECHNICAL REPORTS</type>" & vbCrLf
    out = out & "                <referenceNum>10478</referenceNum>" & vbCrLf
    out = out & "                <source>CCISITES</source>" & vbCrLf
    out = out & "            </doc>" & vbCrLf
    out = out & "            <doc>" & vbCrLf
    out = out & "                <type>4-TOWER FOUNDATION DRAWINGS/DESIGN/SPECS</type>" & vbCrLf
    out = out & "                <referenceNum>599407</referenceNum>" & vbCrLf
    out = out & "                <source>CCISITES</source>" & vbCrLf
    out = out & "            </doc>" & vbCrLf
    out = out & "            <doc>" & vbCrLf
    out = out & "                <type>4-TOWER MANUFACTURER DRAWINGS</type>" & vbCrLf
    out = out & "                <referenceNum>366117</referenceNum>" & vbCrLf
    out = out & "                <source>CCISITES</source>" & vbCrLf
    out = out & "            </doc>" & vbCrLf
    out = out & "            <doc>" & vbCrLf
    out = out & "                <type>4-TOWER STRUCTURAL ANALYSIS REPORTS</type>" & vbCrLf
    out = out & "                <referenceNum>6322415</referenceNum>" & vbCrLf
    out = out & "                <source>CCISITES</source>" & vbCrLf
    out = out & "            </doc>" & vbCrLf
    out = out & "        </documentData>" & vbCrLf
    out = out & "    </towerMetadata>" & vbCrLf
    BuildStaticHeader = out
End Function


' Helper to recreate each section as XML text (same logic as WriteSection)
Function BuildSectionXML( _
    sheetName As String, wrapperTag As String, itemTag As String _
) As String
    Dim ws       As Worksheet
    Dim lastCol  As Long, lastRow As Long
    Dim r As Long, c As Long
    Dim hdr()    As String
    Dim out      As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ReDim hdr(1 To lastCol)
    For c = 1 To lastCol
        hdr(c) = ws.Cells(1, c).Value
    Next c

    out = "  <" & wrapperTag & ">" & vbCrLf
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For r = 2 To lastRow
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then GoTo SkipRow
        out = out & "    <" & itemTag & ">" & vbCrLf
        For c = 1 To lastCol
            out = out & "      <" & hdr(c) & ">" & _
                  XMLEscape(CStr(ws.Cells(r, c).Value)) & _
                  "</" & hdr(c) & ">" & vbCrLf
        Next c
        out = out & "    </" & itemTag & ">" & vbCrLf
SkipRow:
    Next r
    out = out & "  </" & wrapperTag & ">" & vbCrLf

    BuildSectionXML = out
End Function

Function XMLEscape(s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    XMLEscape = s
End Function








Sub Import_Loading_tnxTower()
    Dim fd As FileDialog
    Dim filePath As String
    Dim xmlDoc As Object, nodeList As Object, Node As Object
    Dim ws As Worksheet
    Dim dataType As Variant
    Dim headers As Object, h As Variant
    Dim row As Long, col As Long
    Dim r As Range
    
    Application.ScreenUpdating = False

    ' Create header dictionaries
    Set headers = CreateObject("Scripting.Dictionary")
    
    headers.Add "feedline", Array("database", "USName", "SIName", "type", "cciCode", "count", "endHeight", "selfWeight", "widthOrDiameter")
    headers.Add "discreteLoad", Array("database", "USName", "SIName", "type", "height", "depth", "width", "selfWeight", "cciCode", "count", "face", "startHeight", "offsetType", "horizontalOffset", "verticalOffset")
    headers.Add "dish", Array("database", "USName", "SIName", "cciCode", "count", "face", "selfWeight", "offsetType", "horizontalOffset", "verticalOffset", "heightAboveBase", "outsideDiameter")

    ' Prompt user to select a TML file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Select a .TML XML File"
        .Filters.Clear
        .Filters.Add "TML XML Files", "*.tml"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    ' Load XML
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load filePath

    If xmlDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "Error parsing XML: " & xmlDoc.ParseError.Reason
        Exit Sub
    End If

    ' Loop through each data type
    For Each dataType In headers.Keys
        Dim wsName As String
        wsName = "Extract_" & UCase(Left(dataType, 1)) & Mid(dataType, 2) & "Data"
        Set ws = ThisWorkbook.Sheets(wsName)
        ws.Range("A:O").ClearContents
        
        ' Write headers
        For col = 0 To UBound(headers(dataType))
            ws.Cells(1, col + 1).Value = headers(dataType)(col)
        Next col

        ' Get nodes and write data
        Set nodeList = xmlDoc.getElementsByTagName(dataType)
        row = 2
        For Each Node In nodeList
            For col = 0 To UBound(headers(dataType))
                h = headers(dataType)(col)
                On Error Resume Next
                ws.Cells(row, col + 1).Value = Node.SelectSingleNode(h).Text
                On Error GoTo 0
            Next col
            row = row + 1
        Next Node
    Next dataType
    
    Dim rng As Range
    Dim dataRange As Range
    
    Call Paste_Loading
    
    MsgBox "tnxTower loading successfully imported!", vbInformation
    
    Application.ScreenUpdating = False
    
End Sub


Sub Paste_Loading()
Dim rng As Range
Dim dataRange As Range

Set rng = Sheets("Extract_DiscreteLoadData").Range("Q2").CurrentRegion
Set dataRange = rng.offset(1, 0).Resize(rng.Rows.Count - 2, rng.Columns.Count)

    With Sheets("Discrete Loads")
        For r = 4 To 53
        If .Cells(r, 1) = "" Then
            dataRange.Copy
            .Cells(r, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Exit Sub
        End If
        Next r
    End With

End Sub





