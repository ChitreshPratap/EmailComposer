Attribute VB_Name = "modSetup"
Option Explicit
'==============================================================================
' modSetup - run ONCE in a saved .xlsm
'   SetupWorkbook          : creates Templates / SendLog / Validation sheets,
'                            tblTemplates with 2 sample rows, and buttons
'   CreateSampleTemplates  : creates Templates\T05_CatchUpUpdate.oft and
'                            Templates\T06_StatusUpdate.oft next to the workbook
'==============================================================================

Public Sub SetupWorkbook()
    Dim ws As Worksheet, lo As ListObject

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save this workbook as .xlsm first, then run SetupWorkbook again.", vbExclamation
        Exit Sub
    End If

    '--- Templates sheet + table
    Set ws = EnsureSheet(SHEET_TEMPLATES)
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_TEMPLATES)
    On Error GoTo 0
    If lo Is Nothing Then
        ws.Range("A1:F1").Value = Array("TemplateID", "DisplayName", "FilePath", "To", "CC", "Active")
        ws.Range("A2:F2").Value = Array("T05", "Catch-up Update", "Templates\T05_CatchUpUpdate.oft", "", "", "Y")
        ws.Range("A3:F3").Value = Array("T06", "Status Update", "Templates\T06_StatusUpdate.oft", "{{CustomerEmail}}", "", "Y")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F3"), , xlYes)
        lo.Name = TABLE_TEMPLATES
        lo.TableStyle = "TableStyleMedium2"
        ws.Columns("A:F").AutoFit
        ws.Columns("C").ColumnWidth = 45
    End If

    AddButton ws, "btnCompose", "Compose Email", "ShowComposer", 1
    AddButton ws, "btnValidate", "Validate All Templates", "ValidateAllTemplates", 2
    AddButton ws, "btnSamples", "Create Sample Templates", "CreateSampleTemplates", 3
    AddButton ws, "btnEditTpl", "Edit Template", "EditTemplate", 4
    AddButton ws, "btnSaveTpl", "Save Template", "SaveTemplate", 5

    '--- SendLog sheet
    Set ws = EnsureSheet(SHEET_LOG)
    If Len(ws.Range("A1").Value) = 0 Then
        ws.Range("A1:D1").Value = Array("Timestamp", "User", "TemplateID", "Values")
        ws.Range("A1:D1").Font.Bold = True
        ws.Columns("A").ColumnWidth = 20
        ws.Columns("D").ColumnWidth = 100
    End If

    '--- Validation sheet
    EnsureSheet "Validation"

    ThisWorkbook.Worksheets(SHEET_TEMPLATES).Activate
    MsgBox "Setup complete." & vbCrLf & vbCrLf & _
           "Next: click 'Create Sample Templates', then 'Compose Email'.", vbInformation
End Sub

Public Sub CreateSampleTemplates()
    Dim folder As String, ol As Object

    If Len(ThisWorkbook.Path) = 0 Then MsgBox "Save the workbook first.", vbExclamation: Exit Sub
    folder = ThisWorkbook.Path & "\Templates"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder

    Set ol = GetOutlook()

    SaveOft ol, folder & "\T05_CatchUpUpdate.oft", _
        "Claim {{ClaimNo}} - catch-up update", _
        P("Dear {{CustomerName}},") & P("&nbsp;") & _
        P("We have reviewed claim <b>{{ClaimNo}}</b>.") & P("&nbsp;") & _
        P("{{#ESCALATION}}") & _
        P("This claim was escalated under reference {{EscalationRef}} on {{EscalationDate}}.") & P("&nbsp;") & _
        P("{{/ESCALATION}}") & _
        P("{{#LONG_CATCHUP}}") & _
        P("As the catch-up period of {{CatchUpDays}} days exceeds 30 days, please share approval " & _
          "<b>{{ApprovalRef}}</b> from {{ApproverName}}.") & P("&nbsp;") & _
        P("{{/LONG_CATCHUP}}") & _
        P("{{#SHORT_CATCHUP}}") & _
        P("The catch-up of {{CatchUpDays}} days will be processed by <b>{{TargetDate}}</b>.") & P("&nbsp;") & _
        P("{{/SHORT_CATCHUP}}") & _
        P("Regards,") & P("Claims Team")

    SaveOft ol, folder & "\T06_StatusUpdate.oft", _
        "Claim {{ClaimNo}} - status: {{Status}}", _
        P("Dear {{CustomerName}},") & P("&nbsp;") & _
        P("The current status of claim <b>{{ClaimNo}}</b> is <b>{{Status}}</b>.") & P("&nbsp;") & _
        P("{{#REMARKS}}") & _
        P("Remarks: {{Remarks}}") & P("&nbsp;") & _
        P("{{/REMARKS}}") & _
        P("Regards,") & P("Claims Team")

    MsgBox "Sample templates created in:" & vbCrLf & folder & vbCrLf & vbCrLf & _
           "Open one in Outlook (File > Open > Outlook Item / double-click) to see the markup.", vbInformation
End Sub

'------------------------------------------------------------------------------
Private Sub SaveOft(ol As Object, ByVal path As String, ByVal subj As String, ByVal body As String)
    Dim m As Object
    Set m = ol.CreateItem(0)                'olMailItem
    m.Subject = subj
    m.HTMLBody = "<html><body style=""font-family:Calibri,sans-serif;font-size:11pt"">" & body & "</body></html>"
    m.SaveAs path, 2                        'olTemplate
    m.Close 1                               'olDiscard
End Sub

Private Function P(ByVal s As String) As String
    P = "<p style=""margin:0"">" & s & "</p>"
End Function

Private Function EnsureSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = nm
    End If
End Function

Private Sub AddButton(ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal macro As String, ByVal idx As Long)
    Dim b As Object, leftPos As Double
    On Error Resume Next
    ws.Buttons(nm).Delete
    On Error GoTo 0
    leftPos = ws.Columns("H").Left
    Set b = ws.Buttons.Add(leftPos, 10 + (idx - 1) * 34, 170, 28)
    b.Name = nm
    b.Caption = cap
    b.OnAction = macro
End Sub
