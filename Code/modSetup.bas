Attribute VB_Name = "modSetup"
Option Explicit
'==============================================================================
' modSetup - run ONCE in a saved .xlsm
'   SetupWorkbook         : Templates sheet with tblTemplates and buttons
'   CreateSampleTemplate  : Templates\T07_EducationGuardian.oft (sample for clsT07/frmT07)
'==============================================================================

Public Sub SetupWorkbook()
    Dim ws As Worksheet, lo As ListObject

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save this workbook as .xlsm first, then run SetupWorkbook again.", vbExclamation
        Exit Sub
    End If

    Set ws = EnsureSheet(SHEET_TEMPLATES)
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_TEMPLATES)
    On Error GoTo 0
    If lo Is Nothing Then
        ws.Range("A1:F1").Value = Array("TemplateID", "DisplayName", "FilePath", "To", "CC", "Active")
        ws.Range("A2:F2").Value = Array("T07", "Education & Guardian", "Templates\T07_EducationGuardian.oft", "", "", "Y")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F2"), , xlYes)
        lo.Name = TABLE_TEMPLATES
        lo.TableStyle = "TableStyleMedium2"
        ws.Columns("A:F").AutoFit
        ws.Columns("C").ColumnWidth = 45
    End If

    AddButton ws, "btnMain", "Main Menu", "ShowMainMenu", 1
    AddButton ws, "btnEditTpl", "Edit Template", "EditTemplate", 2
    AddButton ws, "btnSaveTpl", "Save Template", "SaveTemplate", 3

    ws.Activate
    MsgBox "Setup complete. Next: run CreateSampleTemplate.", vbInformation
End Sub

Public Sub CreateSampleTemplate()
    Dim folder As String, m As Object
    If Len(ThisWorkbook.Path) = 0 Then MsgBox "Save the workbook first.", vbExclamation: Exit Sub
    folder = ThisWorkbook.Path & "\Templates"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder

    Set m = GetOutlook().CreateItem(0)
    m.Subject = "Details for {{ContributorName}}"
    m.HTMLBody = "<html><body style=""font-family:Calibri,sans-serif;font-size:11pt"">" & _
        P("Dear {{ContributorName}},") & P("&nbsp;") & _
        P("Thank you for submitting your details.") & P("&nbsp;") & _
        P("{{#EDUCATION}}") & _
        P("<b>Education Details</b>") & _
        P("12th Marks: {{Marks12th}}%") & _
        P("{{#BOARD}}") & P("Board: {{Board12th}}") & P("{{/BOARD}}") & _
        P("{{#YEAR}}") & P("Year of Passing: {{PassingYear12th}}") & P("{{/YEAR}}") & _
        P("&nbsp;") & _
        P("{{/EDUCATION}}") & _
        P("{{#GUARDIAN_VALID}}") & P("Your guardian is validated.") & P("&nbsp;") & P("{{/GUARDIAN_VALID}}") & _
        P("{{#GUARDIAN_INVALID}}") & P("You are not a valid guardian.") & P("&nbsp;") & P("{{/GUARDIAN_INVALID}}") & _
        P("Regards,") & P("Admissions Team") & "</body></html>"
    m.SaveAs folder & "\T07_EducationGuardian.oft", 2      'olTemplate
    m.Close 1                                               'olDiscard

    MsgBox "Created " & folder & "\T07_EducationGuardian.oft", vbInformation
End Sub

'------------------------------------------------------------------------------
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
    Dim b As Object
    On Error Resume Next
    ws.Buttons(nm).Delete
    On Error GoTo 0
    Set b = ws.Buttons.Add(ws.Columns("H").Left, 10 + (idx - 1) * 34, 170, 28)
    b.Name = nm
    b.Caption = cap
    b.OnAction = macro
End Sub
