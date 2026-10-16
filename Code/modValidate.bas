Attribute VB_Name = "modValidate"
Option Explicit
'==============================================================================
' modValidate
' Run after ANY edit to an .oft file. Reports to the "Validation" sheet:
'  - placeholders split by Outlook formatting (look right, but won't be replaced)
'  - malformed placeholders  e.g. {{ ClaimNo}} or {{Claim No}}
'  - placeholders with no field in GetFields
'  - {{#BLOCK}} without matching {{/BLOCK}} (and the reverse)
'  - block markers not on their own line
'  - which blocks are rule-driven vs AUTO
'==============================================================================

Private Const SHEET_VALID As String = "Validation"

Public Sub ValidateTemplate(Optional ByVal tid As String = "")
    Dim rep As New Collection
    If Len(tid) = 0 Then tid = Trim(InputBox("Template ID to validate (e.g. T05):", "Validate template"))
    If Len(tid) = 0 Then Exit Sub
    CheckTemplate tid, rep
    WriteReport rep
End Sub

Public Sub ValidateAllTemplates()
    Dim rep As New Collection, r As ListRow
    For Each r In ThisWorkbook.Worksheets(SHEET_TEMPLATES).ListObjects(TABLE_TEMPLATES).ListRows
        If UCase(Trim(CStr(r.Range(1, COL_ACTIVE).Value))) = "Y" Then
            CheckTemplate CStr(r.Range(1, COL_TID).Value), rep
        End If
    Next r
    WriteReport rep
End Sub

Private Sub CheckTemplate(ByVal tid As String, rep As Collection)
    Dim tplRow As Range, path As String, m As Object, flds As Collection
    Dim html As String, subj As String, addr As String, plainText As String
    Dim rawTok As Object, txtTok As Object, fieldKeys As Object
    Dim re As Object, mt As Object, k As Variant, fd As Variant
    Dim nErr As Long, tok As String, sCnt As Long, eCnt As Long
    Dim sPos As Long, ePos As Long, rule As Variant
    Const PAD As String = "(?:\s|&nbsp;|<(?!/?p[\s>])[^>]*>)*"

    tid = Trim(tid)
    Set tplRow = FindTemplateRow(tid)
    If tplRow Is Nothing Then AddIssue rep, tid, "ERROR", "Not listed in tblTemplates": Exit Sub

    path = ResolvePath(CStr(tplRow.Cells(1, COL_PATH).Value))
    If Len(Dir(path)) = 0 Then AddIssue rep, tid, "ERROR", "Template file not found: " & path: Exit Sub

    'Field definitions
    Set fieldKeys = CreateObject("Scripting.Dictionary")
    fieldKeys.CompareMode = vbTextCompare
    On Error Resume Next
    Set flds = GetFields(tid)
    If Err.Number <> 0 Then
        AddIssue rep, tid, "ERROR", Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If Not flds Is Nothing Then
        For Each fd In flds
            fieldKeys(CStr(fd(0))) = 0
        Next fd
    End If

    'Open template
    Set m = GetOutlook().CreateItemFromTemplate(path)
    html = m.HTMLBody
    subj = m.Subject
    m.Close 1                               'olDiscard
    Set m = Nothing
    addr = CStr(tplRow.Cells(1, COL_TO).Value) & " " & CStr(tplRow.Cells(1, COL_CC).Value)

    plainText = HtmlToText(html) & " " & subj & " " & addr

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True

    'Proper tokens in raw HTML vs visible text
    Set rawTok = CollectTokens(html & " " & subj & " " & addr)
    Set txtTok = CollectTokens(plainText)

    For Each k In txtTok.Keys
        If Not rawTok.Exists(k) Then
            AddIssue rep, tid, "ERROR", k & " is split by formatting in the .oft - delete it and retype it in one go"
            nErr = nErr + 1
        End If
    Next k

    'Malformed tokens like {{ ClaimNo}} / {{Claim No}}
    re.Pattern = "\{\{[^{}]{0,60}\}\}"
    For Each mt In re.Execute(plainText)
        If Not IsProperToken(mt.Value) Then
            AddIssue rep, tid, "ERROR", "Malformed placeholder " & mt.Value & " - use {{Key}} with no spaces"
            nErr = nErr + 1
        End If
    Next mt

    'Token <-> field cross-check
    For Each k In txtTok.Keys
        tok = CStr(k)
        If Left(tok, 3) <> "{{#" And Left(tok, 3) <> "{{/" Then
            tok = Mid(tok, 3, Len(tok) - 4)
            If fieldKeys.Exists(tok) Then
                fieldKeys(tok) = 1
            Else
                AddIssue rep, tid, "ERROR", "{{" & tok & "}} has no field in GetFields - it will stay unfilled"
                nErr = nErr + 1
            End If
        End If
    Next k
    For Each k In fieldKeys.Keys
        If fieldKeys(k) = 0 Then AddIssue rep, tid, "INFO", "Field " & k & " is not used in the template (fine if it only drives a rule)"
    Next k

    'Blocks
    re.Pattern = "\{\{([#/])(\w+)\}\}"
    Dim blocks As Object: Set blocks = CreateObject("Scripting.Dictionary")
    blocks.CompareMode = vbTextCompare
    For Each mt In re.Execute(plainText)
        blocks(mt.SubMatches(1)) = 1
    Next mt

    For Each k In blocks.Keys
        sCnt = CountMatches(plainText, "\{\{#" & k & "\}\}")
        eCnt = CountMatches(plainText, "\{\{/" & k & "\}\}")
        If sCnt <> 1 Or eCnt <> 1 Then
            AddIssue rep, tid, "ERROR", "Block " & k & ": found " & sCnt & " start and " & eCnt & " end markers (need exactly 1 each)"
            nErr = nErr + 1
        Else
            sPos = InStr(1, plainText, "{{#" & k & "}}", vbTextCompare)
            ePos = InStr(1, plainText, "{{/" & k & "}}", vbTextCompare)
            If ePos < sPos Then
                AddIssue rep, tid, "ERROR", "Block " & k & ": end marker appears before start marker"
                nErr = nErr + 1
            End If
            If CountMatches(html, "<p[^>]*>" & PAD & "\{\{#" & k & "\}\}" & PAD & "</p>") = 0 _
               Or CountMatches(html, "<p[^>]*>" & PAD & "\{\{/" & k & "\}\}" & PAD & "</p>") = 0 Then
                AddIssue rep, tid, "WARN", "Block " & k & ": a marker is not on its own line - removal may leave text or empty lines"
            End If
            rule = BlockRule(tid, CStr(k), CreateObject("Scripting.Dictionary"))
            If IsEmpty(rule) Then
                AddIssue rep, tid, "INFO", "Block " & k & ": AUTO (kept only when all its fields are filled)"
            Else
                AddIssue rep, tid, "INFO", "Block " & k & ": controlled by BlockRule in code"
            End If
        End If
    Next k

    If nErr = 0 Then AddIssue rep, tid, "OK", "No errors found"
End Sub

'------------------------------------------------------------------------------
Private Function CollectTokens(ByVal s As String) As Object
    Dim re As Object, mt As Object, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\{\{[#/]?\w+\}\}"
    For Each mt In re.Execute(s)
        d(mt.Value) = 1
    Next mt
    Set CollectTokens = d
End Function

Private Function IsProperToken(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\{\{[#/]?\w+\}\}$"
    IsProperToken = re.Test(s)
End Function

Private Function CountMatches(ByVal s As String, ByVal pattern As String) As Long
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True
    re.Pattern = pattern
    CountMatches = re.Execute(s).Count
End Function

Private Function HtmlToText(ByVal html As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True
    re.Pattern = "<!--[\s\S]*?-->":            html = re.Replace(html, "")
    re.Pattern = "<style[\s\S]*?</style>":     html = re.Replace(html, "")
    re.Pattern = "<[^>]+>":                    html = re.Replace(html, "")
    html = Replace(html, "&nbsp;", " ")
    html = Replace(html, "&#123;", "{")
    html = Replace(html, "&#125;", "}")
    html = Replace(html, "&amp;", "&")
    HtmlToText = html
End Function

Private Sub AddIssue(rep As Collection, ByVal tid As String, ByVal sev As String, ByVal msg As String)
    rep.Add Array(tid, sev, msg)
End Sub

Private Sub WriteReport(rep As Collection)
    Dim ws As Worksheet, i As Long, nErr As Long, nWarn As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_VALID)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_VALID
    End If

    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Checked", "TemplateID", "Severity", "Issue")
    ws.Range("A1:D1").Font.Bold = True
    For i = 1 To rep.Count
        ws.Cells(i + 1, 1).Value = Now
        ws.Cells(i + 1, 2).Value = rep(i)(0)
        ws.Cells(i + 1, 3).Value = rep(i)(1)
        ws.Cells(i + 1, 4).Value = rep(i)(2)
        Select Case rep(i)(1)
            Case "ERROR": ws.Cells(i + 1, 3).Interior.Color = RGB(255, 199, 206): nErr = nErr + 1
            Case "WARN":  ws.Cells(i + 1, 3).Interior.Color = RGB(255, 235, 156): nWarn = nWarn + 1
            Case "OK":    ws.Cells(i + 1, 3).Interior.Color = RGB(198, 239, 206)
        End Select
    Next i
    ws.Columns("A:D").AutoFit
    ws.Activate
    MsgBox "Validation finished: " & nErr & " error(s), " & nWarn & " warning(s)." & vbCrLf & _
           "See the '" & SHEET_VALID & "' sheet.", IIf(nErr > 0, vbExclamation, vbInformation), "Validate templates"
End Sub
