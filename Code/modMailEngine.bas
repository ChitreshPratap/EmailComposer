Attribute VB_Name = "modMailEngine"
Option Explicit
'==============================================================================
' modMailEngine
' Opens the .oft template, resolves conditional blocks, replaces {{tokens}},
' fills Subject/To/CC and DISPLAYS the mail (never sends). Logs each compose.
' Late bound: no reference to Outlook or VBScript needed.
'==============================================================================

Public Const SHEET_TEMPLATES As String = "Templates"
Public Const TABLE_TEMPLATES As String = "tblTemplates"
Public Const SHEET_LOG As String = "SendLog"

'tblTemplates column positions
Public Const COL_TID As Long = 1
Public Const COL_NAME As Long = 2
Public Const COL_PATH As Long = 3
Public Const COL_TO As Long = 4
Public Const COL_CC As Long = 5
Public Const COL_ACTIVE As Long = 6

'------------------------------------------------------------------------------
' Entry point - assign this macro to a button
'------------------------------------------------------------------------------
Public Sub ShowComposer()
    frmCompose.Show
End Sub

'------------------------------------------------------------------------------
' Compose and display the mail
'------------------------------------------------------------------------------
Public Function ComposeMail(ByVal tid As String, ByVal vals As Object) As Boolean
    Dim tplRow As Range, olApp As Object, m As Object
    Dim html As String, subj As String, toAdr As String, ccAdr As String
    Dim k As Variant, path As String, leftover As String, tok As String

    On Error GoTo Fail

    Set tplRow = FindTemplateRow(tid)
    If tplRow Is Nothing Then Err.Raise vbObjectError + 700, , "Template not found in tblTemplates: " & tid

    path = ResolvePath(CStr(tplRow.Cells(1, COL_PATH).Value))
    If Len(Dir(path)) = 0 Then Err.Raise vbObjectError + 701, , "Template file not found:" & vbCrLf & path

    Set olApp = GetOutlook()
    Set m = olApp.CreateItemFromTemplate(path)

    'Blocks FIRST (they need to see which tokens are still empty), tokens second
    html = ProcessBlocks(m.HTMLBody, tid, vals)
    subj = m.Subject
    toAdr = CStr(tplRow.Cells(1, COL_TO).Value)
    ccAdr = CStr(tplRow.Cells(1, COL_CC).Value)

    For Each k In vals.Keys
        tok = "{{" & k & "}}"
        html = Replace(html, tok, HtmlEncode(CStr(vals(k))), , , vbTextCompare)
        subj = Replace(subj, tok, CStr(vals(k)), , , vbTextCompare)
        toAdr = Replace(toAdr, tok, CStr(vals(k)), , , vbTextCompare)
        ccAdr = Replace(ccAdr, tok, CStr(vals(k)), , , vbTextCompare)
    Next k

    leftover = FindLeftoverTokens(html & " " & subj & " " & toAdr & " " & ccAdr)

    m.HTMLBody = html
    m.Subject = subj
    If Len(Trim(toAdr)) > 0 Then m.To = toAdr
    If Len(Trim(ccAdr)) > 0 Then m.CC = ccAdr
    m.Display                   'review before sending - never .Send

    LogSend tid, vals

    If Len(leftover) > 0 Then
        MsgBox "Mail created, but these placeholders were not filled:" & vbCrLf & leftover & vbCrLf & vbCrLf & _
               "Run ValidateTemplate for " & tid & " to find the cause.", vbExclamation, "Check the mail"
    End If
    ComposeMail = True
    Exit Function

Fail:
    MsgBox "Could not compose the mail:" & vbCrLf & Err.Description, vbCritical, "Email Composer"
End Function

'------------------------------------------------------------------------------
' Conditional blocks:  {{#NAME}} ... {{/NAME}}
'------------------------------------------------------------------------------
Public Function ProcessBlocks(ByVal html As String, ByVal tid As String, ByVal vals As Object) As String
    Dim re As Object, mt As Object, ids As Object, k As Variant
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True
    Set ids = CreateObject("Scripting.Dictionary")
    ids.CompareMode = vbTextCompare

    re.Pattern = "\{\{#(\w+)\}\}"
    For Each mt In re.Execute(html)
        ids(mt.SubMatches(0)) = 1           'document order: outer blocks first
    Next mt

    For Each k In ids.Keys
        html = ApplyBlock(html, CStr(k), BlockIsOn(html, CStr(k), tid, vals))
    Next k
    ProcessBlocks = html
End Function

Private Function BlockIsOn(ByVal html As String, ByVal id As String, _
                           ByVal tid As String, ByVal vals As Object) As Boolean
    Dim rule As Variant, re As Object, mc As Object, mt As Object

    'Rule defined in modTemplateRules wins
    rule = BlockRule(tid, id, vals)
    If Not IsEmpty(rule) Then BlockIsOn = CBool(rule): Exit Function

    'AUTO: keep only if every token inside the block has a value
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "\{\{#" & id & "\}\}([\s\S]*?)\{\{/" & id & "\}\}"
    Set mc = re.Execute(html)
    If mc.Count = 0 Then Exit Function

    re.Global = True
    re.Pattern = "\{\{(\w+)\}\}"
    BlockIsOn = True
    For Each mt In re.Execute(mc(0).SubMatches(0))
        If Len(GetVal(vals, mt.SubMatches(0))) = 0 Then BlockIsOn = False: Exit Function
    Next mt
End Function

Private Function ApplyBlock(ByVal html As String, ByVal id As String, ByVal keep As Boolean) As String
    'PAD = whitespace / &nbsp; / inline tags (span, o:p, b...), never crosses a paragraph boundary
    Const PAD As String = "(?:\s|&nbsp;|<(?!/?p[\s>])[^>]*>)*"
    Dim re As Object, sPara As String, ePara As String, sRaw As String, eRaw As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True

    sRaw = "\{\{#" & id & "\}\}"
    eRaw = "\{\{/" & id & "\}\}"
    sPara = "<p[^>]*>" & PAD & sRaw & PAD & "</p>"
    ePara = "<p[^>]*>" & PAD & eRaw & PAD & "</p>"

    If keep Then                            'remove only the marker lines
        re.Pattern = sPara: html = re.Replace(html, "")
        re.Pattern = ePara: html = re.Replace(html, "")
        re.Pattern = sRaw:  html = re.Replace(html, "")
        re.Pattern = eRaw:  html = re.Replace(html, "")
    Else                                    'remove the whole section
        re.Pattern = sPara & "[\s\S]*?" & ePara
        If re.Test(html) Then
            html = re.Replace(html, "")
        Else
            re.Pattern = sRaw & "[\s\S]*?" & eRaw
            html = re.Replace(html, "")
        End If
    End If
    ApplyBlock = html
End Function

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Public Function GetVal(ByVal vals As Object, ByVal fld As String) As String
    fld = Trim(fld)
    If vals.Exists(fld) Then GetVal = Trim("" & vals(fld))
End Function

Private Function HtmlEncode(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, vbCrLf, "<br>")
    s = Replace(s, vbCr, "<br>")
    s = Replace(s, vbLf, "<br>")
    HtmlEncode = s
End Function

Private Function FindLeftoverTokens(ByVal s As String) As String
    Dim re As Object, mt As Object, seen As Object, out As String
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\{\{[#/]?\w+\}\}"
    Set seen = CreateObject("Scripting.Dictionary")
    For Each mt In re.Execute(s)
        If Not seen.Exists(mt.Value) Then
            seen(mt.Value) = 1
            out = out & vbCrLf & "  " & mt.Value
        End If
    Next mt
    FindLeftoverTokens = out
End Function

Public Function GetOutlook() As Object
    Dim o As Object
    On Error Resume Next
    Set o = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If o Is Nothing Then Set o = CreateObject("Outlook.Application")
    Set GetOutlook = o
End Function

Public Function FindTemplateRow(ByVal tid As String) As Range
    Dim lo As ListObject, r As ListRow
    Set lo = ThisWorkbook.Worksheets(SHEET_TEMPLATES).ListObjects(TABLE_TEMPLATES)
    For Each r In lo.ListRows
        If StrComp(Trim(CStr(r.Range(1, COL_TID).Value)), Trim(tid), vbTextCompare) = 0 Then
            Set FindTemplateRow = r.Range
            Exit Function
        End If
    Next r
End Function

'Relative paths (e.g. "Templates\T05.oft") resolve against the workbook folder
Public Function ResolvePath(ByVal p As String) As String
    p = Trim(p)
    If Mid(p, 2, 1) = ":" Or Left(p, 2) = "\\" Then
        ResolvePath = p
    Else
        ResolvePath = ThisWorkbook.Path & "\" & p
    End If
End Function

Private Sub LogSend(ByVal tid As String, ByVal vals As Object)
    Dim ws As Worksheet, r As Long, k As Variant, s As String
    On Error GoTo Done                      'logging must never block the mail
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    For Each k In vals.Keys
        If Len("" & vals(k)) > 0 Then s = s & k & "=" & vals(k) & "; "
    Next k
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = Environ("USERNAME")
    ws.Cells(r, 3).Value = tid
    ws.Cells(r, 4).Value = Left(s, 32000)
Done:
End Sub
