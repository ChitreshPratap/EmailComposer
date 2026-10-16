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
    Dim f As Object
    Set f = LoadedForm("frmCompose")
    If Not f Is Nothing Then
        If f.Visible Then Exit Sub                  'already on screen
    End If
    frmCompose.Show
End Sub

'Open one template directly (e.g. a button per template): ShowTemplate "T05"
'Uses the designed form frm<TID> if it exists, otherwise the automatic form.
Public Sub ShowTemplate(ByVal tid As String)
    Dim f As Object
    tid = Trim(tid)

    Set f = TryLoadDesignedForm(tid)
    If Not f Is Nothing Then
        SafeShow f
        Exit Sub
    End If

    'No designed form - use the automatic form with this template preselected
    Set f = LoadedForm("frmCompose")
    If Not f Is Nothing Then
        If f.Visible Then Exit Sub                  'picker already open
    End If
    frmCompose.PreselectTemplate tid
    If LoadedForm("frmCompose") Is Nothing Then Exit Sub    'it handed over and closed itself
    SafeShow frmCompose
End Sub

'Shows a form modally only if it is not already on screen (prevents run-time error 400)
Public Sub SafeShow(ByVal f As Object)
    If f Is Nothing Then Exit Sub
    If f.Visible Then Exit Sub
    f.Show
End Sub

'Returns the designed form frm<TID> (reusing an already loaded one), or Nothing.
'If the form EXISTS but fails while opening (error in its UserForm_Initialize),
'the real error is shown instead of being silently ignored.
Public Function TryLoadDesignedForm(ByVal tid As String) As Object
    Dim f As Object, nm As String, exists As Variant, errNo As Long, errMsg As String
    nm = "frm" & Trim(tid)

    Set f = LoadedForm(nm)
    If Not f Is Nothing Then Set TryLoadDesignedForm = f: Exit Function

    exists = DesignedFormExists(nm)                 'True / False / Empty (= cannot tell)
    If Not IsEmpty(exists) Then
        If exists = False Then Exit Function        'definitely no designed form
    End If

    On Error Resume Next
    Set f = VBA.UserForms.Add(nm)
    errNo = Err.Number: errMsg = Err.Description
    On Error GoTo 0

    If errNo <> 0 Then
        Set f = Nothing
        Debug.Print "TryLoadDesignedForm(" & nm & "): " & errNo & " " & errMsg
        If Not IsEmpty(exists) Then             'form exists -> its Initialize failed
            MsgBox nm & " exists but could not be opened." & vbCrLf & vbCrLf & _
                   "Error " & errNo & ": " & errMsg & vbCrLf & vbCrLf & _
                   "Run TestDesignedForm """ & Trim(tid) & """ to stop on the failing line.", _
                   vbExclamation, "Email Composer"
        End If
    End If
    Set TryLoadDesignedForm = f
End Function

'True/False when the VBA project can be inspected, Empty when it cannot
'(needs "Trust access to the VBA project object model").
Public Function DesignedFormExists(ByVal nm As String) As Variant
    Dim comp As Object, comps As Object
    On Error Resume Next
    Set comps = ThisWorkbook.VBProject.VBComponents
    If Err.Number <> 0 Or comps Is Nothing Then Exit Function      'Empty
    Set comp = comps(nm)
    DesignedFormExists = Not comp Is Nothing
    If DesignedFormExists Then DesignedFormExists = (comp.Type = 3)  'vbext_ct_MSForm
End Function

'Diagnostic: loads frm<TID> WITHOUT error trapping so the debugger stops on the failing line.
'Set VBA editor > Tools > Options > General > Error Trapping = "Break in Class Module" first.
Public Sub TestDesignedForm(Optional ByVal tid As String = "")
    Dim f As Object
    If Len(tid) = 0 Then tid = Trim(InputBox("Template ID (e.g. T05):", "Test designed form"))
    If Len(tid) = 0 Then Exit Sub
    Debug.Print "Exists in project: "; DesignedFormExists("frm" & tid)
    Set f = VBA.UserForms.Add("frm" & tid)
    MsgBox "frm" & tid & " loaded without errors.", vbInformation
    f.Show
End Sub

'Form already loaded in memory (shown or hidden)?
Public Function LoadedForm(ByVal nm As String) As Object
    Dim f As Object
    For Each f In VBA.UserForms
        If StrComp(f.Name, nm, vbTextCompare) = 0 Then Set LoadedForm = f: Exit Function
    Next f
End Function


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

    'Clean hidden formatting inside placeholders (spell-check marks etc.),
    'then blocks FIRST (they need to see which tokens are still empty), tokens second
    html = ProcessBlocks(NormalizeTokens(m.HTMLBody), tid, vals)
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
' NormalizeTokens - repairs placeholders that Outlook/Word split with hidden code.
'   {{<span class=SpellE>contributorName</span>}}   -> {{contributorName}}<span class=SpellE></span>
'   {<span>{</span>ClaimNo}}  /  {{ ClaimNo }}       -> {{ClaimNo}}...
' The removed tags are put back right AFTER the token, so the HTML stays balanced.
' repaired returns how many placeholders needed fixing.
'------------------------------------------------------------------------------
Public Function NormalizeTokens(ByVal html As String, Optional ByRef repaired As Long) As String
    Dim re As Object, reTag As Object, reKey As Object, mt As Object, tg As Object
    Dim out As String, pos As Long, key As String, tags As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\{(?:<[^>]*>)*\{((?:<[^>]*>|&nbsp;|[\s#/\w]){0,400}?)\}(?:<[^>]*>)*\}"

    Set reTag = CreateObject("VBScript.RegExp")
    reTag.Global = True
    reTag.Pattern = "<[^>]*>"

    Set reKey = CreateObject("VBScript.RegExp")
    reKey.Pattern = "^[#/]?\w+$"

    pos = 1
    For Each mt In re.Execute(html)
        out = out & Mid(html, pos, mt.FirstIndex + 1 - pos)

        key = reTag.Replace(mt.SubMatches(0), "")
        key = Replace(key, "&nbsp;", "")
        key = Replace(Replace(Replace(Replace(key, " ", ""), vbCr, ""), vbLf, ""), vbTab, "")

        If reKey.Test(key) Then
            tags = ""
            For Each tg In reTag.Execute(mt.Value)
                tags = tags & tg.Value
            Next tg
            If mt.Value <> "{{" & key & "}}" Then repaired = repaired + 1
            out = out & "{{" & key & "}}" & tags
        Else
            out = out & mt.Value                    'not a valid placeholder - leave untouched
        End If
        pos = mt.FirstIndex + mt.Length + 1
    Next mt

    NormalizeTokens = out & Mid(html, pos)
End Function

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
'Value of any input control as text. CheckBox: ticked = "Yes", not ticked = "".
Public Function ControlValue(ByVal c As Object) As String
    If c Is Nothing Then Exit Function
    Select Case TypeName(c)
        Case "CheckBox", "OptionButton"
            If c.Value = True Then ControlValue = "Yes"
        Case Else
            ControlValue = Trim("" & c.Value)
    End Select
End Function

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
    If o Is Nothing Then Err.Clear: Set o = CreateObject("Outlook.Application")
    On Error GoTo 0
    If o Is Nothing Then
        Err.Raise vbObjectError + 702, , _
            "Cannot connect to Outlook." & vbCrLf & vbCrLf & _
            "This tool needs CLASSIC Outlook. If you are using the 'New Outlook', " & _
            "switch the 'New Outlook' toggle off and try again." & vbCrLf & _
            "If classic Outlook is not installed, contact IT."
    End If
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
