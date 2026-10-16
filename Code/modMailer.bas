Attribute VB_Name = "modMailer"
Option Explicit
'==============================================================================
' modMailer - the ONLY shared mail logic. You normally never edit this module.
'
'   ComposeFromModel model   -> validates the model, opens the template's .oft,
'                               keeps/removes {{#SECTIONS}}, fills {{tokens}},
'                               fills Subject/To/CC and DISPLAYS the mail.
'   ShowMainMenu             -> opens frmMain (assign to a sheet button)
'
' Each template = frm<TID> (screen) + cls<TID> (model, Implements IMailModel).
'==============================================================================

Public Const SHEET_TEMPLATES As String = "Templates"
Public Const TABLE_TEMPLATES As String = "tblTemplates"

'tblTemplates column positions
Public Const COL_TID As Long = 1
Public Const COL_NAME As Long = 2
Public Const COL_PATH As Long = 3
Public Const COL_TO As Long = 4
Public Const COL_CC As Long = 5
Public Const COL_ACTIVE As Long = 6

'------------------------------------------------------------------------------
' Entry points
'------------------------------------------------------------------------------
Public Sub ShowMainMenu()
    Dim f As Object
    Set f = LoadedForm("frmMain")
    If Not f Is Nothing Then
        If f.Visible Then Exit Sub
    End If
    frmMain.Show
End Sub

'Loads frm<TID> (not shown). Reports a clear message when it is missing or fails.
Public Function LoadTemplateForm(ByVal tid As String) As Object
    Dim f As Object, nm As String, errNo As Long, errMsg As String
    nm = "frm" & Trim(tid)

    Set f = LoadedForm(nm)
    If Not f Is Nothing Then Set LoadTemplateForm = f: Exit Function

    On Error Resume Next
    Set f = VBA.UserForms.Add(nm)
    errNo = Err.Number: errMsg = Err.Description
    On Error GoTo 0

    If errNo <> 0 Then
        MsgBox "Could not open the form " & nm & "." & vbCrLf & vbCrLf & _
               "Error " & errNo & ": " & errMsg & vbCrLf & vbCrLf & _
               "Create a UserForm named " & nm & " for this template.", _
               vbExclamation, "Email Composer"
        Exit Function
    End If
    Set LoadTemplateForm = f
End Function

'------------------------------------------------------------------------------
' Compose and display the mail from a model
'------------------------------------------------------------------------------
Public Function ComposeFromModel(ByVal model As IMailModel) As Boolean
    Dim tid As String, problem As String, tplRow As Range, path As String
    Dim m As Object, html As String, subj As String, toAdr As String, ccAdr As String
    Dim tokens As Object, sections As Object, k As Variant, tok As String, leftover As String

    tid = model.TemplateID

    problem = model.Validate()
    If Len(problem) > 0 Then MsgBox problem, vbExclamation, "Please check": Exit Function

    On Error GoTo Fail
    Set tplRow = FindTemplateRow(tid)
    If tplRow Is Nothing Then Err.Raise vbObjectError + 700, , "Template " & tid & " is not listed in tblTemplates."
    path = ResolvePath(CStr(tplRow.Cells(1, COL_PATH).Value))
    If Len(Dir(path)) = 0 Then Err.Raise vbObjectError + 701, , "Template file not found:" & vbCrLf & path

    Set tokens = model.Tokens()
    Set sections = model.Sections()
    If tokens Is Nothing Then Set tokens = NewMap()
    If sections Is Nothing Then Set sections = NewMap()

    Set m = GetOutlook().CreateItemFromTemplate(path)

    'Sections FIRST (AUTO sections need to see which tokens are empty), tokens second
    html = ApplySections(NormalizeTokens(m.HTMLBody), sections, tokens)
    subj = m.Subject
    toAdr = CStr(tplRow.Cells(1, COL_TO).Value)
    ccAdr = CStr(tplRow.Cells(1, COL_CC).Value)

    For Each k In tokens.Keys
        tok = "{{" & k & "}}"
        html = Replace(html, tok, HtmlEncode(CStr(tokens(k))), , , vbTextCompare)
        subj = Replace(subj, tok, CStr(tokens(k)), , , vbTextCompare)
        toAdr = Replace(toAdr, tok, CStr(tokens(k)), , , vbTextCompare)
        ccAdr = Replace(ccAdr, tok, CStr(tokens(k)), , , vbTextCompare)
    Next k

    leftover = FindLeftoverTokens(html & " " & subj & " " & toAdr & " " & ccAdr)

    m.HTMLBody = html
    m.Subject = subj
    If Len(Trim(toAdr)) > 0 Then m.To = toAdr
    If Len(Trim(ccAdr)) > 0 Then m.CC = ccAdr
    m.Display                                       'review before sending - never .Send

    If Len(leftover) > 0 Then
        MsgBox "Mail created, but these placeholders were not filled:" & leftover & vbCrLf & vbCrLf & _
               "Add them to Tokens() in cls" & tid & ", or check their spelling in the template.", _
               vbExclamation, "Check the mail"
    End If
    ComposeFromModel = True
    Exit Function

Fail:
    MsgBox "Could not compose the mail:" & vbCrLf & Err.Description, vbCritical, "Email Composer"
End Function

'------------------------------------------------------------------------------
' Sections  {{#NAME}} ... {{/NAME}}
'   NAME listed in the model's Sections()  -> kept when True, removed when False
'   NAME not listed (AUTO)                 -> kept only when every {{token}} inside is filled
'------------------------------------------------------------------------------
Public Function ApplySections(ByVal html As String, ByVal sections As Object, ByVal tokens As Object) As String
    Dim re As Object, mt As Object, ids As Object, k As Variant, keep As Boolean
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True
    Set ids = NewMap()

    re.Pattern = "\{\{#(\w+)\}\}"
    For Each mt In re.Execute(html)
        ids(mt.SubMatches(0)) = 1                   'document order: outer sections first
    Next mt

    For Each k In ids.Keys
        If MapHas(sections, CStr(k)) Then
            keep = CBool(sections(CStr(k)))
        Else
            keep = AllTokensFilled(html, CStr(k), tokens)
        End If
        html = ApplyBlock(html, CStr(k), keep)
    Next k
    ApplySections = html
End Function

Private Function AllTokensFilled(ByVal html As String, ByVal id As String, ByVal tokens As Object) As Boolean
    Dim re As Object, mc As Object, mt As Object, nm As String
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "\{\{#" & id & "\}\}([\s\S]*?)\{\{/" & id & "\}\}"
    Set mc = re.Execute(html)
    If mc.Count = 0 Then Exit Function
    re.Global = True
    re.Pattern = "\{\{(\w+)\}\}"
    AllTokensFilled = True
    For Each mt In re.Execute(mc(0).SubMatches(0))
        nm = mt.SubMatches(0)
        If Not MapHas(tokens, nm) Then AllTokensFilled = False: Exit Function
        If Len(Trim("" & tokens(nm))) = 0 Then AllTokensFilled = False: Exit Function
    Next mt
End Function

'------------------------------------------------------------------------------
' Small helpers for model classes
'------------------------------------------------------------------------------
'Case-insensitive Dictionary - use it in Tokens() and Sections()
Public Function NewMap() As Object
    Set NewMap = CreateObject("Scripting.Dictionary")
    NewMap.CompareMode = vbTextCompare
End Function

Public Function MapHas(ByVal d As Object, ByVal k As String) As Boolean
    If d Is Nothing Then Exit Function
    MapHas = d.Exists(k)
End Function

'True when value equals one of the ;-separated options (case-insensitive)
Public Function IsOneOf(ByVal value As String, ByVal options As String) As Boolean
    Dim o As Variant
    If Len(Trim(value)) = 0 Then Exit Function
    For Each o In Split(options, ";")
        If StrComp(Trim(o), Trim(value), vbTextCompare) = 0 Then IsOneOf = True: Exit Function
    Next o
End Function

'------------------------------------------------------------------------------
' Template engine internals
'------------------------------------------------------------------------------
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

'Form already loaded in memory (shown or hidden)?
Public Function LoadedForm(ByVal nm As String) As Object
    Dim f As Object
    For Each f In VBA.UserForms
        If StrComp(f.Name, nm, vbTextCompare) = 0 Then Set LoadedForm = f: Exit Function
    Next f
End Function
