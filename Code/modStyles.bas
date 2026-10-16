Attribute VB_Name = "modStyles"
Option Explicit
'==============================================================================
' modStyles - "format painter" for all forms, driven by ONE sample form.
'
' frmStyleMaster holds one sample control per style, named sty_<Style>:
'   sty_Label  sty_RequiredLabel  sty_TextBox  sty_ComboBox  sty_ListBox
'   sty_Frame  sty_CommandButton  sty_PrimaryButton (btnCreate)
'   sty_CheckBox  sty_OptionButton  sty_MultiPage
'   sty_Header  sty_AppName  sty_TemplateName  sty_Accent  sty_Footer  sty_FooterText
'   + any custom style you add, e.g. sty_Heading
'
' Format a sample in the VBA designer (font, size, bold, colours, border...)
' and every matching control on every form copies it.
'
' Which sample a control uses:
'   1. Tag contains "nostyle"          -> untouched
'   2. Tag contains "style:<Name>"     -> sty_<Name>      (custom styles)
'   3. btnCreate                       -> sty_PrimaryButton
'   4. Label whose caption ends in "*" -> sty_RequiredLabel (if it exists)
'   5. otherwise                       -> sty_<ControlType>  e.g. sty_ComboBox
'
' Copied: font (name, size, bold, italic, underline), ForeColor, BackColor,
' BackStyle, BorderStyle, BorderColor, SpecialEffect, TextAlign.
' NOT copied: position, width, caption, list values, behaviour.
' Height is copied only when COPY_HEIGHT = True (single-line inputs/buttons).
'
' RUNTIME : automatic - every form styles itself when it opens.
' DESIGN  : PaintAllTemplateForms writes the styles into the forms' designs so
'           the VBA designer shows the same look (needs "Trust access to the
'           VBA project object model").
'==============================================================================

Public Const STYLE_FORM As String = "frmStyleMaster"
Public Const COPY_HEIGHT As Boolean = False

Private master As Object
Private masterTried As Boolean

'------------------------------------------------------------------------------
' Master access (loaded once, never shown)
'------------------------------------------------------------------------------
Public Function StyleMaster() As Object
    If Not masterTried Then
        masterTried = True
        On Error Resume Next
        Set master = VBA.UserForms.Add(STYLE_FORM)
        If Err.Number <> 0 Then Set master = Nothing
        On Error GoTo 0
    End If
    Set StyleMaster = master
End Function

'Call after editing frmStyleMaster in the same Excel session
Public Sub ResetStyleCache()
    On Error Resume Next
    If Not master Is Nothing Then Unload master
    Set master = Nothing
    masterTried = False
End Sub

Public Function MasterControl(ByVal styleName As String) As Object
    Dim m As Object
    Set m = StyleMaster()
    If m Is Nothing Then Exit Function
    On Error Resume Next
    Set MasterControl = m.Controls("sty_" & styleName)
End Function

Public Function MasterFor(ByVal c As Object) As Object
    Dim nm As String, src As Object
    nm = TagStyle(c.Tag)
    If Len(nm) > 0 Then Set MasterFor = MasterControl(nm): Exit Function

    If TypeName(c) = "CommandButton" Then
        If c.Name = "btnCreate" Then Set src = MasterControl("PrimaryButton")
    ElseIf TypeName(c) = "Label" Then
        If Right(Trim(c.Caption), 1) = "*" Then Set src = MasterControl("RequiredLabel")
    End If
    If src Is Nothing Then Set src = MasterControl(TypeName(c))
    Set MasterFor = src
End Function

'Style one control from the master. Returns False if no sample applies.
Public Function ApplyMasterStyle(ByVal c As Object) As Boolean
    Dim src As Object
    If IsExcluded(c) Then ApplyMasterStyle = True: Exit Function
    Set src = MasterFor(c)
    If src Is Nothing Then Exit Function
    CopyStyle src, c
    ApplyMasterStyle = True
End Function

'------------------------------------------------------------------------------
' The property copy
'------------------------------------------------------------------------------
Public Sub CopyStyle(ByVal src As Object, ByVal dst As Object)
    On Error Resume Next                    'not every control has every property
    dst.Font.Name = src.Font.Name
    dst.Font.Size = src.Font.Size
    dst.Font.Bold = src.Font.Bold
    dst.Font.Italic = src.Font.Italic
    dst.Font.Underline = src.Font.Underline
    dst.ForeColor = src.ForeColor
    dst.BackColor = src.BackColor
    dst.BackStyle = src.BackStyle
    dst.BorderStyle = src.BorderStyle
    dst.BorderColor = src.BorderColor
    dst.SpecialEffect = src.SpecialEffect
    dst.TextAlign = src.TextAlign
    If COPY_HEIGHT Then
        Select Case TypeName(dst)
            Case "ComboBox", "CommandButton": dst.Height = src.Height
            Case "TextBox": If Not dst.MultiLine Then dst.Height = src.Height
        End Select
    End If
End Sub

'------------------------------------------------------------------------------
' DESIGN-TIME painting (writes into the forms' designs)
'------------------------------------------------------------------------------
Public Sub PaintAllTemplateForms()
    Dim r As ListRow, nm As String, done As String, nCtl As Long, n As Long
    If Not CanAccessProject() Then Exit Sub
    ResetStyleCache
    If StyleMaster() Is Nothing Then MsgBox STYLE_FORM & " not found. Run BuildStyleMaster first.", vbExclamation: Exit Sub

    For Each r In ThisWorkbook.Worksheets(SHEET_TEMPLATES).ListObjects(TABLE_TEMPLATES).ListRows
        nm = "frm" & Trim(CStr(r.Range(1, COL_TID).Value))
        n = PaintDesign(nm)
        If n >= 0 Then done = done & vbCrLf & "  " & nm & " (" & n & " controls)": nCtl = nCtl + n
    Next r
    MsgBox IIf(Len(done) = 0, "No designed template forms found.", _
           "Styles written into:" & done & vbCrLf & vbCrLf & "Save the workbook to keep them."), vbInformation
End Sub

Public Sub PaintOneForm(Optional ByVal formName As String = "")
    Dim n As Long
    If Len(formName) = 0 Then formName = Trim(InputBox("Form to paint (e.g. frmT05):", "Paint form"))
    If Len(formName) = 0 Then Exit Sub
    If Not CanAccessProject() Then Exit Sub
    ResetStyleCache
    If StyleMaster() Is Nothing Then MsgBox STYLE_FORM & " not found. Run BuildStyleMaster first.", vbExclamation: Exit Sub
    n = PaintDesign(formName)
    If n < 0 Then
        MsgBox formName & " not found or its designer is not available.", vbExclamation
    Else
        MsgBox n & " controls styled in " & formName & ". Save the workbook to keep it.", vbInformation
    End If
End Sub

'Returns number of controls styled, or -1 if the form was not found
Private Function PaintDesign(ByVal formName As String) As Long
    Dim comp As Object, des As Object, c As Object, n As Long
    PaintDesign = -1
    If formName = STYLE_FORM Then Exit Function
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(formName)
    If comp Is Nothing Then Exit Function
    If comp.Type <> 3 Then Exit Function                   'vbext_ct_MSForm
    Set des = comp.Designer
    On Error GoTo 0
    If des Is Nothing Then Exit Function

    For Each c In des.Controls
        If Not IsExcluded(c) Then
            If ApplyMasterStyle(c) Then n = n + 1
        End If
    Next c
    On Error Resume Next
    comp.Properties("BackColor").Value = StyleMaster().BackColor
    On Error GoTo 0
    PaintDesign = n
End Function

'------------------------------------------------------------------------------
' Create frmStyleMaster with all samples, pre-set to the current theme colours
'------------------------------------------------------------------------------
Public Sub BuildStyleMaster()
    Dim comp As Object, des As Object, y As Single
    If Not CanAccessProject() Then Exit Sub
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(STYLE_FORM)
    On Error GoTo 0
    If Not comp Is Nothing Then MsgBox STYLE_FORM & " already exists. Edit it in the VBA designer.", vbInformation: Exit Sub

    Set comp = ThisWorkbook.VBProject.VBComponents.Add(3)
    comp.Name = STYLE_FORM
    Set des = comp.Designer
    comp.Properties("Caption").Value = "Style Master - format these samples; all forms copy them"
    comp.Properties("Width").Value = 400
    comp.Properties("Height").Value = 470
    comp.Properties("BackColor").Value = ClrFormBack

    '--- header samples
    With Smp(des, "Forms.Label.1", "sty_Header", "", 0, 0, 400, 46)
        .BackStyle = 1: .BackColor = ClrPrimary
    End With
    With Smp(des, "Forms.Label.1", "sty_AppName", "Application name", 12, 6, 250, 18)
        .BackStyle = 0: .ForeColor = vbWhite: .Font.Size = 12: .Font.Bold = True
    End With
    With Smp(des, "Forms.Label.1", "sty_TemplateName", "T05 | Template name", 12, 25, 250, 14)
        .BackStyle = 0: .ForeColor = ClrSubtitle
    End With
    With Smp(des, "Forms.Label.1", "sty_Accent", "", 0, 43, 400, 3)
        .BackStyle = 1: .BackColor = ClrAccent
    End With

    '--- control samples
    y = 60
    Info des, "Label:", y
    With Smp(des, "Forms.Label.1", "sty_Label", "Field label", 140, y, 200, 15)
        .BackStyle = 0
        .ForeColor = ClrText
    End With
    y = y + 24
    Info des, "Required label:", y
    With Smp(des, "Forms.Label.1", "sty_RequiredLabel", "Required field *", 140, y, 200, 15)
        .BackStyle = 0
        .ForeColor = ClrText
    End With
    y = y + 24
    Info des, "TextBox:", y
    With Smp(des, "Forms.TextBox.1", "sty_TextBox", "", 140, y, 200, 18)
        .SpecialEffect = 0
        .BorderStyle = 1
        .BorderColor = ClrBorder
        .ForeColor = ClrText
    End With
    y = y + 26
    Info des, "ComboBox:", y
    With Smp(des, "Forms.ComboBox.1", "sty_ComboBox", "", 140, y, 200, 18)
        .SpecialEffect = 0
        .BorderStyle = 1
        .BorderColor = ClrBorder
        .ForeColor = ClrText
    End With
    y = y + 26
    Info des, "ListBox:", y
    With Smp(des, "Forms.ListBox.1", "sty_ListBox", "", 140, y, 200, 30)
        .SpecialEffect = 0
        .BorderStyle = 1
        .BorderColor = ClrBorder
        .ForeColor = ClrText
    End With
    y = y + 38
    Info des, "Frame / section:", y
    With Smp(des, "Forms.Frame.1", "sty_Frame", "Section", 140, y, 200, 34)
        .SpecialEffect = 0
        .BorderStyle = 1
        .BorderColor = ClrBorder
        .BackColor = ClrSection
        .ForeColor = ClrPrimary
        .Font.Bold = True
    End With
    y = y + 42
    Info des, "CheckBox:", y
    With Smp(des, "Forms.CheckBox.1", "sty_CheckBox", "Check box", 140, y, 200, 18)
        .BackStyle = 0
        .ForeColor = ClrText
    End With
    y = y + 24
    Info des, "OptionButton:", y
    With Smp(des, "Forms.OptionButton.1", "sty_OptionButton", "Option", 140, y, 200, 18)
        .BackStyle = 0
        .ForeColor = ClrText
    End With
    y = y + 24
    Info des, "Heading (Tag style:Heading):", y
    With Smp(des, "Forms.Label.1", "sty_Heading", "Section heading", 140, y, 200, 16)
        .BackStyle = 0: .ForeColor = ClrPrimary: .Font.Size = 10: .Font.Bold = True
    End With
    y = y + 28
    Info des, "Buttons:", y + 4
    With Smp(des, "Forms.CommandButton.1", "sty_PrimaryButton", "Create Email", 140, y, 95, 24)
        .BackColor = ClrAccent: .ForeColor = vbWhite: .Font.Bold = True
    End With
    With Smp(des, "Forms.CommandButton.1", "sty_CommandButton", "Cancel", 245, y, 95, 24)
        .BackColor = ClrButton: .ForeColor = ClrText
    End With

    '--- footer samples
    With Smp(des, "Forms.Label.1", "sty_Footer", "", 0, 410, 400, 22)
        .BackStyle = 1: .BackColor = ClrFooterBack
    End With
    With Smp(des, "Forms.Label.1", "sty_FooterText", "Footer text", 10, 415, 300, 12)
        .BackStyle = 0: .ForeColor = ClrFooterText: .Font.Size = 7.5
    End With

    MsgBox STYLE_FORM & " created." & vbCrLf & vbCrLf & _
           "Open it in the VBA editor and format the samples. Forms pick up the changes the next time they open." & vbCrLf & _
           "Run PaintAllTemplateForms to also write them into the designs.", vbInformation
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Function Smp(ByVal des As Object, ByVal progId As String, ByVal nm As String, ByVal cap As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = des.Controls.Add(progId)
    c.Name = nm
    On Error Resume Next
    c.Caption = cap
    c.Font.Name = FONT_NAME
    c.Font.Size = FONT_SIZE
    On Error GoTo 0
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    Set Smp = c
End Function

Private Sub Info(ByVal des As Object, ByVal cap As String, ByVal t As Single)
    Dim c As Object
    Set c = des.Controls.Add("Forms.Label.1")
    c.Caption = cap
    c.Left = 10: c.Top = t + 2: c.Width = 125: c.Height = 15
    c.ForeColor = RGB(108, 117, 125)
    c.BackStyle = 0
End Sub

Private Function IsExcluded(ByVal c As Object) As Boolean
    If LCase(Left(c.Name, 4)) = "thm_" Then IsExcluded = True: Exit Function
    If LCase(Left(c.Name, 4)) = "sty_" Then IsExcluded = True: Exit Function
    If InStr(1, c.Tag, "nostyle", vbTextCompare) > 0 Then IsExcluded = True
End Function

Private Function TagStyle(ByVal tagText As String) As String
    Dim re As Object, mc As Object
    If InStr(1, tagText, "style:", vbTextCompare) = 0 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "style:(\w+)"
    Set mc = re.Execute(tagText)
    If mc.Count > 0 Then TagStyle = mc(0).SubMatches(0)
End Function

Private Function CanAccessProject() As Boolean
    Dim p As Object
    On Error Resume Next
    Set p = ThisWorkbook.VBProject.VBComponents
    CanAccessProject = (Err.Number = 0 And Not p Is Nothing)
    On Error GoTo 0
    If Not CanAccessProject Then
        MsgBox "Cannot access the VBA project." & vbCrLf & _
               "Tick 'Trust access to the VBA project object model' in Trust Center > Macro Settings.", vbCritical
    End If
End Function
