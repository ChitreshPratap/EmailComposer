Attribute VB_Name = "modTheme"
Option Explicit
'==============================================================================
' modTheme - ONE place for the look of every form.
'
' Every form (designed frmT05.. and the automatic frmCompose) calls ApplyTheme
' when it opens. ApplyTheme:
'   - moves the designed content down and adds a HEADER
'       (application name + template ID/name + optional logo + accent strip)
'   - adds a FOOTER (footer text, version, user)
'   - applies colours and fonts to every control
'
' Designers draw ONLY the content area. Do not draw header/footer yourself and
' do not name controls starting with thm_ (reserved).
' A control whose Tag contains "nostyle" keeps its own designed look.
'==============================================================================

'================= SETTINGS - edit here =================
Public Const APP_NAME As String = "Claims Email Composer"
Public Const APP_VERSION As String = "v1.0"
Public Const FOOTER_TEXT As String = "Fields marked * are required. The mail opens for review before sending."
Public Const LOGO_FILE As String = "Templates\logo.jpg"     '.jpg/.bmp/.gif only (not .png); "" = no logo

Public Const FONT_NAME As String = "Segoe UI"
Public Const FONT_SIZE As Single = 9
Public Const HEADER_H As Single = 46
Public Const FOOTER_H As Single = 22

'Colours (edit the RGB values)
Public Function ClrPrimary() As Long   'header/footer, frame captions
    ClrPrimary = RGB(31, 58, 109)
End Function
Public Function ClrAccent() As Long   'accent strip, Create button
    ClrAccent = RGB(0, 120, 212)
End Function
Public Function ClrFormBack() As Long   'form background
    ClrFormBack = RGB(245, 247, 250)
End Function
Public Function ClrSection() As Long   'frames
    ClrSection = RGB(255, 255, 255)
End Function
Public Function ClrText() As Long
    ClrText = RGB(33, 37, 41)
End Function
Public Function ClrSubtitle() As Long   'template name in header
    ClrSubtitle = RGB(200, 215, 235)
End Function
Public Function ClrBorder() As Long
    ClrBorder = RGB(206, 212, 218)
End Function
Public Function ClrButton() As Long   'secondary buttons
    ClrButton = RGB(233, 236, 239)
End Function
Public Function ClrFooterBack() As Long
    ClrFooterBack = RGB(233, 236, 239)
End Function
Public Function ClrFooterText() As Long
    ClrFooterText = RGB(108, 117, 125)
End Function

'================= APPLY =================
'Look comes from frmStyleMaster samples when present (modStyles), else from the
'coded colours above. Header/footer heights follow sty_Header / sty_Footer heights.
Public Sub ApplyTheme(ByVal frm As Object, ByVal subtitle As String)
    Dim c As Object, w As Single, lbl As Object, img As Object, logoPath As String
    Dim hH As Single, fH As Single, src As Object

    If HasControl(frm, "thm_Header") Then SetThemeSubtitle frm, subtitle: Exit Sub   'already themed

    hH = HEADER_H: fH = FOOTER_H
    Set src = MasterControl("Header"): If Not src Is Nothing Then hH = src.Height
    Set src = MasterControl("Footer"): If Not src Is Nothing Then fH = src.Height

    frm.BackColor = ClrFormBack
    If Not StyleMaster() Is Nothing Then frm.BackColor = StyleMaster().BackColor
    frm.Caption = APP_NAME

    'Move designed content down (top-level controls only; frame children move with their frame)
    For Each c In frm.Controls
        If IsTopLevel(frm, c) Then c.Top = c.Top + hH
    Next c
    frm.Height = frm.Height + hH + fH
    w = frm.InsideWidth

    'Style the designed controls before adding the theme controls
    For Each c In frm.Controls
        StyleControl c
    Next c

    '--- Header
    Set lbl = AddThemeLabel(frm, "thm_Header", 0, 0, w, hH, "")
    lbl.BackStyle = 1: lbl.BackColor = ClrPrimary
    FromMaster "Header", lbl

    Set lbl = AddThemeLabel(frm, "thm_AppName", 12, 6, w - 24, 18, APP_NAME)
    lbl.Font.Size = 12: lbl.Font.Bold = True: lbl.ForeColor = vbWhite
    FromMaster "AppName", lbl, True

    Set lbl = AddThemeLabel(frm, "thm_Template", 12, 25, w - 24, 14, subtitle)
    lbl.ForeColor = ClrSubtitle
    FromMaster "TemplateName", lbl, True

    Set lbl = AddThemeLabel(frm, "thm_Accent", 0, hH - 3, w, 3, "")
    lbl.BackStyle = 1: lbl.BackColor = ClrAccent
    Set src = MasterControl("Accent")
    If Not src Is Nothing Then
        CopyStyle src, lbl
        lbl.Height = src.Height: lbl.Top = hH - src.Height
    End If

    logoPath = ResolvePath(LOGO_FILE)
    If Len(LOGO_FILE) > 0 And Len(Dir(logoPath)) > 0 Then
        On Error Resume Next
        Set img = frm.Controls.Add("Forms.Image.1", "thm_Logo")
        img.Picture = LoadPicture(logoPath)
        img.PictureSizeMode = 3                     'fmPictureSizeModeZoom
        img.BorderStyle = 0: img.BackStyle = 0
        img.Height = hH - 12: img.Width = 80
        img.Top = 4: img.Left = w - img.Width - 10
        On Error GoTo 0
    End If

    '--- Footer
    Set lbl = AddThemeLabel(frm, "thm_Footer", 0, frm.InsideHeight - fH, w, fH, "")
    lbl.BackStyle = 1: lbl.BackColor = ClrFooterBack
    FromMaster "Footer", lbl

    Set lbl = AddThemeLabel(frm, "thm_FooterText", 10, frm.InsideHeight - fH + 5, w - 150, 12, FOOTER_TEXT)
    lbl.Font.Size = 7.5: lbl.ForeColor = ClrFooterText
    FromMaster "FooterText", lbl, True

    Set lbl = AddThemeLabel(frm, "thm_FooterInfo", w - 140, frm.InsideHeight - fH + 5, 130, 12, _
                            APP_VERSION & "  |  " & Environ("USERNAME"))
    lbl.Font.Size = 7.5: lbl.ForeColor = ClrFooterText
    FromMaster "FooterText", lbl, True
    lbl.TextAlign = 3                               'fmTextAlignRight
End Sub

'Copy a header/footer sample onto a theme label; text labels stay transparent
Private Sub FromMaster(ByVal styleName As String, ByVal lbl As Object, Optional ByVal transparent As Boolean = False)
    Dim src As Object
    Set src = MasterControl(styleName)
    If src Is Nothing Then Exit Sub
    CopyStyle src, lbl
    If transparent Then lbl.BackStyle = 0
End Sub

Public Sub SetThemeSubtitle(ByVal frm As Object, ByVal subtitle As String)
    On Error Resume Next
    frm.Controls("thm_Template").Caption = subtitle
End Sub

'Style a single control - also call this for controls created at runtime
Public Sub StyleControl(ByVal c As Object)
    If LCase(Left(c.Name, 4)) = "thm_" Then Exit Sub
    If InStr(1, c.Tag, "nostyle", vbTextCompare) > 0 Then Exit Sub
    If ApplyMasterStyle(c) Then Exit Sub            'sample in frmStyleMaster wins (modStyles)

    On Error Resume Next                            'not every control has every property
    c.Font.Name = FONT_NAME
    c.Font.Size = FONT_SIZE
    Select Case TypeName(c)
        Case "Label"
            c.BackStyle = 0                         'transparent
            c.ForeColor = ClrText
        Case "TextBox", "ComboBox"
            c.SpecialEffect = 0                     'flat
            c.BorderStyle = 1                       'single
            c.BorderColor = ClrBorder
            c.BackColor = vbWhite
            c.ForeColor = ClrText
        Case "Frame"
            c.SpecialEffect = 0
            c.BorderStyle = 1
            c.BorderColor = ClrBorder
            c.BackColor = ClrSection
            c.ForeColor = ClrPrimary
            c.Font.Bold = True
        Case "CommandButton"
            If c.Name = "btnCreate" Then
                c.BackColor = ClrAccent: c.ForeColor = vbWhite: c.Font.Bold = True
            Else
                c.BackColor = ClrButton: c.ForeColor = ClrText
            End If
        Case "CheckBox", "OptionButton"
            c.BackStyle = 0
            c.ForeColor = ClrText
        Case "MultiPage"
            c.BackColor = ClrFormBack
    End Select
End Sub

Public Function TemplateTitle(ByVal tid As String) As String
    Dim r As Range
    Set r = FindTemplateRow(tid)
    If r Is Nothing Then
        TemplateTitle = tid
    Else
        TemplateTitle = tid & "  |  " & CStr(r.Cells(1, COL_NAME).Value)
    End If
End Function

'================= HELPERS =================
Private Function AddThemeLabel(ByVal frm As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
                               ByVal w As Single, ByVal h As Single, ByVal cap As String) As Object
    Dim lbl As Object
    Set lbl = frm.Controls.Add("Forms.Label.1", nm)
    lbl.Left = l: lbl.Top = t: lbl.Width = w: lbl.Height = h
    lbl.Caption = cap
    lbl.BackStyle = 0
    lbl.WordWrap = False
    lbl.Font.Name = FONT_NAME
    lbl.Font.Size = FONT_SIZE
    Set AddThemeLabel = lbl
End Function

Private Function HasControl(ByVal frm As Object, ByVal nm As String) As Boolean
    Dim c As Object
    On Error Resume Next
    Set c = frm.Controls(nm)
    HasControl = Not c Is Nothing
End Function

Private Function IsTopLevel(ByVal frm As Object, ByVal c As Object) As Boolean
    On Error Resume Next
    IsTopLevel = (c.Parent.Name = frm.Name)
End Function
