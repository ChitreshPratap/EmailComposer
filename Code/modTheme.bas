Attribute VB_Name = "modTheme"
Option Explicit
'==============================================================================
' modTheme - common look for every form. Call in each form's Initialize:
'     ApplyTheme Me, TemplateTitle("T07")
' Adds a header (app name + template) and footer, moves your content below
' the header, and styles every control by type. Edit the settings below to
' change the look of all forms at once.
' Put "nostyle" in a control's Tag to keep its own designed look.
'==============================================================================

'================= SETTINGS =================
Public Const APP_NAME As String = "Email Composer"
Public Const FOOTER_TEXT As String = "The mail opens for review before sending."
Public Const FONT_NAME As String = "Segoe UI"
Public Const FONT_SIZE As Single = 9
Private Const HEADER_H As Single = 46
Private Const FOOTER_H As Single = 22

Private Function ClrPrimary() As Long
    ClrPrimary = RGB(31, 58, 109)
End Function
Private Function ClrAccent() As Long
    ClrAccent = RGB(0, 120, 212)
End Function
Private Function ClrFormBack() As Long
    ClrFormBack = RGB(245, 247, 250)
End Function
Private Function ClrText() As Long
    ClrText = RGB(33, 37, 41)
End Function
Private Function ClrBorder() As Long
    ClrBorder = RGB(206, 212, 218)
End Function
Private Function ClrButton() As Long
    ClrButton = RGB(233, 236, 239)
End Function
Private Function ClrMuted() As Long
    ClrMuted = RGB(108, 117, 125)
End Function

'================= APPLY =================
Public Sub ApplyTheme(ByVal frm As Object, ByVal subtitle As String)
    Dim c As Object, w As Single, lbl As Object

    frm.BackColor = ClrFormBack
    frm.Caption = APP_NAME

    For Each c In frm.Controls                      'move content below the header
        If IsTopLevel(frm, c) Then c.Top = c.Top + HEADER_H
    Next c
    frm.Height = frm.Height + HEADER_H + FOOTER_H
    w = frm.InsideWidth

    For Each c In frm.Controls
        StyleControl c
    Next c

    Set lbl = AddLabel(frm, "thm_Header", 0, 0, w, HEADER_H, "")
    lbl.BackStyle = 1: lbl.BackColor = ClrPrimary
    Set lbl = AddLabel(frm, "thm_AppName", 12, 6, w - 24, 18, APP_NAME)
    lbl.Font.Size = 12: lbl.Font.Bold = True: lbl.ForeColor = vbWhite
    Set lbl = AddLabel(frm, "thm_Template", 12, 25, w - 24, 14, subtitle)
    lbl.ForeColor = RGB(200, 215, 235)
    Set lbl = AddLabel(frm, "thm_Accent", 0, HEADER_H - 3, w, 3, "")
    lbl.BackStyle = 1: lbl.BackColor = ClrAccent

    Set lbl = AddLabel(frm, "thm_Footer", 0, frm.InsideHeight - FOOTER_H, w, FOOTER_H, "")
    lbl.BackStyle = 1: lbl.BackColor = ClrButton
    Set lbl = AddLabel(frm, "thm_FooterText", 10, frm.InsideHeight - FOOTER_H + 5, w - 20, 12, _
                       FOOTER_TEXT & "   |   " & Environ("USERNAME"))
    lbl.Font.Size = 7.5: lbl.ForeColor = ClrMuted
End Sub

Public Sub SetThemeSubtitle(ByVal frm As Object, ByVal subtitle As String)
    On Error Resume Next
    frm.Controls("thm_Template").Caption = subtitle
End Sub

Public Function TemplateTitle(ByVal tid As String) As String
    Dim r As Range
    Set r = FindTemplateRow(tid)
    If r Is Nothing Then TemplateTitle = tid Else TemplateTitle = tid & "  |  " & CStr(r.Cells(1, COL_NAME).Value)
End Function

'Style one control by its type - also usable for controls created at runtime
Public Sub StyleControl(ByVal c As Object)
    If LCase(Left(c.Name, 4)) = "thm_" Then Exit Sub
    If InStr(1, c.Tag, "nostyle", vbTextCompare) > 0 Then Exit Sub

    On Error Resume Next                            'not every control has every property
    c.Font.Name = FONT_NAME
    c.Font.Size = FONT_SIZE
    Select Case TypeName(c)
        Case "Label", "CheckBox", "OptionButton"
            c.BackStyle = 0
            c.ForeColor = ClrText
        Case "TextBox", "ComboBox", "ListBox"
            c.SpecialEffect = 0
            c.BorderStyle = 1
            c.BorderColor = ClrBorder
            c.BackColor = vbWhite
            c.ForeColor = ClrText
        Case "Frame"
            c.SpecialEffect = 0
            c.BorderStyle = 1
            c.BorderColor = ClrBorder
            c.BackColor = vbWhite
            c.ForeColor = ClrPrimary
            c.Font.Bold = True
        Case "CommandButton"
            If c.Name = "btnCreate" Then
                c.BackColor = ClrAccent: c.ForeColor = vbWhite: c.Font.Bold = True
            Else
                c.BackColor = ClrButton: c.ForeColor = ClrText
            End If
    End Select
End Sub

'================= HELPERS =================
Private Function AddLabel(ByVal frm As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single, ByVal cap As String) As Object
    Dim lbl As Object
    Set lbl = frm.Controls.Add("Forms.Label.1", nm)
    lbl.Left = l: lbl.Top = t: lbl.Width = w: lbl.Height = h
    lbl.Caption = cap
    lbl.BackStyle = 0
    lbl.WordWrap = False
    lbl.Font.Name = FONT_NAME
    lbl.Font.Size = FONT_SIZE
    Set AddLabel = lbl
End Function

Private Function IsTopLevel(ByVal frm As Object, ByVal c As Object) As Boolean
    On Error Resume Next
    IsTopLevel = (c.Parent.Name = frm.Name)
End Function
