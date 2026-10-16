VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' clsFormEngine - shared logic for every DESIGNED template form (frmT05, frmT06...)
'
' Naming convention inside a designed form:
'   fld_<Key>   input control for field <Key>  (TextBox or ComboBox)  REQUIRED
'   lbl_<Key>   its caption (optional)  - hidden/shown together with the field
'   sec_<Any>   a Frame - hidden automatically when ALL fld_ controls inside it
'               are hidden. Two sec_ frames can sit on top of each other
'               (e.g. long vs short catch-up) so only one shows at a time.
'
' Layout, captions, fonts, colours, frames, MultiPage tabs: all free in the designer.
' Rules (required, visibility, blocks) still come from modTemplateRules.
'==============================================================================

Private frm As Object
Private tid As String
Private flds As Collection
Private watchers As Collection
Private building As Boolean
Private ctlIndex As Object                      'name -> control, built once per form

Public Property Get TemplateID() As String
    TemplateID = tid
End Property

'------------------------------------------------------------------------------
Public Sub Attach(ByVal f As Object, ByVal templateKey As String)
    Dim fd As Variant, c As Object, w As clsFieldWatch, missing As String, extra As String
    Dim ctl As Object, known As Object

    If f Is Nothing Then Err.Raise vbObjectError + 800, "clsFormEngine.Attach", "Attach was called without a form (use: eng.Attach Me, ""T05"")."
    Set frm = f
    tid = Trim(templateKey)
    BuildIndex
    Set flds = GetFields(tid)                       'raises if the template has no fields
    Set watchers = New Collection
    Set known = CreateObject("Scripting.Dictionary")
    known.CompareMode = vbTextCompare

    building = True
    For Each fd In flds
        known("fld_" & fd(0)) = 1
        Set c = FindCtl("fld_" & fd(0))
        If c Is Nothing Then
            missing = missing & vbCrLf & "  fld_" & fd(0)
        Else
            Select Case TypeName(c)
                Case "ComboBox"
                    If c.ListCount = 0 And Len(fd(4)) > 0 Then c.List = Split(fd(4), ";")
                    Set w = New clsFieldWatch: Set w.cb = c
                Case "TextBox"
                    Set w = New clsFieldWatch: Set w.tb = c
                Case Else
                    Set w = Nothing
            End Select
            If Not w Is Nothing Then
                Set w.Owner = Me
                watchers.Add w
            End If
            If fd(2) = "Date" And Len(c.ControlTipText) = 0 Then c.ControlTipText = "Date, e.g. " & Format(Date, "dd-mmm-yyyy")
        End If
    Next fd

    For Each ctl In frm.Controls
        If LCase(Left(ctl.Name, 4)) = "fld_" And Not known.Exists(ctl.Name) Then extra = extra & vbCrLf & "  " & ctl.Name
    Next ctl
    building = False

    ApplyTheme frm, TemplateTitle(tid)              'header, footer, colours (modTheme)

    If Len(missing) > 0 Or Len(extra) > 0 Then
        MsgBox "Design check for " & frm.Name & " (" & tid & "):" & _
               IIf(Len(missing) > 0, vbCrLf & vbCrLf & "Missing controls (field will be blank):" & missing, "") & _
               IIf(Len(extra) > 0, vbCrLf & vbCrLf & "Controls with no field in GetFields (ignored):" & extra, ""), _
               vbExclamation, "Email Composer - design check"
    End If

    RefreshVisibility
End Sub

Public Sub Detach()
    Set watchers = Nothing
    Set ctlIndex = Nothing
    Set frm = Nothing
End Sub

'------------------------------------------------------------------------------
' Show/hide - positions are NOT changed; the designer's layout is kept
'------------------------------------------------------------------------------
Public Sub RefreshVisibility()
    Dim fd As Variant, vis As Boolean, vals As Object, c As Object
    If building Or frm Is Nothing Then Exit Sub

    Set vals = ReadValues()
    For Each fd In flds
        vis = IsFieldVisible(tid, CStr(fd(0)), vals)
        Set c = FindCtl("fld_" & fd(0)): If Not c Is Nothing Then c.Visible = vis
        Set c = FindCtl("lbl_" & fd(0)): If Not c Is Nothing Then c.Visible = vis
    Next fd
    RefreshSections
End Sub

Private Sub RefreshSections()
    Dim sec As Object, inner As Object, hasFld As Boolean, anyVis As Boolean
    For Each sec In frm.Controls
        If LCase(Left(sec.Name, 4)) = "sec_" Then
            hasFld = False: anyVis = False
            On Error Resume Next
            For Each inner In sec.Controls
                If LCase(Left(inner.Name, 4)) = "fld_" Then
                    hasFld = True
                    If inner.Visible Then anyVis = True
                End If
            Next inner
            On Error GoTo 0
            If hasFld Then sec.Visible = anyVis
        End If
    Next sec
End Sub

'------------------------------------------------------------------------------
' Values
'------------------------------------------------------------------------------
Public Function ReadValues() As Object
    Dim d As Object, fd As Variant, c As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For Each fd In flds
        Set c = FindCtl("fld_" & fd(0))
        If c Is Nothing Then d(CStr(fd(0))) = "" Else d(CStr(fd(0))) = Trim("" & c.Value)
    Next fd
    Set ReadValues = d
End Function

'Pre-fill a field from code (e.g. in the form's Initialize)
Public Sub SetValue(ByVal key As String, ByVal v As Variant)
    Dim c As Object
    Set c = FindCtl("fld_" & key)
    If Not c Is Nothing Then c.Value = v
End Sub

'------------------------------------------------------------------------------
' Validate + compose. Returns True when the mail was created.
'------------------------------------------------------------------------------
Public Function Submit() As Boolean
    Dim vals As Object, fd As Variant, key As String, v As String, c As Object, fmt As String

    Set vals = ReadValues()
    For Each fd In flds
        key = fd(0)
        Set c = FindCtl("fld_" & key)
        If c Is Nothing Then
            vals(key) = ""
        ElseIf Not IsReallyVisible(c) Then
            vals(key) = ""                          'hidden field never leaks into the mail
        Else
            v = vals(key)
            If fd(3) And Len(v) = 0 Then
                MsgBox fd(1) & " is required.", vbExclamation: FocusOn c: Exit Function
            End If
            If Len(v) > 0 Then
                Select Case fd(2)
                    Case "Number"
                        If Not IsNumeric(v) Then MsgBox fd(1) & " must be a number.", vbExclamation: FocusOn c: Exit Function
                    Case "Date"
                        If Not IsDate(v) Then MsgBox fd(1) & " must be a valid date.", vbExclamation: FocusOn c: Exit Function
                        fmt = IIf(Len(fd(5)) = 0, "dd-mmm-yyyy", fd(5))
                        vals(key) = Format(CDate(v), fmt)
                End Select
            End If
        End If
    Next fd

    Submit = ComposeMail(tid, vals)
End Function

'------------------------------------------------------------------------------
'Control lookup through an index - never raises an error, returns Nothing if not found
Private Function FindCtl(ByVal nm As String) As Object
    If ctlIndex Is Nothing Then BuildIndex
    If ctlIndex.Exists(nm) Then Set FindCtl = ctlIndex(nm)
End Function

Private Sub BuildIndex()
    Dim c As Object
    Set ctlIndex = CreateObject("Scripting.Dictionary")
    ctlIndex.CompareMode = vbTextCompare
    If frm Is Nothing Then Exit Sub
    For Each c In frm.Controls                      'includes controls inside frames / pages
        If Not ctlIndex.Exists(c.Name) Then ctlIndex.Add c.Name, c
    Next c
End Sub

'Visible itself AND not inside a hidden sec_ frame
Private Function IsReallyVisible(ByVal c As Object) As Boolean
    Dim p As Object
    IsReallyVisible = c.Visible
    If Not IsReallyVisible Then Exit Function
    On Error Resume Next
    Set p = c.Parent
    Do While Not p Is Nothing
        If p Is frm Then Exit Do
        If TypeName(p) = "Frame" Then
            If Not p.Visible Then IsReallyVisible = False: Exit Function
        End If
        Set p = p.Parent
        If Err.Number <> 0 Then Exit Do
    Loop
End Function

Private Sub FocusOn(ByVal c As Object)
    On Error Resume Next                            'control may sit on another MultiPage tab
    c.SetFocus
End Sub
