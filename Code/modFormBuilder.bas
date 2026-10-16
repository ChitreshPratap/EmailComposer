Attribute VB_Name = "modFormBuilder"
Option Explicit
'==============================================================================
' modFormBuilder - generates a STARTER designed form for a template.
'
'   BuildDesignForm "T05"    -> creates UserForm frmT05 with every fld_/lbl_
'                               control, Create/Cancel buttons and the code.
'                               Then open it in the VBA designer and arrange it.
'   BuildAllDesignForms      -> same for every active template without a form
'
' One-time requirement:
'   File > Options > Trust Center > Trust Center Settings > Macro Settings >
'   tick "Trust access to the VBA project object model".
' (Only needed to RUN the builder; users of the tool do not need it.)
'==============================================================================

Private Const COLS_GAP As Single = 6

Public Sub BuildDesignForm(Optional ByVal tid As String = "")
    If Len(tid) = 0 Then tid = Trim(InputBox("Template ID to build a design form for (e.g. T05):", "Build design form"))
    If Len(tid) = 0 Then Exit Sub
    If CreateForm(tid) Then
        MsgBox "frm" & tid & " created." & vbCrLf & vbCrLf & _
               "Open it in the VBA editor (Alt+F11), arrange the controls," & vbCrLf & _
               "rename captions, add frames/tabs. Keep the fld_ names unchanged.", vbInformation
    End If
End Sub

Public Sub BuildAllDesignForms()
    Dim r As ListRow, tid As String, made As String
    For Each r In ThisWorkbook.Worksheets(SHEET_TEMPLATES).ListObjects(TABLE_TEMPLATES).ListRows
        If UCase(Trim(CStr(r.Range(1, COL_ACTIVE).Value))) = "Y" Then
            tid = Trim(CStr(r.Range(1, COL_TID).Value))
            If Not ComponentExists("frm" & tid) Then
                If CreateForm(tid, True) Then made = made & " frm" & tid
            End If
        End If
    Next r
    MsgBox IIf(Len(made) = 0, "No new forms needed.", "Created:" & made), vbInformation
End Sub

'------------------------------------------------------------------------------
Private Function CreateForm(ByVal tid As String, Optional ByVal quiet As Boolean = False) As Boolean
    Dim proj As Object, comp As Object, des As Object, flds As Collection
    Dim fd As Variant, c As Object, topPos As Single, frmName As String
    Dim lblW As Single, ctlW As Single, formW As Single, h As Single, code As String

    frmName = "frm" & tid

    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    If proj Is Nothing Or Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Cannot access the VBA project." & vbCrLf & _
               "Tick 'Trust access to the VBA project object model' in Trust Center > Macro Settings.", vbCritical
        Exit Function
    End If
    Err.Clear
    Set flds = GetFields(tid)
    If Err.Number <> 0 Then
        Dim msg As String: msg = Err.Description
        On Error GoTo 0
        If Not quiet Then MsgBox msg, vbExclamation
        Exit Function
    End If
    On Error GoTo 0

    If ComponentExists(frmName) Then
        If Not quiet Then MsgBox frmName & " already exists. Delete it first if you want to regenerate it.", vbExclamation
        Exit Function
    End If

    lblW = 130: ctlW = 220: formW = lblW + ctlW + COLS_GAP * 5

    Set comp = proj.VBComponents.Add(3)             'vbext_ct_MSForm
    comp.Name = frmName
    Set des = comp.Designer

    topPos = COLS_GAP * 2
    For Each fd In flds
        Set c = des.Controls.Add("Forms.Label.1")
        c.Name = "lbl_" & fd(0)
        c.Caption = fd(1) & IIf(fd(3), " *", "")
        c.Left = COLS_GAP * 2: c.Top = topPos + 3: c.Width = lblW: c.Height = 15

        Select Case fd(2)
            Case "List"
                Set c = des.Controls.Add("Forms.ComboBox.1")
                c.Style = 2                         'fmStyleDropDownList
                h = 18
            Case "MultiLine"
                Set c = des.Controls.Add("Forms.TextBox.1")
                c.MultiLine = True: c.EnterKeyBehavior = True: c.WordWrap = True
                c.ScrollBars = 2                    'fmScrollBarsVertical
                h = 48
            Case Else
                Set c = des.Controls.Add("Forms.TextBox.1")
                h = 18
        End Select
        c.Name = "fld_" & fd(0)
        c.Left = lblW + COLS_GAP * 3: c.Top = topPos: c.Width = ctlW: c.Height = h
        topPos = topPos + h + COLS_GAP
    Next fd

    topPos = topPos + COLS_GAP
    Set c = des.Controls.Add("Forms.CommandButton.1")
    c.Name = "btnCreate": c.Caption = "Create Email"
    c.Width = 90: c.Height = 24: c.Top = topPos: c.Left = formW - COLS_GAP * 4 - 180 - COLS_GAP

    Set c = des.Controls.Add("Forms.CommandButton.1")
    c.Name = "btnCancel": c.Caption = "Cancel": c.Cancel = True
    c.Width = 90: c.Height = 24: c.Top = topPos: c.Left = formW - COLS_GAP * 4 - 90

    comp.Properties("Caption").Value = tid & " - " & TemplateName(tid)
    comp.Properties("Width").Value = formW
    comp.Properties("Height").Value = topPos + 24 + 40

    code = "Option Explicit" & vbCrLf & _
           "'Designed form for template " & tid & ". Layout is free; keep the fld_ names." & vbCrLf & _
           "Private eng As clsFormEngine" & vbCrLf & vbCrLf & _
           "Private Sub UserForm_Initialize()" & vbCrLf & _
           "    Set eng = New clsFormEngine" & vbCrLf & _
           "    eng.Attach Me, """ & tid & """" & vbCrLf & _
           "    'Optional pre-fill, e.g.: eng.SetValue ""ClaimNo"", Range(""A2"").Value" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub btnCreate_Click()" & vbCrLf & _
           "    If eng.Submit Then Unload Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub btnCancel_Click()" & vbCrLf & _
           "    Unload Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub UserForm_Terminate()" & vbCrLf & _
           "    If Not eng Is Nothing Then eng.Detach" & vbCrLf & _
           "End Sub"

    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString code
    End With
    CreateForm = True
End Function

Private Function ComponentExists(ByVal nm As String) As Boolean
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(nm)
    ComponentExists = Not comp Is Nothing
End Function

Private Function TemplateName(ByVal tid As String) As String
    Dim r As Range
    Set r = FindTemplateRow(tid)
    If Not r Is Nothing Then TemplateName = CStr(r.Cells(1, COL_NAME).Value)
End Function
