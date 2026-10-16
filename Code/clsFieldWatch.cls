VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFieldWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' clsFieldWatch - catches Change events of runtime-created form controls
' and asks the form to re-evaluate which fields are visible.
'==============================================================================

Public WithEvents tb As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1
Public WithEvents cb As MSForms.ComboBox
Attribute cb.VB_VarHelpID = -1
Public WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1
Public WithEvents lb As MSForms.ListBox
Attribute lb.VB_VarHelpID = -1
Public Owner As Object

Private Sub tb_Change()
    If Not Owner Is Nothing Then Owner.RefreshVisibility
End Sub

Private Sub cb_Change()
    If Not Owner Is Nothing Then Owner.RefreshVisibility
End Sub

Private Sub chk_Click()
    If Not Owner Is Nothing Then Owner.RefreshVisibility
End Sub

Private Sub lb_Change()
    If Not Owner Is Nothing Then Owner.RefreshVisibility
End Sub
