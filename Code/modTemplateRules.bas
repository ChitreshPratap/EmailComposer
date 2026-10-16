Attribute VB_Name = "modTemplateRules"
Option Explicit
'==============================================================================
' modTemplateRules  -  THE ONLY MODULE YOU EDIT WHEN ADDING / CHANGING TEMPLATES
'
' 1. GetFields       : which fields the form shows for each template
' 2. IsFieldVisible  : show/hide a field based on other values
' 3. BlockRule       : keep/remove a {{#BLOCK}}...{{/BLOCK}} section
' 4. Shared rules    : one function used by BOTH 2 and 3, so form and mail agree
'
' Field array layout: 0 Key | 1 Label | 2 Type | 3 Required | 4 ListValues | 5 Format
' Types: Text, Number, Date, List, MultiLine
' Field Key must match the {{Key}} token in the .oft exactly (case-insensitive).
'==============================================================================

'================= 1. FIELD DEFINITIONS =================
Public Function GetFields(ByVal tid As String) As Collection
    Dim f As New Collection
    Select Case UCase(Trim(tid))

        Case "T05"      'Catch-up update
            AddF f, "CustomerName", "Customer Name", "Text", True
            AddF f, "ClaimNo", "Claim Number", "Text", True
            AddF f, "CatchUpDays", "Catch-up Days", "Number", True
            AddF f, "ApprovalRef", "Approval Ref", "Text", True              'only when > 30
            AddF f, "ApproverName", "Approver Name", "Text", True            'only when > 30
            AddF f, "TargetDate", "Target Date", "Date", True, , "dd-mmm-yyyy" 'only when <= 30
            AddF f, "EscalationRef", "Escalation Ref", "Text", False
            AddF f, "EscalationDate", "Escalation Date", "Date", False, , "dd-mmm-yyyy"

        Case "T06"      'Status update
            AddF f, "CustomerName", "Customer Name", "Text", True
            AddF f, "CustomerEmail", "Customer Email", "Text", True
            AddF f, "ClaimNo", "Claim Number", "Text", True
            AddF f, "Status", "Status", "List", True, "Open;Pending;Closed"
            AddF f, "Remarks", "Remarks", "MultiLine", False

        'Case "T07"    'copy a block above and adjust

        Case Else
            Err.Raise vbObjectError + 600, , "No field definition in modTemplateRules.GetFields for template " & tid
    End Select
    Set GetFields = f
End Function

'================= 2. FIELD VISIBILITY =================
'Default = visible. Hidden fields are never required and are sent as blank.
Public Function IsFieldVisible(ByVal tid As String, ByVal key As String, ByVal v As Object) As Boolean
    IsFieldVisible = True
    Select Case UCase(Trim(tid))
        Case "T05"
            Select Case UCase(key)
                Case "APPROVALREF", "APPROVERNAME": IsFieldVisible = IsLongCatchUp(v)
                Case "TARGETDATE":                  IsFieldVisible = IsShortCatchUp(v)
            End Select
    End Select
End Function

'================= 3. BLOCK RULES =================
'Return True (keep) / False (remove).
'Return Empty (Case Else) = AUTO: keep only if every {{token}} inside the block is filled.
Public Function BlockRule(ByVal tid As String, ByVal id As String, ByVal v As Object) As Variant
    Select Case UCase(Trim(tid)) & "|" & UCase(Trim(id))
        Case "T05|LONG_CATCHUP":  BlockRule = IsLongCatchUp(v)
        Case "T05|SHORT_CATCHUP": BlockRule = IsShortCatchUp(v)
        'Example of a richer rule:
        'Case "T07|OVERDUE_NOTE": BlockRule = (Txt(v, "Status") = "Open" And DaysSince(v, "ClaimDate") > 60)
        Case Else:                BlockRule = Empty
    End Select
End Function

'================= 4. SHARED BUSINESS RULES =================
Private Function IsLongCatchUp(ByVal v As Object) As Boolean
    If IsNum(v, "CatchUpDays") Then IsLongCatchUp = (Num(v, "CatchUpDays") > 30)
End Function

Private Function IsShortCatchUp(ByVal v As Object) As Boolean
    'Exactly 30 days falls here. Change to < 30 if 30 belongs nowhere / elsewhere.
    If IsNum(v, "CatchUpDays") Then IsShortCatchUp = (Num(v, "CatchUpDays") <= 30)
End Function

'================= HELPERS =================
Private Sub AddF(f As Collection, ByVal key As String, ByVal lbl As String, ByVal typ As String, _
                 ByVal req As Boolean, Optional ByVal listVals As String = "", Optional ByVal fmt As String = "")
    f.Add Array(key, lbl, typ, req, listVals, fmt), key
End Sub

Public Function Has(ByVal v As Object, ByVal k As String) As Boolean
    If v.Exists(k) Then Has = Len(Trim("" & v(k))) > 0
End Function

Public Function Txt(ByVal v As Object, ByVal k As String) As String
    If v.Exists(k) Then Txt = Trim("" & v(k))
End Function

Public Function IsNum(ByVal v As Object, ByVal k As String) As Boolean
    If Has(v, k) Then IsNum = IsNumeric(v(k))
End Function

Public Function Num(ByVal v As Object, ByVal k As String) As Double
    If IsNum(v, k) Then Num = CDbl(v(k))
End Function

'Days from a date field to today; returns -1 if the field is blank or not a date
Public Function DaysSince(ByVal v As Object, ByVal k As String) As Long
    DaysSince = -1
    If Has(v, k) Then
        If IsDate(v(k)) Then DaysSince = DateDiff("d", CDate(v(k)), Date)
    End If
End Function
