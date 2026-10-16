Attribute VB_Name = "modTemplateEditor"
Option Explicit
'==============================================================================
' modTemplateEditor - how users create / change a template's wording
'   1. EditTemplate  -> opens the .oft in Outlook (a blank mail if it is new)
'   2. user edits the text normally; placeholders typed as {{Name}}
'   3. SaveTemplate  -> backs up the old file to Templates\_backup and saves the .oft
' Do NOT press Ctrl+S in Outlook (that saves to Drafts, not to the template).
'==============================================================================

Private editItem As Object
Private editInsp As Object
Private editPath As String

Public Sub EditTemplate(Optional ByVal tid As String = "")
    Dim tplRow As Range, ol As Object

    If Len(tid) = 0 Then tid = Trim(InputBox("Template ID to edit (e.g. T07):", "Edit template"))
    If Len(tid) = 0 Then Exit Sub

    If IsStillOpen() Then
        If MsgBox("Another template is still open for editing. Discard those changes?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        On Error Resume Next
        editItem.Close 1                            'olDiscard
        On Error GoTo 0
    End If

    Set tplRow = FindTemplateRow(tid)
    If tplRow Is Nothing Then MsgBox "Add " & tid & " to tblTemplates first.", vbExclamation: Exit Sub
    editPath = ResolvePath(CStr(tplRow.Cells(1, COL_PATH).Value))

    Set ol = GetOutlook()
    If Len(Dir(editPath)) > 0 Then
        Set editItem = ol.CreateItemFromTemplate(editPath)
    Else
        If MsgBox("No file yet for " & tid & ":" & vbCrLf & editPath & vbCrLf & vbCrLf & "Create a new template?", _
                  vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Set editItem = ol.CreateItem(0)             'olMailItem
    End If
    editItem.Display
    Set editInsp = editItem.GetInspector

    MsgBox "Template " & tid & " is open in Outlook." & vbCrLf & vbCrLf & _
           "Edit it, then come back and click 'Save Template'." & vbCrLf & _
           "Do NOT press Ctrl+S or use Save As in Outlook.", vbInformation, "Edit template"
End Sub

Public Sub SaveTemplate()
    If Not IsStillOpen() Then
        MsgBox "No template is open for editing (or its window was closed). Click 'Edit Template' first.", vbExclamation
        Exit Sub
    End If
    On Error GoTo Fail
    EnsureFolder editPath
    BackupFile editPath
    editItem.SaveAs editPath, 2                     'olTemplate
    editItem.Close 1                                'olDiscard - nothing lands in Drafts
    Set editItem = Nothing: Set editInsp = Nothing
    MsgBox "Template saved:" & vbCrLf & editPath, vbInformation
    Exit Sub
Fail:
    MsgBox "Could not save the template:" & vbCrLf & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
Private Function IsStillOpen() As Boolean
    Dim s As String
    If editItem Is Nothing Or editInsp Is Nothing Then Exit Function
    On Error Resume Next
    s = editInsp.Caption                            'a closed window raises an error
    IsStillOpen = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub BackupFile(ByVal path As String)
    Dim folder As String, nm As String
    If Len(Dir(path)) = 0 Then Exit Sub
    folder = Left(path, InStrRev(path, "\")) & "_backup"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    nm = Mid(path, InStrRev(path, "\") + 1)
    FileCopy path, folder & "\" & Left(nm, InStrRev(nm, ".") - 1) & "_" & Format(Now, "yyyymmdd_hhnnss") & ".oft"
End Sub

Private Sub EnsureFolder(ByVal path As String)
    Dim folder As String
    folder = Left(path, InStrRev(path, "\") - 1)
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
End Sub
