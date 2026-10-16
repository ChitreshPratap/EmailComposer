Attribute VB_Name = "modTemplateEditor"
Option Explicit
'==============================================================================
' modTemplateEditor - the simple way for users to create / change templates
'
'   1. EditTemplate   -> opens the template (or a blank mail for a new one) in Outlook
'   2. user edits the text normally in Outlook (bold, colours, tables... all fine)
'   3. SaveTemplate   -> backs up the old file, saves the .oft, repairs hidden
'                        formatting inside {{placeholders}}, and validates it
'
' Users never use Outlook's File > Save As, and never press Ctrl+S in the
' template window (Ctrl+S saves a copy into Drafts, not into the template).
'
'   RepairTemplate / RepairAllTemplates -> clean existing .oft files in place
' Backups go to <Templates folder>\_backup\ with a date-time stamp.
'==============================================================================

Private editItem As Object
Private editInsp As Object                       'the Outlook window showing the template
Private editTid As String
Private editPath As String

'------------------------------------------------------------------------------
Public Sub EditTemplate(Optional ByVal tid As String = "")
    Dim tplRow As Range, ol As Object

    If Len(tid) = 0 Then tid = Trim(InputBox("Template ID to edit (e.g. T05):", "Edit template"))
    If Len(tid) = 0 Then Exit Sub

    If Not editItem Is Nothing Then
        If IsStillOpen() Then
            If MsgBox("Template " & editTid & " is still open for editing." & vbCrLf & _
                      "Discard those changes and open " & tid & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            On Error Resume Next
            editItem.Close 1                        'olDiscard
            On Error GoTo 0
        End If
        Set editItem = Nothing
    End If

    Set tplRow = FindTemplateRow(tid)
    If tplRow Is Nothing Then MsgBox "Add " & tid & " to tblTemplates first.", vbExclamation: Exit Sub
    editPath = ResolvePath(CStr(tplRow.Cells(1, COL_PATH).Value))
    editTid = tid

    Set ol = GetOutlook()
    If Len(Dir(editPath)) > 0 Then
        Set editItem = ol.CreateItemFromTemplate(editPath)
    Else
        If MsgBox("The file for " & tid & " does not exist yet:" & vbCrLf & editPath & vbCrLf & vbCrLf & _
                  "Create a new template?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Set editItem = ol.CreateItem(0)             'olMailItem
        editItem.Subject = TemplateTitle(tid)
    End If
    editItem.Display
    Set editInsp = editItem.GetInspector

    MsgBox "Template " & tid & " is open in Outlook." & vbCrLf & vbCrLf & _
           "1. Change the text and formatting as you like." & vbCrLf & _
           "2. Type placeholders as {{FieldName}}." & vbCrLf & _
           "3. Come back to Excel and click 'Save Template'." & vbCrLf & vbCrLf & _
           "Do NOT press Ctrl+S or use Save As in Outlook.", vbInformation, "Edit template"
End Sub

'------------------------------------------------------------------------------
Public Sub SaveTemplate()
    Dim n As Long

    If editItem Is Nothing Then MsgBox "No template is open for editing. Click 'Edit Template' first.", vbExclamation: Exit Sub
    If Not IsStillOpen() Then
        Set editItem = Nothing
        MsgBox "The template window was closed before saving, so the changes were lost." & vbCrLf & _
               "Click 'Edit Template' and try again.", vbExclamation
        Exit Sub
    End If

    EnsureFolder editPath
    BackupFile editPath

    On Error GoTo Fail
    editItem.SaveAs editPath, 2                     'olTemplate - captures what is on screen
    editItem.Close 1                                'olDiscard - nothing lands in Drafts
    Set editItem = Nothing
    Set editInsp = Nothing
    On Error GoTo 0

    n = RepairFile(editPath)                        'clean hidden formatting in placeholders
    ValidateTemplate editTid
    If n > 0 Then Debug.Print "SaveTemplate " & editTid & ": " & n & " placeholder(s) repaired"
    Exit Sub

Fail:
    MsgBox "Could not save the template:" & vbCrLf & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
Public Sub RepairTemplate(Optional ByVal tid As String = "")
    Dim tplRow As Range, path As String, n As Long
    If Len(tid) = 0 Then tid = Trim(InputBox("Template ID to repair (e.g. T05):", "Repair template"))
    If Len(tid) = 0 Then Exit Sub
    Set tplRow = FindTemplateRow(tid)
    If tplRow Is Nothing Then MsgBox tid & " is not in tblTemplates.", vbExclamation: Exit Sub
    path = ResolvePath(CStr(tplRow.Cells(1, COL_PATH).Value))
    If Len(Dir(path)) = 0 Then MsgBox "File not found: " & path, vbExclamation: Exit Sub

    BackupFile path
    n = RepairFile(path)
    MsgBox tid & ": " & n & " placeholder(s) repaired.", vbInformation
End Sub

Public Sub RepairAllTemplates()
    Dim r As ListRow, path As String, n As Long, total As Long, files As Long
    For Each r In ThisWorkbook.Worksheets(SHEET_TEMPLATES).ListObjects(TABLE_TEMPLATES).ListRows
        path = ResolvePath(CStr(r.Range(1, COL_PATH).Value))
        If Len(Dir(path)) > 0 Then
            BackupFile path
            n = RepairFile(path)
            total = total + n
            If n > 0 Then files = files + 1
        End If
    Next r
    MsgBox total & " placeholder(s) repaired in " & files & " template file(s).", vbInformation
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
'Opens the .oft, repairs placeholders and saves it back only if something changed
Private Function RepairFile(ByVal path As String) As Long
    Dim m As Object, html As String, n As Long
    Set m = GetOutlook().CreateItemFromTemplate(path)
    html = NormalizeTokens(m.HTMLBody, n)
    If n > 0 Then
        m.HTMLBody = html
        m.SaveAs path, 2
    End If
    m.Close 1
    RepairFile = n
End Function

Private Sub BackupFile(ByVal path As String)
    Dim folder As String, nm As String, bak As String, p As Long
    If Len(Dir(path)) = 0 Then Exit Sub
    p = InStrRev(path, "\")
    folder = Left(path, p) & "_backup"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    nm = Mid(path, p + 1)
    bak = folder & "\" & Left(nm, InStrRev(nm, ".") - 1) & "_" & Format(Now, "yyyymmdd_hhnnss") & ".oft"
    FileCopy path, bak
End Sub

Private Sub EnsureFolder(ByVal path As String)
    Dim folder As String
    folder = Left(path, InStrRev(path, "\") - 1)
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
End Sub

'True while the template window is still open in Outlook.
'A closed window raises an error when touched.
Private Function IsStillOpen() As Boolean
    Dim s As String
    If editItem Is Nothing Or editInsp Is Nothing Then Exit Function
    On Error Resume Next
    s = editInsp.Caption
    IsStillOpen = (Err.Number = 0)
    On Error GoTo 0
End Function
