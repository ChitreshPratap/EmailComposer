VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IMailModel
'==============================================================================
' clsT07 - model for template T07 "Education & Guardian"
' frmT07 fills the properties; this class decides what goes into the mail.
'==============================================================================

'---------------- data (filled by frmT07) ----------------
Public ContributorName As String
Public EducationDetails As Boolean
Public Marks12th As String
Public Board12th As String
Public PassingYear12th As String
Public Guardian As String

'---------------- business rules ----------------
Public Function GuardianIsValid() As Boolean
    GuardianIsValid = IsOneOf(Guardian, "Mother;Father;Friend")
End Function

'---------------- IMailModel ----------------
Private Property Get IMailModel_TemplateID() As String
    IMailModel_TemplateID = "T07"
End Property

Private Function IMailModel_Validate() As String
    If Len(ContributorName) = 0 Then IMailModel_Validate = "Contributor Name is required.": Exit Function
    If EducationDetails Then
        If Len(Marks12th) = 0 Then IMailModel_Validate = "12th Marks are required.": Exit Function
        If Not IsNumeric(Marks12th) Then IMailModel_Validate = "12th Marks must be a number.": Exit Function
    End If
    If Len(Guardian) = 0 Then IMailModel_Validate = "Select a guardian.": Exit Function
End Function

'{{placeholders}} in the .oft
Private Function IMailModel_Tokens() As Object
    Dim t As Object: Set t = NewMap()
    t("ContributorName") = ContributorName
    t("Marks12th") = Marks12th
    t("Board12th") = Board12th
    t("PassingYear12th") = PassingYear12th
    t("Guardian") = Guardian
    Set IMailModel_Tokens = t
End Function

'{{#SECTIONS}} in the .oft  (BOARD and YEAR are left out on purpose = AUTO)
Private Function IMailModel_Sections() As Object
    Dim s As Object: Set s = NewMap()
    s("EDUCATION") = EducationDetails
    s("GUARDIAN_VALID") = GuardianIsValid()
    s("GUARDIAN_INVALID") = (Len(Guardian) > 0 And Not GuardianIsValid())
    Set IMailModel_Sections = s
End Function
