VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMailModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' IMailModel - the contract every template model class follows.
' Do not put code here. Each cls<TID> says "Implements IMailModel" and
' fills in these four members.
'==============================================================================

'Template ID as listed in tblTemplates, e.g. "T07"
Public Property Get TemplateID() As String
End Property

'Return "" when the data is complete, otherwise the message to show the user
Public Function Validate() As String
End Function

'Placeholder name -> value, e.g. t("ContributorName") = ContributorName
'Fills {{ContributorName}} in body, subject, To and CC
Public Function Tokens() As Object
End Function

'Section name -> True (keep) / False (remove), e.g. s("EDUCATION") = EducationDetails
'Controls {{#EDUCATION}} ... {{/EDUCATION}}
'Sections you leave out are AUTO: kept only when all their {{tokens}} are filled
Public Function Sections() As Object
End Function
