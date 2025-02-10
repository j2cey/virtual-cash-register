Option Explicit

Private m_user As CUser
Private m_action_date As Date


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CActor
    Set Self = Me
End Property

' User
Public Property Set User(oUser As CUser)
    Set m_user = oUser
End Property

Public Property Get User() As CUser
    Set User = m_user
End Property

' ActionDate
Public Property Let ActionDate(dActionDate As Date)
    m_action_date = dActionDate
End Property

Public Property Get ActionDate() As Date
    ActionDate = m_action_date
End Property


'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(lUserId As Long, dActionDate As Date)
    Set Me.User = NewUser(lUserId)
    Me.ActionDate = dActionDate
End Sub


'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    
End Sub


