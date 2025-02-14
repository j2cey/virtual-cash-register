Option Explicit

Private m_lLeft As Long
Private m_lTop As Long


'*****************************************************************************************************************************
'   Getters & Setters
'*****************************************************************************************************************************

'* * * GETTERS
Public Property Get Self() As CPoint
    Set Self = Me
End Property

Public Property Get Left() As Long
    Left = m_lLeft
End Property

Public Property Get Top() As Long
    Top = m_lTop
End Property

'* * * SETTERS
Public Property Let Left(lLeft As Long)
    m_lLeft = lLeft
End Property

Public Property Let Top(lTop As Long)
    m_lTop = lTop
End Property


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************
Public Sub Init(lLeft As Long, lTop As Long)
    Left = lLeft
    Top = lTop
End Sub

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub