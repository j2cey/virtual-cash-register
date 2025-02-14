Option Explicit

Private m_id As Long
Private m_payment_id As Long

Private m_status As CPaymentStatus
Private m_set_by As CActor


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CPaymentStatusHistory
    Set Self = Me
End Property

' Id
Public Property Let Id(lId As Long)
    m_id = lId
End Property

Public Property Get Id() As Long
    Id = m_id
End Property

' PaymentId
Public Property Let PaymentId(lPaymentId As Long)
    m_payment_id = lPaymentId
End Property

Public Property Get PaymentId() As Long
    PaymentId = m_payment_id
End Property

' Status
Public Property Set Status(oStatus As CPaymentStatus)
    Set m_status = oSetBy
End Property

Public Property Get Status() As CPaymentStatus
    Set Status = m_status
End Property

' SetBy
Public Property Set SetBy(oSetBy As CActor)
    Set m_set_by = oSetBy
End Property

Public Property Get SetBy() As CActor
    Set SetBy = m_set_by
End Property



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(lId As Long, lPaymentId As Long, lPaymentstatusId As Long, lUserId As Long, dStatusSetAt As Date)
    Me.Id = lId
    
    Me.PaymentId = lPaymentId
    Set Me.Status = NewPaymentStatus(lPaymentstatusId)
    Set Me.SetBy = NewActor(lUserId, dStatusSetAt)
    
End Sub


'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    m_id = 0
End Sub



