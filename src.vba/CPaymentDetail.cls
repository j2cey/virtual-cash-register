Option Explicit

Private m_id As Long
Private m_payment_id As Long

Private m_employee_matricule As String
Private m_employee_lastname As String
Private m_employee_firstname As String
Private m_receiver_identitfier_type As String
Private m_receiver_identifier As String
Private m_validation_kyc As String
Private m_validation_kyc_value As String
Private m_amount As Long
Private m_comment As String

Private m_created_by As CActor
Private m_updated_by As CActor


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CPaymentDetail
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

' EmployeeMatricule
Public Property Let EmployeeMatricule(sEmployeeMatricule As String)
    m_employee_matricule = sEmployeeMatricule
End Property

Public Property Get EmployeeMatricule() As String
    EmployeeMatricule = m_employee_matricule
End Property

' EmployeeLastName
Public Property Let EmployeeLastName(sEmployeeLastName As String)
    m_employee_lastname = sEmployeeLastName
End Property

Public Property Get EmployeeLastName() As String
    EmployeeLastName = m_employee_lastname
End Property

' EmployeeFirstName
Public Property Let EmployeeFirstName(sEmployeeFirstName As String)
    m_employee_firstname = sEmployeeFirstName
End Property

Public Property Get EmployeeFirstName() As String
    EmployeeFirstName = m_employee_firstname
End Property

' ReceiverIdentitfierType
Public Property Let ReceiverIdentitfierType(sReceiverIdentitfierType As String)
    m_receiver_identitfier_type = sReceiverIdentitfierType
End Property

Public Property Get ReceiverIdentitfierType() As String
    ReceiverIdentitfierType = m_receiver_identitfier_type
End Property

' ReceiverIdentifier
Public Property Let ReceiverIdentifier(sReceiverIdentifier As String)
    m_receiver_identifier = sReceiverIdentifier
End Property

Public Property Get ReceiverIdentifier() As String
    ReceiverIdentifier = m_receiver_identifier
End Property

' ValidationKyc
Public Property Let ValidationKyc(sValidationKyc As String)
    m_validation_kyc = sValidationKyc
End Property

Public Property Get ValidationKyc() As String
    ValidationKyc = m_validation_kyc
End Property

' ValidationKycValue
Public Property Let ValidationKycValue(sValidationKycValue As String)
    m_validation_kyc_value = sValidationKycValue
End Property

Public Property Get ValidationKycValue() As String
    ValidationKycValue = m_validation_kyc_value
End Property

' Amount
Public Property Let Amount(lAmount As Long)
    m_amount = lAmount
End Property

Public Property Get Amount() As Long
    Amount = m_amount
End Property

' Comment
Public Property Let Comment(sComment As String)
    m_comment = sComment
End Property

Public Property Get Comment() As String
    Comment = m_comment
End Property


' CreatedBy
Public Property Set CreatedBy(oCreatedBy As CActor)
    Set m_created_by = oCreatedBy
End Property

Public Property Get CreatedBy() As CActor
    Set CreatedBy = m_created_by
End Property

' UpdatedBy
Public Property Set UpdatedBy(oUpdatedBy As CActor)
    Set m_updated_by = oUpdatedBy
End Property

Public Property Get UpdatedBy() As CActor
    Set UpdatedBy = m_updated_by
End Property



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(lId As Long, lPaymentId As Long, sEmployeeMatricule As String, sEmployeeLastName As String, sEmployeeFirstName As String, _
sReceiverIdentitfierType As String, sReceiverIdentifier As String, sValidationKyc As String, sValidationKycValue As String, lAmount As Long, sComment As String, _
lCreatorId As Long, dCreatedAt As Date, lUpdatorId As Long, dUpdatedAt As Date)
    
    Me.Id = lId
    Me.PaymentId = lPaymentId
    
    Me.EmployeeMatricule = sEmployeeMatricule
    Me.EmployeeLastName = sEmployeeLastName
    Me.EmployeeFirstName = sEmployeeFirstName
    
    Me.ReceiverIdentitfierType = sReceiverIdentitfierType
    Me.ReceiverIdentifier = sReceiverIdentifier
    
    Me.ValidationKyc = sValidationKyc
    Me.ValidationKycValue = sValidationKycValue
    
    Me.Amount = lAmount
    Me.Comment = sComment
    
    Me.CreatedBy = NewActor(lCreatorId, dCreatedAt)
    Me.UpdatedBy = NewActor(lUpdatorId, dUpdatedAt)
    
End Sub



'*****************************************************************************************************************************
'    Private functions and subs
'*****************************************************************************************************************************


Public Function UpdateDetail(Optional auditAction As Boolean = True) As Boolean
    Dim userData As Variant, sqlRst As String, recordCount As Long, i As Integer, audit As clsAudit
    
    If auditAction Then
        Set audit = loggedUser.StartNewAudit("Modification Mot de Passe Utilisateur (" & Me.Id & ") " & Me.Name)
    End If
    
    sqlRst = "UPDATE users SET userpwd = '" & encrypter.encrypt(sPwd) & "', userpwd_lastupdate_at = '" & Now & "' WHERE Id = " & CStr(Me.Id)
    
    Call PrepareDatabase
    If gobjDB.ExecuteActionQuery(sqlRst) Then
        Me.Pwd = encrypter.encrypt(sPwd)
        Me.PwdLastUpdateAt = Now
        SetPwdExpiry
        
        If auditAction Then
            audit.EndWithSuccess
        End If
        
        SavePwd = True
    Else
        If auditAction Then
            audit.EndWithFailure
        End If
        
        SavePwd = False
    End If
End Function

'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    m_id = 0
End Sub



