Option Explicit

Private m_id As Long

Private m_title As String
Private m_amount As Long
Private m_file_name As String

Private m_status As CPaymentStatus

Private m_statuses_history As Collection
Private m_details As Collection

Private m_created_by As CActor
Private m_updated_by As CActor

'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CPayment
    Set Self = Me
End Property

' Id
Public Property Let Id(lId As Long)
    m_id = lId
End Property

Public Property Get Id() As Long
    Id = m_id
End Property

' Title
Public Property Let Title(sTitle As String)
    m_title = sTitle
End Property

Public Property Get Title() As String
    Title = m_title
End Property

' Amount
Public Property Let Amount(lAmount As Long)
    m_amount = lAmount
End Property

Public Property Get Amount() As Long
    Amount = m_amount
End Property

' FileName
Public Property Let FileName(sFileName As String)
    m_file_name = sFileName
End Property

Public Property Get FileName() As String
    FileName = m_file_name
End Property

' Status
Public Property Set Status(oStatus As CPaymentStatus)
    Set m_status = oStatus
End Property

Public Property Get Status() As CPaymentStatus
    Set Status = m_status
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

' StatusesHistory
Public Property Set StatusesHistory(oStatusesHistory As Collection)
    Set m_statuses_history = oStatusesHistory
End Property

Public Property Get StatusesHistory() As Collection
    Set StatusesHistory = m_statuses_history
End Property

' Details
Public Property Set Details(oDetails As Collection)
    Set m_details = oDetails
End Property

Public Property Get Details() As Collection
    Set Details = m_details
End Property



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(lId As Long, Optional sTitle As String = "", Optional lAmount As Long = -1, Optional sFileName As String = "", Optional lStatusId As Long = -1)
    Me.Id = lId
    
    If sTitle = "" Then
        LoadFromDB
    Else
        Me.Title = sTitle
        Me.Amount = lAmount
        Me.FileName = sFileName
        
        If lStatusId > 0 Then
            Set Me.Status = NewPaymentStatus(lStatusId)
        End If
    End If
    
    LoadStatusesHistoryFromDB
    LoadDetailsFromDB
End Sub

Public Function LoadFromDB() As Boolean
    Dim paymentData As Variant, sqlRst As String, recordCount As Long, i As Integer
    
    If Me.Id > 0 Then
        sqlRst = "SELECT * FROM payments_view WHERE Id = " & Me.Id & ""
        
        Call PrepareDatabase
        paymentData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
        
        If recordCount = 0 Then
            LoadFromDB = False
        Else
            Me.Id = CLng(paymentData(0, 0))
            Me.Title = CStr(paymentData(1, 0))
            Set Me.Status = NewPaymentStatus(CLng(paymentData(2, 0)), CStr(paymentData(3, 0)))
            Me.Amount = CLng(paymentData(4, 0))
            Me.FileName = CStr(paymentData(5, 0))
            
            Set Me.CreatedBy = NewActor(CLng(paymentData(7, 0)), CDate(paymentData(6, 0)))
            Set Me.UpdatedBy = NewActor(CLng(paymentData(9, 0)), CDate(paymentData(8, 0)))
            
            LoadFromDB = True
        End If
    Else
        LoadFromDB = False
    End If
End Function

Public Function LoadStatusesHistoryFromDB(Optional editProgr As clsProgression = Nothing) As Boolean
    Dim statushistoryData As Variant, sqlRst As String, recordCount As Long, i As Integer
    
    If Me.Id > 0 Then
        sqlRst = "SELECT * FROM paymentstatuses WHERE payment_id = " & Me.Id & " ORDER BY status_set_at DESC"
        
        Call PrepareDatabase
        statushistoryData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
        
        If recordCount = 0 Then
            LoadStatusesHistoryFromDB = False
        Else
            
            ClearStatusesHistory
            
            TryStartNewSubProgression editProgr, "Chargement Historique Statut Paiement", recordCount
            For i = 0 To recordCount - 1
                Me.StatusesHistory.Add NewPaymentStatusHistory(CLng(paymentData(0, i)), CLng(paymentData(1, i)), CLng(paymentData(2, i)), CLng(paymentData(3, i)), CDate(paymentData(4, i))), CStr(paymentData(0, i))
                TryAddDoneLastSubProgression editProgr, 1, True
            Next i
            'TryAddDoneProgression editProgr, 1, True
            
            LoadStatusesHistoryFromDB = True
        End If
    Else
        LoadStatusesHistoryFromDB = False
    End If
End Function

Public Function LoadDetailsFromDB(Optional editProgr As clsProgression = Nothing) As Boolean
    Dim detailsData As Variant, sqlRst As String, recordCount As Long, i As Integer
    
    If Me.Id > 0 Then
        sqlRst = "SELECT * FROM paymentdetails WHERE payment_id = " & Me.Id & " ORDER BY Id ASC"
        
        Call PrepareDatabase
        detailsData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
        
        If recordCount = 0 Then
            LoadDetailsFromDB = False
        Else
            
            ClearDetails
            
            TryStartNewSubProgression editProgr, "Chargement Détails Paiement", recordCount
            For i = 0 To recordCount - 1
                Me.AddDetail CLng(detailsData(0, i)), CLng(detailsData(1, i)), CStr(detailsData(2, i)), CStr(detailsData(3, i)), CStr(detailsData(4, i)), CStr(detailsData(5, i)), CStr(detailsData(6, i)), CStr(detailsData(7, i)), CStr(detailsData(8, i)), CLng(detailsData(9, i)), CStr(detailsData(10, i)), CLng(detailsData(11, i)), CDate(detailsData(12, i)), CLng(detailsData(13, i)), CDate(detailsData(14, i))
                TryAddDoneLastSubProgression editProgr, 1, True
            Next i
            'TryAddDoneProgression editProgr, 1, True
            
            LoadDetailsFromDB = True
        End If
    Else
        LoadDetailsFromDB = False
    End If
End Function

Public Function AddDetail(lId As Long, lPaymentId As Long, sEmployeeMatricule As String, sEmployeeLastName As String, sEmployeeFirstName As String, _
sReceiverIdentitfierType As String, sReceiverIdentifier As String, sValidationKyc As String, sValidationKycValue As String, lAmount As Long, sComment As String, _
lCreatorId As Long, dCreatedAt As Date, lUpdatorId As Long, dUpdatedAt As Date)
    
    If lId > 0 Then
        Me.Details.Add NewPaymentDetail(lId, lPaymentId, sEmployeeMatricule, sEmployeeLastName, sEmployeeFirstName, sReceiverIdentitfierType, _
        sReceiverIdentifier, sValidationKyc, sValidationKycValue, lAmount, sComment, lCreatorId, dCreatedAt, lUpdatorId, dUpdatedAt), CStr(lId)
    Else
        Me.Details.Add NewPaymentDetail(lId, lPaymentId, sEmployeeMatricule, sEmployeeLastName, sEmployeeFirstName, sReceiverIdentitfierType, _
        sReceiverIdentifier, sValidationKyc, sValidationKycValue, lAmount, sComment, lCreatorId, dCreatedAt, lUpdatorId, dUpdatedAt)
    End If
End Function

Private Sub ClearStatusesHistory()
    Set m_statuses_history = New Collection
End Sub

Private Sub ClearDetails()
    Set m_details = New Collection
End Sub


'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    m_id = -1
    Set m_statuses_history = New Collection
    Set m_details = New Collection
End Sub


