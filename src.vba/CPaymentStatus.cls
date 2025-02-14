Option Explicit

Private m_id As Long
Private m_title As String
Private m_description As String

Private m_created_by As CActor
Private m_updated_by As CActor


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CPaymentStatus
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

' Description
Public Property Let Description(sDescription As String)
    m_description = sDescription
End Property

Public Property Get Description() As String
    Description = m_description
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

Public Sub Init(lId As Long, Optional sTitle As String = "", Optional sDescription As String = "")
    Me.Id = lId
    
    If sTitle = "" Then
        Me.LoadFromDB
    Else
        Me.Title = sTitle
        Me.Description = sDescription
    End If
    
End Sub

Public Function LoadFromDB() As Boolean
    Dim paymentstatusData As Variant, sqlRst As String, recordCount As Long, i As Integer
    
    If Me.Id > 0 Then
        sqlRst = "SELECT * FROM paymentstatus WHERE Id = " & Me.Id & ""
        
        Call PrepareDatabase
        paymentstatusData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
        
        If recordCount = 0 Then
            LoadFromDB = False
        Else
            Me.Id = CLng(paymentstatusData(0, 0))
            Me.Title = CStr(paymentstatusData(1, 0))
            Me.Description = CStr(paymentstatusData(2, 0))
            Set Me.CreatedBy = NewActor(CLng(paymentstatusData(4, 0)), CDate(paymentData(3, 0)))
            Set Me.UpdatedBy = NewActor(CLng(paymentstatusData(6, 0)), CDate(paymentData(5, 0)))
            
            LoadFromDB = True
        End If
    Else
        LoadFromDB = False
    End If
End Function



'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    m_id = 0
End Sub


