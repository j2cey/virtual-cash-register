Option Explicit

Private m_user As CUser

Private m_id As Long
Private m_action As String
Private m_action_result As Long

Private m_startedAt As Date
Private m_endedAt As Date

Private m_hostname As String
Private m_hostip As String
Private m_hostuser As String
Private m_hostmodel As String
Private m_hostos As String
Private m_hostosversion As String


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CAudit
    Set Self = Me
End Property

' User
Public Property Set User(oUser As CUser)
    Set m_user = oUser
End Property

Public Property Get User() As CUser
    Set User = m_user
End Property

' Id
Public Property Let id(lId As Long)
    m_id = lId
End Property

Public Property Get id() As Long
    id = m_id
End Property

' Action
Public Property Let Action(sAction As String)
    m_action = sAction
End Property

Public Property Get Action() As String
    Action = m_action
End Property

' ActionResult
Public Property Let ActionResult(lAction As Long)
    m_action_result = lAction
End Property

Public Property Get ActionResult() As Long
    ActionResult = m_action_result
End Property

'StartedAt
Public Property Let StartedAt(dStartedAt As Date)
    m_startedAt = dStartedAt
End Property

Public Property Get StartedAt() As Date
    StartedAt = m_startedAt
End Property

'EndedAt
Public Property Let EndedAt(dEndedAt As Date)
    m_endedAt = dEndedAt
End Property

Public Property Get EndedAt() As Date
    EndedAt = m_endedAt
End Property

' AuditHostName
Public Property Let AuditHostName(sHostName As String)
    m_hostname = sHostName
End Property

Public Property Get AuditHostName() As String
    AuditHostName = m_hostname
End Property

' AuditHostIP
Public Property Let AuditHostIP(sHostIP As String)
    m_hostip = sHostIP
End Property

Public Property Get AuditHostIP() As String
    AuditHostIP = m_hostip
End Property

' AuditHostUser
Public Property Let AuditHostUser(sHostUser As String)
    m_hostuser = sHostUser
End Property

Public Property Get AuditHostUser() As String
    AuditHostUser = m_hostuser
End Property

' AuditHostModel
Public Property Let AuditHostModel(sHostModel As String)
    m_hostmodel = sHostModel
End Property

Public Property Get AuditHostModel() As String
    AuditHostModel = m_hostmodel
End Property

' AuditHostOS
Public Property Let AuditHostOS(sHostOS As String)
    m_hostos = sHostOS
End Property

Public Property Get AuditHostOS() As String
    AuditHostOS = m_hostos
End Property

' AuditHostOSVersion
Public Property Let AuditHostOSVersion(sHostOSVersion As String)
    m_hostosversion = sHostOSVersion
End Property

Public Property Get AuditHostOSVersion() As String
    AuditHostOSVersion = m_hostosversion
End Property



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(oUser As CUser, sAction As String)
    Set Me.User = oUser
    Me.Action = sAction
    Me.ActionResult = 0

    StartAudit
End Sub

Public Sub StartAudit()
    Me.StartedAt = Now
    Me.id = InsertNew
End Sub

Public Sub EndAudit(lActionResult As Long)
    Me.EndedAt = Now
    Me.ActionResult = lActionResult
    SaveAudit
End Sub

Public Sub EndWithSuccess()
    EndAudit (1)
End Sub

Public Sub EndWithFailure()
    EndAudit (-1)
End Sub

'*****************************************************************************************************************************
'    Private functions and subs
'*****************************************************************************************************************************

Private Function InsertNew() As Long
    Dim queryStr As String, newAuditId As Variant
    
    ' Add New Audit
    queryStr = "INSERT INTO audittrail (username,audit_action,started_at,action_result,host_name,host_ip,host_user,host_model,host_os,host_osversion) VALUES ('" _
    & SqlStringVar(Me.User.Name) & "', '" & SqlStringVar(Me.Action) & "', '" & CStr(Me.StartedAt) & "', " & CStr(Me.ActionResult) _
    & ", '" & CStr(Me.AuditHostName) & "', '" & CStr(Me.AuditHostIP) & "', '" & CStr(Me.AuditHostUser) & "'" _
    & ", '" & CStr(Me.AuditHostModel) & "', '" & CStr(Me.AuditHostOS) & "', '" & CStr(Me.AuditHostOSVersion) & "'" _
    & ")"
    
    Call PrepareDatabase
    If gobjDB.ExecuteActionQuery(queryStr, newAuditId) Then
        InsertNew = CLng(newAuditId)
    Else
        InsertNew = -1
    End If
End Function

Private Function SaveAudit() As Boolean
    Dim sqlRst As String, recordCount As Long, i As Integer
    
    sqlRst = "UPDATE audittrail SET username = '" & SqlStringVar(Me.User.Name) & "', audit_action = '" & SqlStringVar(Me.Action) _
    & "', started_at = '" & SqlStringVar(CStr(Me.StartedAt)) & "', ended_at = '" & SqlStringVar(CStr(Me.EndedAt)) _
    & "', action_result = " & CStr(Me.ActionResult) & ", host_name = '" & CStr(Me.AuditHostName) & "', host_ip = '" & CStr(Me.AuditHostIP) & "'" _
    & ", host_user = '" & CStr(Me.AuditHostUser) & "'" & ", host_model = '" & CStr(Me.AuditHostModel) & "'" _
    & ", host_os = '" & CStr(Me.AuditHostOS) & "'" & ", host_osversion = '" & CStr(Me.AuditHostOSVersion) & "'" _
    & " WHERE Id = " & CStr(Me.id)
    
    Call PrepareDatabase
    SaveAudit = gobjDB.ExecuteActionQuery(sqlRst)
End Function


'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    'm_hostname = Environ_GetComputerName
    'm_hostip = GetLocalIP
    'm_hostuser = Environ_GetComputerUser
    BiosGetMainInfos m_hostname, m_hostuser, m_hostip, m_hostmodel, m_hostos, m_hostosversion
End Sub
