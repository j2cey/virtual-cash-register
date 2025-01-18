'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CUser
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/09
' Purpose   : A User of the System
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_id As Long
Private m_login As String
Private m_name As String
Private m_role As CUserRole
Private m_pwd As String

Private m_isLogged As Boolean
Private m_loggedAt As Date
Private m_pwdLastUpdateAt As Date
Private m_isPwdExpired As Boolean

Private m_auditsColl As Collection

Private m_record As CRecord


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CUser
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Record
'   Purpose       : The Record of the user
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set Record(oRecord As CRecord)
    Set m_record = oRecord
End Property

Public Property Get Record() As CRecord
    Set Record = m_record
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Id
'   Purpose       : The (database record) Unique Identificator
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Id(lId As Long)
    m_id = lId
End Property

Public Property Get Id() As Long
    Id = m_id
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : LogIn
'   Purpose       : The User LogIn
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let LogIn(sLogin As String)
    m_login = sLogin
End Property

Public Property Get LogIn() As String
    LogIn = m_login
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Name
'   Purpose       : The User Name
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Name(sName As String)
    m_name = sName
End Property

Public Property Get Name() As String
    Name = m_name
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Role
'   Purpose       : The User Role
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set Role(oRole As CUserRole)
    Set m_role = oRole
End Property

Public Property Get Role() As CUserRole
    Set Role = m_role
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Pwd
'   Purpose       : The User Password
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Pwd(sPwd As String)
    m_pwd = sPwd
End Property

Public Property Get Pwd() As String
    Pwd = m_pwd
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IsLogged
'   Purpose       : Determine whether the User is Logged (correctly)
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsLogged(bIsLogged As Boolean)
    m_isLogged = bIsLogged
End Property

Public Property Get IsLogged() As Boolean
    IsLogged = m_isLogged
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : LoggedAt
'   Purpose       : User last Login date
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let LoggedAt(dLoggedAt As Date)
    m_loggedAt = dLoggedAt
End Property

Public Property Get LoggedAt() As Date
    LoggedAt = m_loggedAt
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : PwdLastUpdateAt
'   Purpose       : User last password update date
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let PwdLastUpdateAt(dPwdLastUpdateAt As Date)
    m_pwdLastUpdateAt = dPwdLastUpdateAt
End Property

Public Property Get PwdLastUpdateAt() As Date
    PwdLastUpdateAt = m_pwdLastUpdateAt
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IsPwdExpired
'   Purpose       : Determine whether the User password is expired
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsPwdExpired(bIsPwdExpired As Boolean)
    m_isPwdExpired = bIsPwdExpired
End Property

Public Property Get IsPwdExpired() As Boolean
    IsPwdExpired = m_isPwdExpired
End Property

' AuditsList
Public Property Get AuditsList() As Collection
    Set AuditsList = m_auditsColl
End Property


'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(lId As Long, Optional oRecord As CRecord = Nothing, Optional sLogin As String = "", Optional sName As String = "")
    Me.Id = lId
    
    If Not oRecord Is Nothing Then
        Set Me.Record = oRecord
    End If
    
    'If sLogin = "" Then
    '    LoadFromDB
    'Else
    '    Me.LogIn = sLogin
    '    Me.Name = sName
    'End If
End Sub

Public Function TryLogIn(sLogin As String, sPwd As String) As Boolean
    Dim userData As Variant, sqlRst As String, recordCount As Long, i As Integer, audit As CAudit
    
    Set audit = Me.StartNewAudit("Connexion à l'Application, Login: " & sLogin)
    
    sqlRst = "SELECT * FROM users_view WHERE userlogin = '" & SqlStringVar(sLogin) & "' AND userpwd = '" & encrypter.encrypt(sPwd) & "'"
    
    Call PrepareDatabase
    userData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    If recordCount = 0 Then
        MsgBox "Login et/ou Mot de Passe Incorrect !", vbCritical, GetAppName
        Me.IsLogged = False
        audit.EndWithFailure
    Else
        SetFromDB userData
        
        Me.LoggedAt = Now
        Me.IsLogged = True
        audit.EndWithSuccess
    End If
End Function

Public Function Logout() As Boolean
    Dim audit As CAudit
    
    Set audit = Me.StartNewAudit("Déconnexion de l'Application")
    audit.EndWithSuccess
End Function

Public Sub SetRole(lId As Long, sTitle As String)
    Set Me.Role = NewUserRole(lId, sTitle)
End Sub

Public Function SavePwd(sPwd As String, Optional auditAction As Boolean = True) As Boolean
    Dim userData As Variant, sqlRst As String, recordCount As Long, i As Integer, audit As CAudit
    
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

Public Function IsPwdSameAs(clearPwdToCheck As String)
    IsPwdSameAs = Not (Me.Pwd <> encrypter.encrypt(clearPwdToCheck))
End Function

Public Function SaveLastLogin(Optional auditAction As Boolean = True) As Boolean
    Dim userData As Variant, sqlRst As String, recordCount As Long, i As Integer, audit As CAudit
    
    If auditAction Then
        Set audit = loggedUser.StartNewAudit("Modification Derniere date Login Utilisateur (" & Me.Id & ") " & Me.Name)
    End If
    
    sqlRst = "UPDATE users SET user_lastlogin_at = '" & Now & "' WHERE Id = " & CStr(Me.Id)
    
    Call PrepareDatabase
    If gobjDB.ExecuteActionQuery(sqlRst) Then
        Me.LoggedAt = Now
        
        If auditAction Then
            audit.EndWithSuccess
        End If
        
        SaveLastLogin = True
    Else
        If auditAction Then
            audit.EndWithFailure
        End If
        
        SaveLastLogin = False
    End If
End Function

Public Function HasRole(sRole As String) As Boolean
    If Me.Role Is Nothing Then
        HasRole = False
        Exit Function
    End If
    
    HasRole = (Me.Role.Title = sRole)
End Function

Public Function Can(arrPermissions As Variant) As Boolean
    Dim i As Integer
    
    If Me.Role Is Nothing Then
        Can = False
        Exit Function
    End If
    
    For i = 0 To UBound(arrPermissions)
        If Me.Role.HasPermission(CStr(arrPermissions(i))) Then
            Can = True
            Exit Function
        End If
    Next i
    
    Can = False
End Function

Public Function StartNewAudit(sAction As String) As CAudit
    Dim oAudit As CAudit
    
    Set oAudit = NewAudit(Me, sAction)
    
    ' Add Audit to The Audits Collection
    Me.AuditsList.Add oAudit, CStr(oAudit.Id)
    
    Set StartNewAudit = oAudit
End Function


'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Private Function LoadFromDB() As Boolean
    Dim userData As Variant, sqlRst As String, recordCount As Long, i As Integer
    
    sqlRst = "SELECT * FROM users_view WHERE Id = " & Me.Id & " "
    
    Call PrepareDatabase
    userData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    If recordCount = 0 Then
        GetById = False
    Else
        SetFromDB userData
        
        GetById = True
    End If
End Function

Private Sub SetFromDB(userData As Variant)
    Me.Id = CLng(userData(0, 0))
    Me.LogIn = sLogin
    Me.Name = CStr(userData(2, 0))
    Me.SetRole CLng(userData(4, 0)), CStr(userData(5, 0))
    Me.Pwd = CStr(userData(3, 0))
    
    Me.PwdLastUpdateAt = CDate(userData(7, 0))
    SetPwdExpiry
End Sub


Private Sub SetPwdExpiry()
    Dim result As Long, validitydays As Long
    
    validitydays = GetPwdValidity
    If validitydays < 1 Then
        Me.IsPwdExpired = False
        Exit Sub
    End If
    
    result = DateDiff("D", Me.PwdLastUpdateAt, Now)
    Me.IsPwdExpired = result > validitydays
End Sub


'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    m_name = "Invité"
    Set m_auditsColl = New Collection
End Sub

