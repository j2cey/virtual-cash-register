'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CModelUser
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/09
' Purpose   : A User of the System
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_id As Long
Private m_upper_model As CModel

Private Const MODULE_NAME As String = "Model User"

' Login
Private m_login As String
Private Const LOGIN_FIELD_SAVING_KEY As String = "userlogin"
Private Const LOGIN_FIELD_SELECT_KEY As String = "userlogin"
Private Const LOGIN_FIELD_LABEL_KEY As String = "User Login"

'Pwd
Private m_pwd As String
Private Const PWD_FIELD_SAVING_KEY As String = "userpwd"
Private Const PWD_FIELD_SELECT_KEY As String = "userpwd"
Private Const PWD_FIELD_LABEL_KEY As String = "User Password"

' Name
'Private m_name As String
Private Const USERNAME_FIELD_SAVING_KEY As String = "username"
Private Const USERNAME_FIELD_SELECT_KEY As String = "username"
Private Const USERNAME_FIELD_LABEL_KEY As String = "User Name"

' Role
Private m_role As CModelUserRole

Private m_is_logged As Boolean
Private m_logged_at As Date

' LastLoginAt
Private Const LAST_LOGIN_AT_FIELD_SAVING_KEY As String = "user_lastlogin_at"
Private Const LAST_LOGIN_AT_FIELD_SELECT_KEY As String = "user_lastlogin_at"
Private Const LAST_LOGIN_AT_FIELD_LABEL_KEY As String = "User Password Last Update Date"

' PwdLastUpdateAt
'Private m_pwd_last_update_at As Date
Private Const PWD_LAST_UPDATE_AT_FIELD_SAVING_KEY As String = "userpwd_lastupdate_at"
Private Const PWD_LAST_UPDATE_AT_FIELD_SELECT_KEY As String = "userpwd_lastupdate_at"
Private Const PWD_LAST_UPDATE_AT_FIELD_LABEL_KEY As String = "User Last Login Date"

Private m_is_pwd_expired As Boolean

Private m_audits As Collection


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CModelUser
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : UpperModel
'   Purpose       : The Upper Model Object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set UpperModel(oUpperModel As CModel)
    Set m_upper_model = oUpperModel
End Property

Public Property Get UpperModel() As CModel
    Set UpperModel = m_upper_model
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Id
'   Purpose       : The (database record) Unique Identificator
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Id(lngId As Long)
    If Not Me.UpperModel Is Nothing Then
        Me.UpperModel.SetId lngId
    Else
        m_id = lngId
    End If
End Property

Public Property Get Id() As Long
    If Not Me.UpperModel Is Nothing Then
        Id = Me.UpperModel.GetId()
    Else
        Id = m_id
    End If
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : LogIn
'   Purpose       : The User LogIn
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let LogIn(strLogin As String)
    If Not Me.UpperModel Is Nothing Then
        Me.UpperModel.SetFieldValue NewFieldValueString, LOGIN_FIELD_SAVING_KEY, LOGIN_FIELD_LABEL_KEY, LOGIN_FIELD_SELECT_KEY, strLogin
    Else
        m_login = strLogin
    End If
End Property

Public Property Get LogIn() As String
    Dim vrnValue As Variant
    
    vrnValue = Me.UpperModel.GetFieldValue(NewFieldValueString, LOGIN_FIELD_SAVING_KEY, LOGIN_FIELD_LABEL_KEY, LOGIN_FIELD_SELECT_KEY)
    
    LogIn = IIf(vrnValue = Null, Null, CStr(vrnValue))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Pwd
'   Purpose       : The User Password
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Pwd(strPwd As String)
    If Not Me.UpperModel Is Nothing Then
        Me.UpperModel.SetFieldValue NewFieldValueString, PWD_FIELD_SAVING_KEY, PWD_FIELD_LABEL_KEY, PWD_FIELD_SELECT_KEY, strPwd
    Else
        m_pwd = strPwd
    End If
End Property

Public Property Get Pwd() As String
    Dim vrnValue As Variant
    
    vrnValue = Me.UpperModel.GetFieldValue(NewFieldValueString, PWD_FIELD_SAVING_KEY, PWD_FIELD_LABEL_KEY, PWD_FIELD_SELECT_KEY)
    
    Pwd = IIf(vrnValue = Null, Null, CStr(vrnValue))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Name
'   Purpose       : The User Name
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Name(strName As String)
    If Not Me.UpperModel Is Nothing Then
        Me.UpperModel.SetFieldValue NewFieldValueString, USERNAME_FIELD_SAVING_KEY, USERNAME_FIELD_LABEL_KEY, USERNAME_FIELD_LABEL_KEY, strName
    End If
End Property

Public Property Get Name() As String
    Dim vrnValue As Variant
    
    vrnValue = Me.UpperModel.GetFieldValue(NewFieldValueString, USERNAME_FIELD_SAVING_KEY, USERNAME_FIELD_LABEL_KEY, USERNAME_FIELD_LABEL_KEY)
    
    Name = IIf(vrnValue = Null, Null, CStr(vrnValue))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Role
'   Purpose       : The User Role
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set Role(oRole As CModelUserRole)
    Set m_role = oRole
End Property

Public Property Get Role() As CModelUserRole
    Set Role = m_role
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
    m_is_logged = bIsLogged
End Property

Public Property Get IsLogged() As Boolean
    IsLogged = m_is_logged
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : LoggedAt
'   Purpose       : User last Login date
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let LoggedAt(dteLoggedAt As Date)
    m_logged_at = dteLoggedAt
End Property

Public Property Get LoggedAt() As Date
    LoggedAt = m_logged_at
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : LastLoginAt
'   Purpose       : User last login date
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/15      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let LastLoginAt(dtePwdLastUpdateAt As Date)
    If Not Me.UpperModel Is Nothing Then
        Me.UpperModel.SetFieldValue NewFieldValueDate, LAST_LOGIN_AT_FIELD_SAVING_KEY, LAST_LOGIN_AT_FIELD_LABEL_KEY, LAST_LOGIN_AT_FIELD_SELECT_KEY, dtePwdLastUpdateAt
    End If
End Property

Public Property Get LastLoginAt() As Date
    Dim vrnValue As Variant
    
    vrnValue = Me.UpperModel.GetFieldValue(NewFieldValueString, LAST_LOGIN_AT_FIELD_SAVING_KEY, LAST_LOGIN_AT_FIELD_LABEL_KEY, LAST_LOGIN_AT_FIELD_SELECT_KEY)
    
    LastLoginAt = IIf(vrnValue = Null, Null, CDate(vrnValue))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : PwdLastUpdateAt
'   Purpose       : User last password update date
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let PwdLastUpdateAt(dtePwdLastUpdateAt As Date)
    If Not Me.UpperModel Is Nothing Then
        Me.UpperModel.SetFieldValue NewFieldValueDate, PWD_LAST_UPDATE_AT_FIELD_SAVING_KEY, PWD_LAST_UPDATE_AT_FIELD_LABEL_KEY, PWD_LAST_UPDATE_AT_FIELD_SELECT_KEY, dtePwdLastUpdateAt
    End If
End Property

Public Property Get PwdLastUpdateAt() As Date
    Dim vrnValue As Variant
    
    vrnValue = Me.UpperModel.GetFieldValue(NewFieldValueString, PWD_LAST_UPDATE_AT_FIELD_SAVING_KEY, PWD_LAST_UPDATE_AT_FIELD_LABEL_KEY, PWD_LAST_UPDATE_AT_FIELD_SELECT_KEY)
    
    PwdLastUpdateAt = IIf(vrnValue = Null, Null, CDate(vrnValue))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IsPwdExpired
'   Purpose       : Determine whether the User password is expired
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsPwdExpired(blnIsPwdExpired As Boolean)
    m_is_pwd_expired = blnIsPwdExpired
End Property

Public Property Get IsPwdExpired() As Boolean
    IsPwdExpired = m_is_pwd_expired
End Property

' AuditsList
Public Property Get AuditsList() As Collection
    Set AuditsList = m_audits
End Property


'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(Optional ByVal lId As Long = -1, Optional ByVal oModel As CModel = Nothing, Optional ByVal sLogin As String = "", Optional ByVal sName As String = "", Optional ByVal sPwd As String = "")
    Me.Id = lId
    
    If Not oModel Is Nothing Then
        Set Me.UpperModel = oModel
    End If
    
    'If sLogin = "" Then
    '    LoadFromDB
    'Else
    '    Me.LogIn = sLogin
    '    Me.Name = sName
    'End If
End Sub

Public Sub LoadValues()
    If Not Me.UpperModel Is Nothing Then
        Me.UpperModel.LoadValues
    End If
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
'    Private functions and subs
'*****************************************************************************************************************************

Private Sub SetFieldValue(ByVal strFieldKey As String, ByVal varFieldValue As Variant)
    Me.UpperModel.SetFieldValue strFieldKey, varFieldValue
End Sub

Private Function GetFieldValue(ByVal oFieldType As IFieldValue, ByVal strNameForSaving As String, ByVal strFieldLabel As String, ByVal strNameForSelecting As String) As Variant
    GetFieldValue = Me.UpperModel.GetFieldValue(oFieldType, strNameForSaving, strFieldLabel, strNameForSelecting)
End Function



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
    Me.Name = "Invité"
    Set m_audits = New Collection
End Sub

