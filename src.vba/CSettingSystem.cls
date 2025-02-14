'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CSettingSystem
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/05
' Purpose   : Class for Data source Settings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_business_logic As CBusinessLogic
Private m_last_result As CResult

Private Const MODULE_NAME As String = "Setting System"
Private Const USER_LOGIN_FIELD_KEY As String = "UserLogin"
Private Const USER_PWD_FIELD_KEY As String = "UserPwd"

Private Const APP_NAME_KEY As String = "AppName"
Private Const MAIL_PROVIDER_KEY As String = "MailProvider"

'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CSettingSystem
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : BusinessLogic
'   Purpose       : The Business Logic object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/09      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set BusinessLogic(oBusinessLogic As CBusinessLogic)
    Set m_business_logic = oBusinessLogic
End Property

Public Property Get BusinessLogic() As CBusinessLogic
    Set BusinessLogic = m_business_logic
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : LastResult
'   Purpose       : The Last Result
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/08      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set LastResult(oLastResult As CResult)
    Set m_last_result = oLastResult
End Property

Public Property Get LastResult() As CResult
    Set LastResult = m_last_result
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : UserLogin
' Purpose       : The Data source operating user login
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let UserLogin(strUserLogin As String)
    SetFieldValue USER_LOGIN_FIELD_KEY, strUserLogin
End Property

Public Property Get UserLogin() As String
    UserLogin = CStr(GetFieldValue(USER_LOGIN_FIELD_KEY))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : UserPwd
' Purpose       : The Data source operating user password
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let UserPwd(strUserPwd As String)
    SetFieldValue USER_PWD_FIELD_KEY, strUserPwd
End Property

Public Property Get UserPwd() As String
    UserPwd = CStr(GetFieldValue(USER_PWD_FIELD_KEY))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : AppName
' Purpose       : The Application Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/12    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let AppName(strAppName As String)
    SetFieldValue APP_NAME_KEY, CStr(strAppName)
End Property

Public Property Get AppName() As String
    AppName = CStr(GetFieldValue(APP_NAME_KEY))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : MailProvider
' Purpose       : The Application s MailProvider
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/12    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let MailProvider(strMailProvider As String)
    SetFieldValue MAIL_PROVIDER_KEY, CStr(strMailProvider)
End Property

Public Property Get MailProvider() As String
    MailProvider = CStr(GetFieldValue(MAIL_PROVIDER_KEY))
End Property


'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : blnIsOK             Tells if the operation has been well performed
'                   lngCode             The Result Code
'                   strMessage          The Result Message
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal oUser As CUser, ByVal oDataSource As IDataSource, ByVal strRecordTable As String)
    
    Call InitBusinessLogic(oUser, oDataSource, strRecordTable)
End Sub

Public Sub LoadValues()
    Dim result As CResult, oRecordValue As CRecord
    
    Set result = Me.BusinessLogic.ReadValues(False)
    
    If Not result.IsOK Then
        result.Notify
    End If
End Sub

Public Sub SaveValues()
    Dim result As CResult, oRecordValue As CRecord
    
    Set result = Me.BusinessLogic.Save(False)
    
    If Not result.IsOK Then
        result.Notify
    End If
End Sub

Public Function IsComplete() As Boolean
    Dim blnAnyImportantFieldIsEmpty As Boolean
    
    Set Me.LastResult = SetNewResult("Check Is Complete", MODULE_NAME)
    
    If Not Me.BusinessLogic Is Nothing Then
        blnAnyImportantFieldIsEmpty = _
            Me.BusinessLogic.GetField(APP_NAME_KEY).FieldValue.IsEmpty Or _
            Me.BusinessLogic.GetField(MAIL_PROVIDER_KEY).FieldValue.IsEmpty
    
        IsComplete = Not blnAnyImportantFieldIsEmpty
    Else
        Me.LastResult.SetFailure "The Business Logic Object not set yet !", True
        IsComplete = False
    End If
End Function



'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************

Private Sub InitBusinessLogic(ByVal oUser As CUser, ByVal oDataSource As IDataSource, ByVal strRecordTable As String)
    Set Me.BusinessLogic = NewBusinessLogic(oUser, NewDataAccess(oUser, oDataSource, strRecordTable))
    
    ' TODO: From Excel, Get the first cell empty, and use it as start row for offset
    ' UserLogin
    Me.BusinessLogic.AddField(NewFieldValueString, APP_NAME_KEY, "User Login").SetOffsetStart(2).SetOffsetEnd (2)
    ' UserPwd
    Me.BusinessLogic.AddField(NewFieldValueString, MAIL_PROVIDER_KEY, "User Pwd").SetOffsetStart(3).SetOffsetEnd (2)
End Sub

Private Sub SetFieldValue(ByVal strFieldKey As String, ByVal varFieldValue As Variant)
    Set Me.LastResult = SetNewResult("Set Field Value", MODULE_NAME)
    
    If Not Me.BusinessLogic Is Nothing Then
        Set Me.LastResult = Me.BusinessLogic.SetFieldValue(strFieldKey, varFieldValue)
    Else
        Me.LastResult.SetFailure "The Business Logic Object not set yet !", True
    End If
End Sub

Private Function GetFieldValue(ByVal strFieldKey As String) As Variant
    Set Me.LastResult = SetNewResult("Get Field Value", MODULE_NAME)
    
    If Not Me.BusinessLogic Is Nothing Then
        GetFieldValue = Me.BusinessLogic.GetFieldValue(strFieldKey).GetValue()
    Else
        Me.LastResult.SetFailure "The Business Logic Object not set yet !", True
    End If
End Function


'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    
End Sub




