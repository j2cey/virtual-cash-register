'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CSettingDataSource
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/05
' Purpose   : Class for Data source Settings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_business_logic As CBusinessLogic
Private m_last_result As CResult

Private Const MODULE_NAME As String = "Setting Data Source Class"
Private Const USER_LOGIN_FIELD_KEY As String = "UserLogin"
Private Const USER_PWD_FIELD_KEY As String = "UserPwd"

Private Const DATA_SOURCE_CLASS_KEY As String = "DataSourceClass"
Private Const DATABASE_CLASS_KEY As String = "DatabaseClass"
Private Const SERVER_OR_PATH_KEY As String = "ServerOrPath"
Private Const DATABASE_OR_FILE_NAME_KEY As String = "DatabaseOrFileName"
Private Const INTEGRATED_SECURITY_KEY As String = "IntegratedSecurity"

'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CSettingDataSource
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
' Property      : DataSourceClass
' Purpose       : The Data source Class
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let DataSourceClass(eDataSourceClass As enuDataSourceClass)
    SetFieldValue DATA_SOURCE_CLASS_KEY, CInt(eDataSourceClass)
End Property

Public Property Get DataSourceClass() As enuDataSourceClass
    DataSourceClass = CInt(GetFieldValue(DATA_SOURCE_CLASS_KEY))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : DatabaseClass
' Purpose       : The Database Class
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/06    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let DatabaseClass(eDatabaseClass As enuDatabaseClass)
    SetFieldValue DATABASE_CLASS_KEY, CInt(eDatabaseClass)
End Property

Public Property Get DatabaseClass() As enuDatabaseClass
    DatabaseClass = CInt(GetFieldValue(DATABASE_CLASS_KEY))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : ServerOrPath
' Purpose       : The Data source Server Name or Path
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ServerOrPath(strServerOrPath As String)
    SetFieldValue SERVER_OR_PATH_KEY, strServerOrPath
End Property

Public Property Get ServerOrPath() As String
    ServerOrPath = CStr(GetFieldValue(SERVER_OR_PATH_KEY))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : DatabaseOrFileName
' Purpose       : The Data source Database Name Or File Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let DatabaseOrFileName(strDatabaseOrFileName As String)
    SetFieldValue DATABASE_OR_FILE_NAME_KEY, strDatabaseOrFileName
End Property

Public Property Get DatabaseOrFileName() As String
    DatabaseOrFileName = CStr(GetFieldValue(DATABASE_OR_FILE_NAME_KEY))
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : IntegratedSecurity
' Purpose       : The Data source Integrated Security Connection, if any
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IntegratedSecurity(blnIntegratedSecurity As Boolean)
    SetFieldValue INTEGRATED_SECURITY_KEY, blnIntegratedSecurity
End Property

Public Property Get IntegratedSecurity() As Boolean
    IntegratedSecurity = CStr(GetFieldValue(INTEGRATED_SECURITY_KEY))
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
            Me.BusinessLogic.GetField(SERVER_OR_PATH_KEY).FieldValue.IsEmpty Or _
            Me.BusinessLogic.GetField(DATABASE_OR_FILE_NAME_KEY).FieldValue.IsEmpty Or _
            Me.BusinessLogic.GetField(DATA_SOURCE_CLASS_KEY).FieldValue.IsEmpty Or _
            Me.BusinessLogic.GetField(DATABASE_CLASS_KEY).FieldValue.IsEmpty Or _
            Me.BusinessLogic.GetField(USER_LOGIN_FIELD_KEY).FieldValue.IsEmpty Or _
            Me.BusinessLogic.GetField(USER_PWD_FIELD_KEY).FieldValue.IsEmpty Or _
            Me.BusinessLogic.GetField(INTEGRATED_SECURITY_KEY).FieldValue.IsEmpty
    
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
    
    ' UserLogin
    Me.BusinessLogic.AddField(NewFieldValueString, USER_LOGIN_FIELD_KEY, "User Login").SetOffsetStart(2).SetOffsetEnd (2)
    ' UserPwd
    Me.BusinessLogic.AddField(NewFieldValueString, USER_PWD_FIELD_KEY, "User Pwd").SetOffsetStart(3).SetOffsetEnd (2)
    ' Data Source Class
    Me.BusinessLogic.AddField(NewFieldValueString, DATA_SOURCE_CLASS_KEY, "Data Source Class Name").SetOffsetStart(4).SetOffsetEnd (2)
    ' Database Class
    Me.BusinessLogic.AddField(NewFieldValueString, DATABASE_CLASS_KEY, "Data Source Class Name").SetOffsetStart(5).SetOffsetEnd (2)
    ' Server or Path
    Me.BusinessLogic.AddField(NewFieldValueString, SERVER_OR_PATH_KEY, "Server or Path").SetOffsetStart(6).SetOffsetEnd (2)
    ' Database or File Name
    Me.BusinessLogic.AddField(NewFieldValueString, DATABASE_OR_FILE_NAME_KEY, "Database or File Name").SetOffsetStart(7).SetOffsetEnd (2)
    ' Integrated Security
    Me.BusinessLogic.AddField(NewFieldValueBoolean, INTEGRATED_SECURITY_KEY, "Integrated Security").SetOffsetStart(8).SetOffsetEnd (2)
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


