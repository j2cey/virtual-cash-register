'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CSettingDataSource
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/05
' Purpose   : Class for Data source Settings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_upper_setting As CSetting

Private Const MODULE_NAME As String = "Setting Data Source Class"
Private Const COLUMN_VALUE As Long = 2

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
'   Property      : UpperSetting
'   Purpose       : The Super Setting object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/14      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set UpperSetting(oUpperSetting As CSetting)
    Set m_upper_setting = oUpperSetting
End Property

Public Property Get UpperSetting() As CSetting
    Set UpperSetting = m_upper_setting
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
    Dim vrnValue As Variant
    
    vrnValue = GetFieldValue(DATA_SOURCE_CLASS_KEY)
    
    If vrnValue = Empty Or CStr(vrnValue) = "" Then
        DataSourceClass = noDataSource
    Else
        DataSourceClass = CInt(vrnValue)
    End If
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
Public Sub Init(ByVal oUpperSetting As CSetting)
    Set Me.UpperSetting = oUpperSetting
    
    InitFields
End Sub

Public Sub LoadValues()
    Me.UpperSetting.LoadValues
End Sub

Public Sub SaveValues()
    Me.UpperSetting.SaveValues
End Sub

Public Function IsComplete() As Boolean
    Dim blnAnyImportantFieldIsEmpty As Boolean
    
    Set Me.UpperSetting.LastResult = SetNewResult("Check Is Complete", MODULE_NAME)
    
    If Not Me.UpperSetting.BusinessLogic Is Nothing Then
        blnAnyImportantFieldIsEmpty = _
            Me.UpperSetting.IsFieldEmpty(SERVER_OR_PATH_KEY) Or _
            Me.UpperSetting.IsFieldEmpty(DATABASE_OR_FILE_NAME_KEY) Or _
            Me.UpperSetting.IsFieldEmpty(DATA_SOURCE_CLASS_KEY) Or _
            Me.UpperSetting.IsFieldEmpty(DATABASE_CLASS_KEY) Or _
            Me.UpperSetting.IsFieldEmpty(USER_LOGIN_FIELD_KEY) Or _
            Me.UpperSetting.IsFieldEmpty(USER_PWD_FIELD_KEY) Or _
            Me.UpperSetting.IsFieldEmpty(INTEGRATED_SECURITY_KEY)
    
        IsComplete = Not blnAnyImportantFieldIsEmpty
    Else
        Me.UpperSetting.LastResult.SetFailure "The Business Logic Object not set yet !", True
        IsComplete = False
    End If
End Function



'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************

Private Sub InitFields()
    ' UserLogin
    Me.UpperSetting.AddField NewFieldValueString, USER_LOGIN_FIELD_KEY, "User Login", COLUMN_VALUE, -1, toTheLeft
    ' UserPwd
    Me.UpperSetting.AddField NewFieldValueString, USER_PWD_FIELD_KEY, "User Pwd", COLUMN_VALUE, -1, toTheLeft
    ' Data Source Class
    Me.UpperSetting.AddField NewFieldValueString, DATA_SOURCE_CLASS_KEY, "Data Source Class Name", COLUMN_VALUE, -1, toTheLeft
    ' Database Class
    Me.UpperSetting.AddField NewFieldValueString, DATABASE_CLASS_KEY, "Data Source Class Name", COLUMN_VALUE, -1, toTheLeft
    ' Server or Path
    Me.UpperSetting.AddField NewFieldValueString, SERVER_OR_PATH_KEY, "Server or Path", COLUMN_VALUE, -1, toTheLeft
    ' Database or File Name
    Me.UpperSetting.AddField NewFieldValueString, DATABASE_OR_FILE_NAME_KEY, "Database or File Name", COLUMN_VALUE, -1, toTheLeft
    ' Integrated Security
    Me.UpperSetting.AddField NewFieldValueBoolean, INTEGRATED_SECURITY_KEY, "Integrated Security", COLUMN_VALUE, -1, toTheLeft
End Sub

Private Sub SetFieldValue(ByVal strFieldKey As String, ByVal varFieldValue As Variant)
    Me.UpperSetting.SetFieldValue strFieldKey, varFieldValue
End Sub

Private Function GetFieldValue(ByVal strFieldKey As String) As Variant
    GetFieldValue = Me.UpperSetting.GetFieldValue(strFieldKey)
End Function


'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    
End Sub


