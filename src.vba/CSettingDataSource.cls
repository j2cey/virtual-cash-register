'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CSettingDataSource
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/05
' Purpose   : Class for Data source Settings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_record As CRecord
Private m_data_access As CDataAccess

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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : DataAccess
'   Purpose       : The Data Access object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set DataAccess(oDataAccess As CDataAccess)
    Set m_data_access = oDataAccess
End Property

Public Property Get DataAccess() As CDataAccess
    Set DataAccess = m_data_access
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : Record
' Purpose       : The Record
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set Record(oRecord As CRecord)
    Set m_record = oRecord
End Property

Public Property Get Record() As CRecord
    Set Record = m_record
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
    
    Call InitRecord(oUser, oDataSource, strRecordTable)
End Sub

Public Sub LoadValues()
    Dim result As CResult, oRecordValue As CRecord
    
    Set oRecordValue = Me.DataAccess.GetRecord(result, True)
    
    If result.IsOK Then
        Dim fdCount As Long
        
        ' TODO: Set The get record to the DataAccess s Record
        Set Me.DataAccess.Record.FieldList = oRecordValue.FieldList
        
        fdCount = Me.DataAccess.Record.FieldList.FieldList.Count
    End If
End Sub

Public Sub SaveValues()
    Dim result As CResult, oRecordValue As CRecord
    
    Set result = Me.DataAccess.Save()
    
End Sub

Public Function IsComplete() As Boolean
    Dim blnAnyImportantFieldIsEmpty As Boolean
    
    blnAnyImportantFieldIsEmpty = _
        GetField(SERVER_OR_PATH_KEY).FieldValue.IsEmpty Or _
        GetField(DATABASE_OR_FILE_NAME_KEY).FieldValue.IsEmpty Or _
        GetField(DATA_SOURCE_CLASS_KEY).FieldValue.IsEmpty Or _
        GetField(DATABASE_CLASS_KEY).FieldValue.IsEmpty Or _
        GetField(USER_LOGIN_FIELD_KEY).FieldValue.IsEmpty Or _
        GetField(USER_PWD_FIELD_KEY).FieldValue.IsEmpty Or _
        GetField(INTEGRATED_SECURITY_KEY).FieldValue.IsEmpty

    IsComplete = Not blnAnyImportantFieldIsEmpty
End Function



'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************

Private Sub InitRecord(ByVal oUser As CUser, ByVal oDataSource As IDataSource, ByVal strRecordTable As String)
    'Set Me.Record = NewRecord(NewDataAccess(oUser, oDataSource, strRecordTable), oUser)
    Set Me.DataAccess = NewDataAccess(oUser, oDataSource, strRecordTable)
    
    ' UserLogin
    Me.DataAccess.Record.FieldList.AddField NewField(NewFieldValueString, USER_LOGIN_FIELD_KEY, "User Login").SetOffsetStart(2).SetOffsetEnd(2)
    ' UserPwd
    Me.DataAccess.Record.FieldList.AddField NewField(NewFieldValueString, USER_PWD_FIELD_KEY, "User Pwd").SetOffsetStart(3).SetOffsetEnd(2)
    ' Data Source Class
    Me.DataAccess.Record.FieldList.AddField NewField(NewFieldValueString, DATA_SOURCE_CLASS_KEY, "Data Source Class Name").SetOffsetStart(4).SetOffsetEnd(2)
    ' Database Class
    Me.DataAccess.Record.FieldList.AddField NewField(NewFieldValueString, DATABASE_CLASS_KEY, "Data Source Class Name").SetOffsetStart(5).SetOffsetEnd(2)
    ' Server or Path
    Me.DataAccess.Record.FieldList.AddField NewField(NewFieldValueString, SERVER_OR_PATH_KEY, "Server or Path").SetOffsetStart(6).SetOffsetEnd(2)
    ' Database or File Name
    Me.DataAccess.Record.FieldList.AddField NewField(NewFieldValueString, DATABASE_OR_FILE_NAME_KEY, "Database or File Name").SetOffsetStart(7).SetOffsetEnd(2)
    ' Integrated Security
    Me.DataAccess.Record.FieldList.AddField NewField(NewFieldValueBoolean, INTEGRATED_SECURITY_KEY, "Integrated Security").SetOffsetStart(8).SetOffsetEnd(2)
End Sub

Private Sub SetFieldValue(ByVal strFieldKey As String, ByVal varFieldValue As Variant)
    Me.DataAccess.Record.FieldList.GetFieldByName(strFieldKey).SetValue (varFieldValue)
End Sub

Private Function GetField(ByVal strFieldKey As String) As CField
    Set GetField = Me.DataAccess.Record.FieldList.GetFieldByName(strFieldKey)
End Function

Private Function GetFieldValue(ByVal strFieldKey As String) As Variant
    GetFieldValue = GetField(strFieldKey).FieldValue.GetValue()
End Function


'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    
End Sub


