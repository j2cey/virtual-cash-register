'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CDataSourceDatabase
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/18
' Purpose   : Database Data Source
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_user As CModelUser
Private m_server_or_path As String
Private m_database_or_file_name As String

Private m_connection As CDatabaseConnection
Private m_command As CDatabaseCommand
Private m_request_recordset As New ADODB.Recordset

Private m_integrated_security As Boolean
Private m_affected_record_id As Long

Private m_current_data_access As CDataAccess
Private m_inner_database As IDataSourceDatabase

Private Const MODULE_NAME As String = "Data-Source Database"

Implements IDataSource

'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CDataSourceDatabase
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : InnerDatabase
'   Purpose       : The Inner Database
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set InnerDatabase(oInnerDatabase As IDataSourceDatabase)
    Set m_inner_database = oInnerDatabase
End Property

Public Property Get InnerDatabase() As IDataSourceDatabase
    Set InnerDatabase = m_inner_database
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : ServerOrPath
'   Purpose       : The server or database file path
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ServerOrPath(strServerOrPath As String)
    Let m_server_or_path = strServerOrPath
End Property

Public Property Get ServerOrPath() As String
    ServerOrPath = m_server_or_path
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : User
'   Purpose       : The performing user
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set User(oUser As CModelUser)
    Set m_user = oUser
End Property

Public Property Get User() As CModelUser
    Set User = m_user
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : DatabaseOrFileName
'   Purpose       : The Database Or File Name
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let DatabaseOrFileName(strDatabaseOrFileName As String)
    Let m_database_or_file_name = strDatabaseOrFileName
End Property

Public Property Get DatabaseOrFileName() As String
    DatabaseOrFileName = m_database_or_file_name
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : Connection
' Purpose       : The Connection Object
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/19    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set Connection(oConnection As CDatabaseConnection)
    Set m_connection = oConnection
End Property

Public Property Get Connection() As CDatabaseConnection
    Set Connection = m_connection
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : Command
' Purpose       : The Request Command
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/21    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set Command(oCommand As CDatabaseCommand)
    Set m_command = oCommand
End Property

Public Property Get Command() As CDatabaseCommand
    Set Command = m_command
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : RequestRecordset
' Purpose       : The Request Recordset
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/21    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set RequestRecordset(oRequestRecordset As ADODB.Recordset)
    Set m_request_recordset = oRequestRecordset
End Property

Public Property Get RequestRecordset() As ADODB.Recordset
    Set RequestRecordset = m_request_recordset
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IntegratedSecurity
'   Purpose       : Determine whether the Integrated security must be used
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IntegratedSecurity(blnIntegratedSecurity As Boolean)
    Let m_integrated_security = blnIntegratedSecurity
End Property

Public Property Get IntegratedSecurity() As Boolean
    IntegratedSecurity = m_integrated_security
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : AffectedRecordId
'   Purpose       : Affected record id after command execution
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let AffectedRecordId(lngAffectedRecordId As Long)
    Let m_affected_record_id = lngAffectedRecordId
End Property

Public Property Get AffectedRecordId() As Long
    AffectedRecordId = m_affected_record_id
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : CurrentDataAccess
'   Purpose       : Current Data Access Object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/15      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set CurrentDataAccess(oCurrentDataAccess As CDataAccess)
    Set m_current_data_access = oCurrentDataAccess
End Property

Public Property Get CurrentDataAccess() As CDataAccess
    Set CurrentDataAccess = m_current_data_access
End Property



'***************************************************************************************************************************************************************
'   Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : oRecordDA           The Data Access object
'                   oUser               The Performing User
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal oUser As CModelUser, ByVal strServerOrPath As String, ByVal strDatabaseOrFileName As String, Optional ByVal eDatabaseClass As enuDatabaseClass = noDatabase, Optional ByVal blnIntegratedSecurity As Boolean = True)
    
    Set Me.User = oUser
    Me.ServerOrPath = strServerOrPath
    Me.DatabaseOrFileName = strDatabaseOrFileName
    
    Me.IntegratedSecurity = blnIntegratedSecurity
    
    If Not eDatabaseClass = noDatabase Then
        Me.SetInnerDatabase eDatabaseClass
    End If
End Sub

Public Function SetInnerDatabase(ByVal eDatabaseClass As enuDatabaseClass) As CDataSourceDatabase
    Dim oInnerDatabase As IDataSourceDatabase
    
    Select Case eDatabaseClass
        Case access2007
          Set oInnerDatabase = New CDatabaseAccess2007
        Case sqlserver2014
          Set oInnerDatabase = New CDatabaseSqlServer2014
    End Select
    
    Set Me.InnerDatabase = oInnerDatabase
    Set Me.InnerDatabase.UpperDatabase = Me
    
    Set SetInnerDatabase = Me
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : ShowErrorMessages
'   Purpose       : IConstruct a comprehensive error message based on the passed objects.
'   Arguments     : StandardErrorObject
'                   SourceModule
'                   SourceMethod
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowErrorMessages(ByVal StandardErrorObject As VBA.ErrObject, _
    ByVal SourceModule As String, ByVal SourceMethod As String)
    
    Dim strMsg As String
    'Handle the standard error, if any.
    If StandardErrorObject.Number <> 0 Then
        strMsg = "Error: " & CStr(Err.Number) & vbCrLf & vbCrLf & "Description: " & _
        Err.Description & vbCrLf & vbCrLf
    End If
    'Include the database-related errors, if any.
    If Not (gobjDB Is Nothing) Then
        strMsg = strMsg & gobjDB.ADOErrors()
    End If
    'Finally, tack on the module/method names.
    strMsg = strMsg & SourceModule & "::" & SourceMethod & vbCrLf
    
    Err.Clear
    MsgBox strMsg, vbCritical, "Contact Technical Support For Assistance"
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : BuildConnectionString
'   Purpose     : Construct a full connection string, set local properties, and
'                 return string to calling routine.
'
'   Arguments   :
'
'   Returns     : String
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/19  Jude Parfait        Created
'   2025/01/21  Jude Parfait        Set as funtion to return the connection string string
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildConnectionString() As String
    Dim rslt As CResult, strConnStr As String
    
    Set rslt = InnerDatabase.BuildConnectionString(strConnStr)
    Me.Connection.ConnectionString = strConnStr
    
    BuildConnectionString = strConnStr
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : OpenDatabase
'   Purpose     : Open a connection to the specified server and database, if not already open.
'                 If the parameters are not specified, then just re-use the last connection string that was created.
'   Arguments   :
'
'   Returns     : Boolean
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/19  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenDatabase() As Boolean
    OpenDatabase = Me.Connection.OpenConnection()
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : CloseDatabase
'   Purpose     : Close connection to database.
'   Arguments   :
'
'   Returns     : Boolean
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/20  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CloseDatabase() As Boolean
    If Me.Connection Is Nothing Then
        CloseDatabase = False
    Else
        CloseDatabase = CloseDatabase = Me.Connection.CloseConnection()
    End If
End Function



'***************************************************************************************************************************************************************
'   Private Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : CloseRecordset
'   Purpose     : Close the recordset.
'   Arguments   :
'
'   Returns     : Boolean
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/31  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CloseRecordset() As Boolean
    If Me.RequestRecordset Is Nothing Then
        CloseRecordset = False
    Else
        Set Me.RequestRecordset = Nothing
        CloseRecordset = True
    End If
End Function

Private Function SetValueField() As CResult
    
End Function

Private Function SetRecordIdValue(ByRef oRecordValue As CRecord) As CResult
    Dim result As CResult, strIdName As String
    
    strIdName = "Id"
    Set result = NewResult(MODULE_NAME, "SetIdFieldValue", False)
    
    With Me.RequestRecordset
        If Not IsNull(.fields(strIdName).Value) Then
            oRecordValue.RecordId = CLng(.fields(strIdName).Value)
            result.SetSuccess
        Else
            result.SetFailure "field " & strIdName & " NOT in Record"
        End If
    End With
    
    Set SetRecordIdValue = result
End Function

Private Function SetSelectableFieldValue(ByVal oFieldValue As CField) As CResult
    Dim result As CResult
    
    Set result = NewResult(MODULE_NAME, "SetSelectableFieldValue", False)
    
    If Not oFieldValue Is Nothing Then
        With Me.RequestRecordset
            If Not IsNull(.fields(oFieldValue.Name).Value) Then
                If oFieldValue.IsSelectable Then
                    oFieldValue.FieldValue.SetValue CVar(.fields(oFieldValue.Name).Value)
                    result.SetSuccess
                Else
                    result.SetFailure "field " & oFieldValue.Name & " Is NOT Selectable"
                End If
            Else
                oFieldValue.FieldValue.ClearValue
                result.SetFailure "field " & oFieldValue.Name & " NOT in Record"
            End If
        End With
    Else
        result.SetFailure "the field is NULL"
    End If
    
    Set SetSelectableFieldValue = result
End Function

Private Function SetSelectableRecord(ByRef oRecordValue As CRecord) As CResult
    Dim result As CResult, loopDone As Boolean, lngTest As Long
    
    ' create new Record, if not done yet
    If oRecordValue Is Nothing Then
        Set oRecordValue = NewRecord(Me.CurrentDataAccess, Me.CurrentDataAccess.Record.User)
    End If
    
    lngTest = Me.CurrentDataAccess.Record.FieldListCount
    lngTest = oRecordValue.FieldListCount
    ' set record ID
    AddSubResult MODULE_NAME, "SetSelectableRecord", result, SetRecordIdValue(oRecordValue)
    lngTest = Me.CurrentDataAccess.Record.FieldListCount
    
    ' Reset Record field list's fields picked
    Me.CurrentDataAccess.Record.FieldList.ResetFieldsPicked
    loopDone = False
    
    Do
        Dim currField As CField
        Set currField = Me.CurrentDataAccess.Record.FieldList.GetNextField().ResetCode()
        AddSubResult MODULE_NAME, "SetSelectableRecord", result, SetSelectableFieldValue(currField)
        
        If result.LastResult.IsOK Then
            oRecordValue.FieldList.AddField currField
        End If
        
        If Me.CurrentDataAccess.Record.FieldList.LastFieldReached() Then
            ' The end of list is reached
            Exit Do
            loopDone = True
        End If
        
    Loop While (Not loopDone)
    
    result.SetSuccess
    Set SetSelectableRecord = result
End Function

Private Function GetValue(ByVal oDataAccess As CDataAccess, ByVal strFieldName As String, ByRef oResult As CResult) As CField
    Dim oFieldValue As CField
    
    Set oFieldValue = oDataAccess.Record.FieldList.GetFieldByName(strFieldName)
    
    ' Execute  the Select Command and Set The Relevant RecordSet
    AddSubResult MODULE_NAME, "GetValue", oResult, ExecSelect(oDataAccess)
    
    If oResult.LastResult.IsOK Then
        If Not (Me.RequestRecordset Is Nothing) Then
            If Not Me.RequestRecordset.EOF Then
                SetSelectableFieldValue oFieldValue
                'Me.RequestRecordset.MoveNext
            End If
        End If
    End If
    
    Set GetValue = oFieldValue
End Function

Private Function GetRecord(ByRef oResult As CResult) As CRecord
    Dim oRecordValue As CRecord, loopDone As Boolean, oField As CField, lngTest As Long
    
    lngTest = Me.CurrentDataAccess.Record.FieldListCount
    
    ' Execute the Select Command and Set The Relevant RecordSet
    AddSubResult MODULE_NAME, "GetRecord", oResult, ExecSelect
    
    lngTest = Me.CurrentDataAccess.Record.FieldListCount
    
    ' Create the OUTPUT Record
    Set oRecordValue = NewRecord(Me.CurrentDataAccess, Me.CurrentDataAccess.User)
    
    lngTest = Me.CurrentDataAccess.Record.FieldListCount
    
    If oResult.LastResult.IsOK Then
        If Not (Me.RequestRecordset Is Nothing) Then
            
            If Not Me.RequestRecordset.EOF Then
                ' Loop through all the fields of the first recordset s record
                'AddToDoLastSubFromParent 1, selectProgr
                
                ' Set record values
                AddSubResult MODULE_NAME, "GetRecord", oResult, SetSelectableRecord(oRecordValue)
            End If
        End If
    End If
    
    Set GetRecord = oRecordValue
End Function

Private Function GetRecordList(ByVal oDataAccess As CDataAccess, ByRef oResult As CResult) As CRecordList
    Dim oRecordListValue As CRecordList, loopDone As Boolean
    
    ' Execute the Select Command and Set The Relevant RecordSet
    AddSubResult MODULE_NAME, "GetRecordList", oResult, ExecSelect(oDataAccess)
    
    ' Create the OUTPUT RecordList
    Set oRecordListValue = NewRecordList(oDataAccess, oDataAccess.User)
    
    If oResult.LastResult.IsOK Then
        If Not (Me.RequestRecordset Is Nothing) Then
            
            If Not Me.RequestRecordset.EOF Then
                
                Do
                    Dim currRecord As CRecord
                    
                    ' Set record values
                    AddSubResult MODULE_NAME, "GetRecordList", oResult, SetSelectableRecord(oDataAccess, currRecord)
                    
                    If oResult.LastResult.IsOK Then
                        oRecordListValue.AddRecord currRecord
                    End If
                    
                    ' Move to Next Record
                    Me.RequestRecordset.MoveNext
                Loop Until Me.RequestRecordset.EOF
                
            End If
        End If
    End If
    
    Set GetRecordList = oRecordListValue
End Function

Private Function ExecSelect() As CResult
    Dim result As CResult, strRequestString As String, lngCriteriaCount As Long
    
    'StartNewSubProgressionFromParent "Exécution de la Requête SELECT", 1, selectProgr
    
    'Call PrepareDatabase
    
    'Call RsltListInit
    
    'BuildSelectCmdText oCriterionFields, lIdCriteria, lLimitLines
    
    'Set result = NewResult("ExecSelect")
    ' Build Select Command Text
    AddSubResult MODULE_NAME, "ExecSelect", result, Me.InnerDatabase.BuildSelectStatement(Me.CurrentDataAccess, strRequestString, lngCriteriaCount)
    
    If result.LastResult.IsOK Then
        ' Open Connection
        Me.Connection.OpenConnection
        
        If Me.Connection.IsOpened Then
            
            'BuildRequestCommand True, oCriterionFields, Nothing, lIdCriteria
            ' Build Request Command
            Set result = Me.Command.BuildRequestCommand(Me.CurrentDataAccess, strRequestString)
            
            Set Me.RequestRecordset = Me.Command.ExecuteCommand(Me.AffectedRecordId)
            
            'Set ExecSelect = result
        Else
            result.IsOK = False
            MsgBox "Erreur Ouverture de la Base de Donnees"
            'Set ExecSelect = result
        End If
    Else
        result.SetFailure result.LastResult.Message
    End If
    
    'AddDoneLastSubFromParent 1, True, selectProgr
    
    'MsgBox Me.RequestStr
    Set ExecSelect = result
End Function

Public Function StoreRecord(ByVal oDataAccess As CDataAccess) As CResult
End Function

Public Function UpdateRecord(ByVal oDataAccess As CDataAccess) As CResult
End Function

Public Function DeleteRecord(ByVal oDataAccess As CDataAccess) As CResult
End Function



'***************************************************************************************************************************************************************
'   Interfaces Implementation
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : IDataSource_User
' Purpose       : The Record Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set IDataSource_User(oUser As CModelUser)
    Set Me.User = oUser
End Property

Public Property Get IDataSource_User() As CModelUser
    Set IDataSource_User = Me.User
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : ServerOrPath
' Purpose       : The Record Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IDataSource_ServerOrPath(strServerOrPath As String)
    Me.ServerOrPath = strServerOrPath
End Property

Public Property Get IDataSource_ServerOrPath() As String
    IDataSource_ServerOrPath = Me.ServerOrPath
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : DatabaseOrFileName
' Purpose       : The Record Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IDataSource_DatabaseOrFileName(strDatabaseOrFileName As String)
    Me.DatabaseOrFileName = strDatabaseOrFileName
End Property

Public Property Get IDataSource_DatabaseOrFileName() As String
    IDataSource_DatabaseOrFileName = Me.DatabaseOrFileName
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : IntegratedSecurity
' Purpose       : The Record Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IDataSource_IntegratedSecurity(blnIntegratedSecurity As Boolean)
    Me.IntegratedSecurity = blnIntegratedSecurity
End Property

Public Property Get IDataSource_IntegratedSecurity() As Boolean
    IDataSource_IntegratedSecurity = Me.IntegratedSecurity
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function      : IDataSource_GetValue
'   Purpose       : Get a value from database (select scalar)
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IDataSource_GetValue(ByVal oDataAccess As CDataAccess, ByVal strFieldName As String, ByRef oResult As CResult) As CField
    Set IDataSource_GetValue = GetValue(oDataAccess, strFieldName, oResult)
End Function

Public Function IDataSource_GetRecord(ByVal oDataAccess As CDataAccess, ByRef oResult As CResult) As CRecord
    Set Me.CurrentDataAccess = oDataAccess
    Set IDataSource_GetRecord = GetRecord(oResult)
End Function

Public Function IDataSource_GetRecordList(ByVal oDataAccess As CDataAccess, ByRef oResult As CResult) As CRecordList
    Set IDataSource_GetRecordList = GetRecordList(oDataAccess, oResult)
End Function



Public Function IDataSource_StoreRecord(ByVal oDataAccess As CDataAccess) As CResult
    IDataSource_StoreRecord = StoreRecord(oDataAccess)
End Function

Public Function IDataSource_UpdateRecord(ByVal oDataAccess As CDataAccess) As CResult
    IDataSource_UpdateRecord = UpdateRecord(oDataAccess)
End Function

Public Function IDataSource_DeleteRecord(ByVal oDataAccess As CDataAccess) As CResult
    IDataSource_DeleteRecord = DeleteRecord(oDataAccess)
End Function



'***************************************************************************************************************************************************************
'    Private events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    Set m_connection = NewDatabaseConnection(Me)
    Set m_command = NewDatabaseCommand(Me)
    Set m_request_recordset = New ADODB.Recordset
    m_affected_record_id = 0
End Sub

Private Sub Class_Terminate()
    Me.CloseDatabase
End Sub