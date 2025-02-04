'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CDatabaseDS
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/18
' Purpose   : Database Data Source
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_module_name As String
Private m_user As CUser
Private m_server_or_path As String
Private m_database_or_file_name As String

Private m_connection As CDatabaseConnection
Private m_command As CDatabaseCommand
Private m_request_recordset As New ADODB.Recordset

Private m_integrated_security As Boolean
Private m_affected_record_id As Long

Private m_inner_database As IDatabaseDS


Implements IDataSource

'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CDatabaseDS
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
Public Property Set InnerDatabase(oInnerDatabase As IDatabaseDS)
    Set m_inner_database = oInnerDatabase
End Property

Public Property Get InnerDatabase() As IDatabaseDS
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
Public Property Set User(oUser As CUser)
    Set m_user = oUser
End Property

Public Property Get User() As CUser
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
'   Property      : ModuleName
'   Purpose       : The Name of this Module
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/20      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ModuleName(strModuleName As String)
    Let m_module_name = strModuleName
End Property

Public Property Get ModuleName() As String
    ModuleName = m_module_name
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
Public Sub Init(ByVal oUser As CUser, ByVal strServerOrPath As String, ByVal strDatabaseOrFileName As String, ByVal oInnerDatabase As IDatabaseDS, Optional ByVal blnIntegratedSecurity As Boolean = True)
    
    Set Me.User = oUser
    Me.ServerOrPath = strServerOrPath
    Me.DatabaseOrFileName = strDatabaseOrFileName
    
    Set Me.InnerDatabase = oInnerDatabase
    Set Me.InnerDatabase.UpperDatabase = Me
    
    Me.IntegratedSecurity = blnIntegratedSecurity
End Sub


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

Private Function SetSelectableFieldValue(ByVal oFieldValue As CField) As CResult
    Dim result As CResult
    
    Set result = NewResult("SetSelectableFieldValue", False)
    
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

Private Function SetSelectableRecord(ByVal oRecordTemplate As CRecord, ByRef oRecordValue As CRecord) As CResult
    Dim result As CResult, loopDone As Boolean, oField As CField
    
    ' Reset Record field list's fields picked
    oRecordTemplate.FieldList.ResetFieldsPicked
    loopDone = False
                
    Do
        Set oField = oRecordTemplate.FieldList.GetNextField()
        AddSubResult "SetSelectableRecord", result, SetSelectableFieldValue(oField)
        
        If result.LastResult.IsOK Then
            oRecordValue.FieldList.AddField oField
        End If
        
        If oRecordTemplate.FieldList.LastFieldReached() Then
            ' The end of list is reached
            Exit Do
            loopDone = True
        End If
        
    Loop While (Not loopDone)
    
    result.SetSuccess
    Set SetSelectableRecord = result
End Function

Private Function GetValue(ByVal dataAccess As CRecordableDA, ByVal strFieldName As String, ByRef oResult As CResult) As CField
    Dim oFieldValue As CField
    
    Set oFieldValue = dataAccess.Record.FieldList.GetFieldByName(strFieldName)
    
    ' Execute  the Select Command and Set The Relevant RecordSet
    AddSubResult "GetValue", oResult, ExecSelect(dataAccess)
    
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

Private Function GetRecord(ByVal dataAccess As CRecordableDA, ByRef oResult As CResult) As CRecord
    Dim oRecordValue As CRecord, loopDone As Boolean, oField As CField
    
    ' Execute the Select Command and Set The Relevant RecordSet
    AddSubResult "GetRecord", oResult, ExecSelect(dataAccess)
    
    ' Create the OUTPUT Record
    Set oRecordValue = NewRecord(dataAccess, dataAccess.User)
    
    If oResult.LastResult.IsOK Then
        If Not (Me.RequestRecordset Is Nothing) Then
            
            If Not Me.RequestRecordset.EOF Then
                ' Loop through all the fields of the first recordset s record
                'AddToDoLastSubFromParent 1, selectProgr
                
                ' Set record values
                AddSubResult "GetRecord", oResult, SetSelectableRecord(dataAccess.Record, oRecordValue)
            End If
        End If
    End If
    
    Set GetRecord = oRecordValue
End Function

Private Function ExecSelect(ByVal dataAccess As CRecordableDA) As CResult
    Dim result As CResult, strRequestString As String, lngCriteriaCount As Long
    
    'StartNewSubProgressionFromParent "Exécution de la Requête SELECT", 1, selectProgr
    
    'Call PrepareDatabase
    
    'Call RsltListInit
    
    'BuildSelectCmdText oCriterionFields, lIdCriteria, lLimitLines
    
    'Set result = NewResult("ExecSelect")
    ' Build Select Command Text
    AddSubResult "ExecSelect", result, Me.InnerDatabase.BuildSelectStatement(dataAccess, strRequestString, lngCriteriaCount)
    
    If result.LastResult.IsOK Then
        ' Open Connection
        Me.Connection.OpenConnection
        
        If Me.Connection.IsOpened Then
            
            'BuildRequestCommand True, oCriterionFields, Nothing, lIdCriteria
            ' Build Request Command
            Set result = Me.Command.BuildRequestCommand(dataAccess, strRequestString)
            
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

Public Function InsertValues(ByVal dataAccess As CRecordableDA) As Variant
End Function

Public Function UpdateValues(ByVal dataAccess As CRecordableDA) As Variant
End Function

Public Function DeleteValues(ByVal dataAccess As CRecordableDA) As Variant
End Function



'***************************************************************************************************************************************************************
'   Interfaces Implementation
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function      : IDataSource_GetValue
'   Purpose       : Get a value from database (select scalar)
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IDataSource_GetValue(ByVal dataAccess As CRecordableDA, ByVal strFieldName As String, ByRef oResult As CResult) As CField
    Set IDataSource_GetValue = GetValue(dataAccess, strFieldName, oResult)
End Function

Public Function IDataSource_GetRecord(ByVal dataAccess As CRecordableDA, ByRef oResult As CResult) As CRecord
    Set IDataSource_GetRecord = GetRecord(dataAccess, oResult)
End Function

Public Function IDataSource_GetRecordList(ByVal dataAccess As CRecordableDA, ByRef oRecordListValue As CRecordList) As CResult
End Function



Public Function IDataSource_AddValues(ByVal dataAccess As CRecordableDA) As Variant
    IDataSource_AddValues = InsertValues(dataAccess)
End Function

Public Function IDataSource_UpdateValues(ByVal dataAccess As CRecordableDA) As Variant
    IDataSource_UpdateValues = UpdateValues(dataAccess)
End Function

Public Function IDataSource_DeleteValues(ByVal dataAccess As CRecordableDA) As Variant
    IDataSource_DeleteValues = DeleteValues(dataAccess)
End Function



'***************************************************************************************************************************************************************
'    Private events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    m_module_name = "CDatabaseDS"
    Set m_connection = NewDatabaseConnection(Me)
    Set m_command = NewDatabaseCommand(Me)
    Set m_request_recordset = New ADODB.Recordset
    m_affected_record_id = 0
End Sub

Private Sub Class_Terminate()
    Me.CloseDatabase
End Sub