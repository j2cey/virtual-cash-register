'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CDatabaseDS
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/18
' Purpose   : Database Data Source
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_user As CUser
Private m_server_or_path As String
Private m_db_or_file_name As String

Private m_connection As ADODB.Connection
Private m_connection_string As String
Private m_request_command As ADODB.Command

Private m_integrated_security As Boolean
Private m_connection_timeout As Long
Private m_module_name As String

Private m_inner_database As IDatabaseDS



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
'   Property      : DbOrFileName
'   Purpose       : The Database Or File Name
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let DbOrFileName(strDbOrFileName As String)
    Let m_db_or_file_name = strDbOrFileName
End Property

Public Property Get DbOrFileName() As String
    DbOrFileName = m_db_or_file_name
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : Connection
' Purpose       : The Connection Object
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/19    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set Connection(oConnection As ADODB.Connection)
    Set m_connection = oConnection
End Property

Public Property Get Connection() As ADODB.Connection
    Set Connection = m_connection
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : ConnectionString
' Purpose       : The Connection String
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/19    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ConnectionString(strConnectionString As String)
    m_connection_string = strConnectionString
End Property

Public Property Get ConnectionString() As String
    ConnectionString = m_connection_string
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : RequestCommand
' Purpose       : The Request Command
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/21    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set RequestCommand(oRequestCommand As ADODB.Command)
    Set m_request_command = oRequestCommand
End Property

Public Property Get RequestCommand() As ADODB.Command
    Set RequestCommand = m_request_command
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
'   Property      : ConnectionTimeout
'   Purpose       : The Database Connection Timeout amount
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/20      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ConnectionTimeout(lngConnectionTimeout As Long)
    Let m_connection_timeout = lngConnectionTimeout
End Property

Public Property Get ConnectionTimeout() As Long
    ConnectionTimeout = m_connection_timeout
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
'   Subroutine  : BuildConnectionString
'   Purpose     : Construct a full connection string, set local properties, and
'                 return string to calling routine.
'
'   Arguments   :
'
'   Returns     : void
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/19  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildConnectionString()
    Me.ConnectionString = InnerDatabase.BuildConnectionString(Me)
End Sub

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
    Dim blnNewConnect As Boolean, blnReturn As Boolean, strDataSource As String
    Dim strDB As String, strConn As String
    
    blnReturn = True
    blnNewConnect = True
    
    On Error GoTo ODError
        If Not Me.Connection Is Nothing Then
            If Me.Connection.State And adStateOpen Then
                'Connection is alread defined and opened.
                blnNewConnect = False
            End If
        End If
        
        If blnNewConnect Then
            'Must create a new connection
            'Calling routine has not specified Server or Database values;
            ' default to last used.
            If Len(Me.ConnectionString) <= 0 Then
                'If full connection string is blank, use the individual Server/Environment and
                ' Database variables to determine connection.
                Me.BuildConnectionString
            End If
            
            Me.Connection.ConnectionString = Me.BuildConnectionString
            Me.Connection.ConnectionTimeout = Me.ConnectionTimeout
            Me.Connection.Open
        End If
        
ODResume:
    OpenDatabase = blnReturn
    Exit Function
ODError:
    blnReturn = False
    ShowErrorMessages Err, Me.ModuleName, "OpenDatabase"
    Resume ODResume
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
    
    On Error Resume Next
        If Not Me.Connection Is Nothing Then
            If Me.Connection.State And adStateOpen Then
                'Connection is defined and still open; close it.
                Me.Connection.Close
            End If
        End If
        CloseDatabase = True
    On Error GoTo 0
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetRecordset
'   Purpose     : Return a disconnected recordset from the database.
'   Arguments   :
'
'   Returns     : ADODB.Recordset
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/20  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetRecordset(strSQL As String) As ADODB.Recordset
    
    Dim rst As ADODB.Recordset
    On Error GoTo GRError
    
    If Me.OpenDatabase() Then
        Set rst = New ADODB.Recordset
        rst.CursorLocation = adUseClient
        rst.Open strSQL, Me.Connection, adOpenDynamic
        Set GetRecordset = rst
        rst.ActiveConnection = Nothing
    End If
GRResume:
    Me.CloseDatabase
    Exit Function
GRError:
    ShowErrorMessages Err, Me.ModuleName, "GetRecordset"
    Resume GRResume
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetRecordsetToArray
'   Purpose     : Return recordset data into an array. If, for any reason, the recordset
'                 does not return any data, this function simply creates a single element
'                 array and populates it with "No matching records...".
'
'   Arguments   : strSQL        The SQL tatement
'                 recordCount   The number of returned records
'
'   Returns     : Variant
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/20  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetRecordsetToArray(strSQL As String, ByRef recordCount As Long) As Variant
    
    Dim rst As ADODB.Recordset
    Dim arrData As Variant
    Dim lngX As Long
    
    recordCount = 0
    
    On Error GoTo GRAError
        
        If Me.OpenDatabase() Then
            Set rst = New ADODB.Recordset
            With rst
                .CursorLocation = adUseClient
                .LockType = adLockReadOnly
                .Open strSQL, mobjConn, adOpenStatic
                If Not (rst.EOF) Then
                    'Disconnect the recordset
                    .ActiveConnection = Nothing
                    'Get the field count
                    lngX = .fields.count
                    recordCount = CLng(.recordCount)
                    arrData = .GetRows()
                Else
                    'Recordset is empty; create dummy array record
                    ReDim arrData(0, 0)
                    arrData(0, 0) = "No matching records found in the database"
                End If
            End With
        Else
        
            'Connection failed for some reason; in order not to crash the
            ' calling routine, create dummy array record.
            ReDim arrData(0, 0)
            arrData(0, 0) = "Could not open database connection"
        End If
GRAResume:
    GetRecordsetToArray = arrData
    Me.CloseDatabase
    Exit Function
GRAError:
    ShowErrorMessages Err, Me.ModuleName, "GetRecordsetToArray"
    Resume GRAResume
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetMultipleRecordsets
'   Purpose     : Allows for multiple recordsets to be returned to the calling routine.
'                 IMPORTANT - Although this method is similar to the GetRecordSet function,
'                 there are significant differences.
'                 1. Do not disconnect the Recordset (i.e. don't set ActiveConnection to Nothing).
'                 2. You must not close the database connection! It must remain open so that
'                 the calling routine can retrieve all of the resultsets in the returned Recordset.
'                 IT IS UP TO THE CALLING ROUTINE TO CLOSE THE CONNECTION WHEN DONE.
'
'   Arguments   : strSQL        The SQL tatement
'
'   Returns     : ADODB.Recordset
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/20  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetMultipleRecordsets(strSQL As String) As ADODB.Recordset
    
    Dim rst As ADODB.Recordset
    
    On Error GoTo GMRError
        If Me.OpenDatabase() Then
            Set rst = New ADODB.Recordset
            rst.Open strSQL, mobjConn
            Set GetMultipleRecordsets = rst
        End If
GMRResume:
    'REMEMBER! Do NOT close the connection here. Calling routine must close it.
    Exit Function
GMRError:
    ShowErrorMessages Err, Me.ModuleName, "GetMultipleRecordsets"
    Resume GMRResume
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : ExecuteActionQuery
'   Purpose     : Execute an action query, or stored procedure, which is not expected
'                 to return a resultset.
'
'   Arguments   : strSQLToExecute   The SQL to be executed
'                 varLastIdentify   The returned Identity, if any
'
'   Returns     : Boolean
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/20  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExecuteActionQuery(strSQLToExecute As String, Optional ByRef varLastIdentify As Variant) As Boolean
    
    Dim blnReturn As Boolean
    blnReturn = True
    On Error GoTo EAQError
    If Me.OpenDatabase() Then
        Me.Connection.Execute strSQLToExecute
        If Not IsMissing(varLastIdentify) Then
            varLastIdentify = Me.Connection.Execute("select @@identity").fields(0)
        End If
    End If
EAQResume:
    Me.CloseDatabase
    ExecuteActionQuery = blnReturn
    Exit Function
EAQError:
    blnReturn = False
    ShowErrorMessages Err, Me.ModuleName, "ExecuteActionQuery"
    Resume EAQResume
End Function







Private Function ExecSelect()
    
End Function

Public Function SelectOneField()

End Function

Public Function SelectMultipleFields()

End Function






'***************************************************************************************************************************************************************
'    Private events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    m_connection_timeout = 30
    m_module_name = "CDatabaseDS"
    Set Me.Connection = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
        Me.Connection.Close
        Set Me.Connection = Nothing
    On Error GoTo 0
End Sub