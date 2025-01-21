'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CDatabaseCommand
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/21
' Purpose   : Database Command Object
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_module_name As String

Private m_request_command As ADODB.Command
Private m_request_recordset As New ADODB.Recordset
Private m_request_statement As String

Private m_upper_database As CDatabaseDS

Private m_last_result As CResult



'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CDatabaseCommand
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : UpperDatabase
'   Purpose       : The Upper Database
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set UpperDatabase(oUpperDatabase As CDatabaseDS)
    Set m_upper_database = oUpperDatabase
End Property

Public Property Get UpperDatabase() As CDatabaseDS
    Set UpperDatabase = m_upper_database
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
' Property      : RequestStatement
' Purpose       : The Request Statement
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/21    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let RequestStatement(strRequestStatement As String)
    m_request_statement = strRequestStatement
End Property

Public Property Get RequestStatement() As String
    RequestStatement = m_request_statement
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
'   Property      : ModuleName
'   Purpose       : The Name of this Module
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ModuleName(strModuleName As String)
    Let m_module_name = strModuleName
End Property

Public Property Get ModuleName() As String
    ModuleName = m_module_name
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : LastResult
'   Purpose     : Get / Set The Last Result
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LastResult() As CResult
    Set LastResult = m_last_result
End Property

Public Property Set LastResult(oLastResult As CResult)
    Set m_last_result = oLastResult
End Property



'***************************************************************************************************************************************************************
'   Public Functions and Subs
'***************************************************************************************************************************************************************

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
    Me.ConnectionString = UpperDatabase.BuildConnectionString()
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : OpenConnection
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
Public Function OpenConnection() As Boolean
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
            
            Me.Connection.ConnectionString = Me.ConnectionString
            Me.Connection.ConnectionTimeout = Me.ConnectionTimeout
            Me.Connection.Open
        End If
        
ODResume:
    OpenConnection = blnReturn
    Exit Function
ODError:
    blnReturn = False
    UpperDatabase.ShowErrorMessages Err, Me.ModuleName, "OpenConnection"
    Resume ODResume
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : CloseConnection
'   Purpose     : Close connection.
'   Arguments   :
'
'   Returns     : Boolean
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/20  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CloseConnection() As Boolean
    
    On Error Resume Next
        If Not Me.Connection Is Nothing Then
            If Me.Connection.State And adStateOpen Then
                'Connection is defined and still open; close it.
                Me.Connection.Close
            End If
        End If
        CloseConnection = True
    On Error GoTo 0
End Function



'***************************************************************************************************************************************************************
'   Public Functions and Subs
'***************************************************************************************************************************************************************

Private Function ExecSelect(Optional oCriterionFields As Collection = Nothing, Optional lIdCriteria As Long = -1, Optional lLimitLines As Long = 0) As Long
    Dim connectionOpened As Boolean, recAffected As Long
    
    'StartNewSubProgressionFromParent "Exécution de la Requête SELECT", 1, selectProgr
    
    'Call PrepareDatabase
    
    'Call RsltListInit
    
    'BuildSelectCmdText oCriterionFields, lIdCriteria, lLimitLines
    
    'Build the Select Statement
    Me.RequestStatement = Me.UpperDatabase.dd
    
    ' Open The Database
    connectionOpened = Me.UpperDatabase.OpenDatabase()
    
    If connectionOpened Then
        
        BuildRequestCommand True, oCriterionFields, Nothing, lIdCriteria
        
        Set Me.reqstRecordset = Me.RequestCommand.Execute(recAffected)
        
        ExecSelect = 1
    Else
        MsgBox "Erreur Ouverture de la Base de Donnees"
        ExecSelect = -1
    End If
    
    AddDoneLastSubFromParent 1, True, selectProgr
    
    'MsgBox Me.RequestStr
End Function



'***************************************************************************************************************************************************************
'    Private events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    m_connection_timeout = 30
    m_module_name = "CDatabaseConnection"
    Set Me.Connection = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
        Me.Connection.Close
        Set Me.Connection = Nothing
    On Error GoTo 0
End Sub
