'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CDatabaseConnection
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/21
' Purpose   : Database Connection Object
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_module_name As String
Private m_connection As ADODB.Connection
Private m_connection_string As String

Private m_connection_timeout As Long
Private m_upper_database As CDatabaseDS

Private m_is_opened As Boolean
Private m_last_result As CResult

Private Const DEFAULT_TIMEOUT As Long = 30


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CDatabaseConnection
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
' Property      : Connection
' Purpose       : The Connection Object
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/21    Jude Parfait        Created
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
' 2025/01/21    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ConnectionString(strConnectionString As String)
    m_connection_string = strConnectionString
End Property

Public Property Get ConnectionString() As String
    ConnectionString = m_connection_string
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : ConnectionTimeout
'   Purpose       : The Database Connection Timeout amount
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21      Jude Parfait        Created
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IsOpened
'   Purpose     : Determine whether this database connection is opened or not
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsOpened(blnIsOpened As Boolean)
    m_is_opened = blnIsOpened
End Property

Public Property Get IsOpened() As Boolean
    IsOpened = m_is_opened
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
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(oUpperDatabase As CDatabaseDS, Optional lngConnectionTimeout As Long = -1)
    
    Set Me.UpperDatabase = oUpperDatabase
    
    If Not lngConnectionTimeout = -1 Then
        Me.ConnectionTimeout = lngConnectionTimeout
    End If
    
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
    Dim blnNewConnect As Boolean, strDataSource As String
    Dim strDB As String, strConn As String
    
    'blnReturn = True
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
            Me.IsOpened = True
        End If
        
ODResume:
    OpenConnection = Me.IsOpened
    Exit Function
ODError:
    Me.IsOpened = False
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
'    Private events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    m_connection_timeout = DEFAULT_TIMEOUT
    m_module_name = "CDatabaseConnection"
    Set Me.Connection = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
        Me.Connection.Close
        Set Me.Connection = Nothing
    On Error GoTo 0
End Sub