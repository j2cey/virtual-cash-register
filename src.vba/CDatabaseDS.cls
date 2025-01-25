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
Private m_database_or_file_name As String

Private m_connection As CDatabaseConnection
Private m_command As CDatabaseCommand

Private m_integrated_security As Boolean
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
    BuildConnectionString = InnerDatabase.BuildConnectionString(Me)
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
    CloseDatabase = Me.Connection.CloseConnection()
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
    m_module_name = "CDatabaseDS"
End Sub

Private Sub Class_Terminate()
End Sub