'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : mFactoryData
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/28
' Purpose   : Manage all factories for Data related Classes instantiation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit



Public Function NewDataSource(ByVal eDataSourceClass As enuDataSourceClass, ByVal oUser As CUser, ByVal strServerOrPath As String, ByVal strDatabaseOrFileName As String, Optional ByVal blnIntegratedSecurity As Boolean = True) As IDataSource
    Dim oDataSource As IDataSource
    
    Select Case eDataSourceClass
        Case databaseSource
          Set oDataSource = New CDataSourceDatabase
        Case fileSource
          Set oDataSource = New CDataSourceFile
        Case sheetSource
          Set oDataSource = New CDataSourceSheet
    End Select
    
    Set oDataSource.User = oUser
    oDataSource.ServerOrPath = strServerOrPath
    oDataSource.DatabaseOrFileName = strDatabaseOrFileName
    oDataSource.IntegratedSecurity = blnIntegratedSecurity
    
    Set NewDataSource = oDataSource
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewDatabase
'   Purpose     : Create and Initialize a New Database Data Source
'   Arguments   : oUser                     The user performing
'                 strServerOrPath           The server IP or File Path
'                 strDatabaseOrFileName     The Database Or File Name
'                 oInnerDatabase            The Inner Database
'                 blnIntegratedSecurity     Determine whether the integrated security must be set or not
'
'   Returns     : CDataSourceDatabase
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewDatabase(ByVal oUser As CUser, ByVal strServerOrPath As String, ByVal strDatabaseOrFileName As String, Optional ByVal eDatabaseClass As enuDatabaseClass = None, Optional ByVal blnIntegratedSecurity As Boolean = True) As CDataSourceDatabase
    Dim oInnerDatabase As IDataSourceDatabase
    
    With New CDataSourceDatabase
        .Init oUser, strServerOrPath, strDatabaseOrFileName, eDatabaseClass, blnIntegratedSecurity
        Set NewDatabase = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewDatabaseConnection
'   Purpose     : Create and Initialize a Database Connection
'   Arguments   : oUpperDatabase            The upper database
'                 lngConnectionTimeout      The timemout
'
'   Returns     : CDatabaseConnection
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewDatabaseConnection(oUpperDatabase As CDataSourceDatabase, Optional lngConnectionTimeout As Long = -1) As CDatabaseConnection
    With New CDatabaseConnection
        .Init oUpperDatabase, lngConnectionTimeout
        Set NewDatabaseConnection = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewDatabaseCommand
'   Purpose     : Create and Initialize a Database Command
'   Arguments   : oUpperDatabase            The upper database
'
'   Returns     : CDatabaseConnection
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewDatabaseCommand(oUpperDatabase As CDataSourceDatabase) As CDatabaseCommand
    With New CDatabaseCommand
        .Init oUpperDatabase
        Set NewDatabaseCommand = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewAccess2007
'   Purpose     : Create and Initialize a New Access 2007 Database
'   Arguments   : oUser                     The user performing
'                 strServerOrPath           The server IP or File Path
'                 strDatabaseOrFileName     The Database Or File Name
'                 oInnerDatabase            The Inner Database
'                 blnIntegratedSecurity     Determine whether the integrated security must be set or not
'
'   Returns     : CDatabaseAccess2007
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewAccess2007(Optional oUser As CUser = Nothing, Optional oDataAccess As CDataAccess = Nothing) As CDatabaseAccess2007
    With New CDatabaseAccess2007
        .Init oUser, oDataAccess
        Set NewAccess2007 = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewDataAccess
'   Purpose     : Create and Initialize a New DataAccess Data Access
'   Arguments   : oUser                     The user performing
'                 oDataSource               The Data Source
'                 strRecordTable            The Record Table
'                 lngRecordId               The Record ID (if any)
'
'   Returns     : CDataAccess
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewDataAccess(ByVal oUser As CUser, ByVal oDataSource As IDataSource, ByVal strRecordTable As String, Optional ByVal oRecord As CRecord = Nothing) As CDataAccess
    With New CDataAccess
        .Init oUser, oDataSource, strRecordTable, oRecord
        Set NewDataAccess = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewDataSourceSheet
'   Purpose     : Create and Initialize a New Sheet Data Source
'   Arguments   : oUser                     The user performing
'
'   Returns     : CDataSourceSheet
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/05      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewDataSourceSheet(Optional oUser As CUser = Nothing) As CDataSourceSheet
    With New CDataSourceSheet
        .Init oUser
        Set NewDataSourceSheet = .Self 'returns the newly created instance
    End With
End Function