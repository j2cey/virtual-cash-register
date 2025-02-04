'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : mFactoryData
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/28
' Purpose   : Manage all factories for Data related Classes instantiation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


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
Public Function NewDatabaseConnection(oUpperDatabase As CDatabaseDS, Optional lngConnectionTimeout As Long = -1) As CDatabaseConnection
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
Public Function NewDatabaseCommand(oUpperDatabase As CDatabaseDS) As CDatabaseCommand
    With New CDatabaseCommand
        .Init oUpperDatabase
        Set NewDatabaseCommand = .Self 'returns the newly created instance
    End With
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
'   Returns     : CDatabaseDS
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewDatabase(ByVal oUser As CUser, ByVal strServerOrPath As String, ByVal strDatabaseOrFileName As String, ByVal oInnerDatabase As IDatabaseDS, Optional ByVal blnIntegratedSecurity As Boolean = True) As CDatabaseDS
    With New CDatabaseDS
        .Init oUser, strServerOrPath, strDatabaseOrFileName, oInnerDatabase, blnIntegratedSecurity
        Set NewDatabase = .Self 'returns the newly created instance
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
'   Returns     : CAccess2007
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewAccess2007(Optional oUser As CUser = Nothing, Optional oRecordDA As CRecordDA = Nothing) As CAccess2007
    With New CAccess2007
        .Init oUser, oRecordDA
        Set NewAccess2007 = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewRecordable
'   Purpose     : Create and Initialize a New Recordable Data Access
'   Arguments   : oUser                     The user performing
'                 oDataSource               The Data Source
'                 strRecordTable            The Record Table
'                 lngRecordId               The Record ID (if any)
'
'   Returns     : CRecordableDA
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewRecordable(ByVal oUser As CUser, ByVal oDataSource As IDataSource, ByVal strRecordTable As String, Optional ByVal oRecord As CRecord = Nothing) As CRecordableDA
    With New CRecordableDA
        .Init oUser, oDataSource, strRecordTable, oRecord
        Set NewRecordable = .Self 'returns the newly created instance
    End With
End Function
