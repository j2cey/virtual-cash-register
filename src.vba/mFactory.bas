'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : mFactory
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/08
' Purpose   : Manage all factories for Classes instantiation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewResult
'   Purpose     : Initialize the Object
'   Arguments   : blnIsOK           Tells if the operation has been well performed
'                 lngCode           The Result Code
'                 strMessage        The Result Message
'
'   Returns     : CResult
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewResult(Optional strLabel As String = "", Optional blnIsOK As Boolean = False, Optional lngCode As Long = 0, Optional strMessage As String = "") As CResult
    With New CResult
        .Init strLabel, blnIsOK, lngCode, strMessage
        Set NewResult = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewRecord
'   Purpose     : Create and Initialize a New Record
'   Arguments   : blnIsOK           Tells if the operation has been well performed
'                 lngCode           The Result Code
'                 strMessage        The Result Message
'
'   Returns     : CResult
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewRecord(Optional oRecordDA As CRecordDA = Nothing, Optional oUser As CUser = Nothing) As CRecord
    With New CRecord
        .Init oRecordDA, oUser
        Set NewRecord = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewRecordable
'   Purpose     : Create and Initialize a New Recordable
'   Arguments   : blnIsOK           Tells if the operation has been well performed
'                 lngCode           The Result Code
'                 strMessage        The Result Message
'
'   Returns     : CResult
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewRecordable(ByVal oUser As CUser, ByVal oDataSource As IDataSource, ByVal strRecordTable As String, Optional ByVal lngRecordId As Long = -1) As CRecordableDA
    With New CRecordableDA
        .Init oUser, oDataSource, strRecordTable, lngRecordId
        Set NewRecordable = .Self 'returns the newly created instance
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
'   Returns     : CDatabaseDS
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewAccess2007(ByVal oUser As CUser, ByVal strServerOrPath As String, ByVal strDatabaseOrFileName As String, ByVal oInnerDatabase As IDatabaseDS, Optional ByVal blnIntegratedSecurity As Boolean = True) As CAccess2007
    With New CAccess2007
        .Init oUser, strServerOrPath, strDatabaseOrFileName, oInnerDatabase, blnIntegratedSecurity
        Set NewAccess2007 = .Self 'returns the newly created instance
    End With
End Function