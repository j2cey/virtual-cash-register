Option Explicit



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewSettingSystem
'   Purpose     : Create and Initialize a New System Setting
'   Arguments   :
'
'   Returns     : CSettingSystem
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewSettingSystem(ByVal oUser As CUser, ByVal oDataSource As IDataSource, ByVal strRecordTable As String) As CSettingSystem
    With New CSettingSystem
        .Init ByVal oUser, oDataSource, strRecordTable
        Set NewSettingSystem = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewSettingDataSource
'   Purpose     : Create and Initialize a New Setting DataSource
'   Arguments   :
'
'   Returns     : CSettingDataSource
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewSettingDataSource(ByVal oUser As CUser, ByVal oDataSource As IDataSource, ByVal strRecordTable As String) As CSettingDataSource
    With New CSettingDataSource
        .Init ByVal oUser, oDataSource, strRecordTable
        Set NewSettingDataSource = .Self 'returns the newly created instance
    End With
End Function