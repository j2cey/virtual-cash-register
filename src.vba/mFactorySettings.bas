Option Explicit



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewSetting
'   Purpose     : Create and Initialize a New Super Setting Object
'   Arguments   :
'
'   Returns     : CSetting
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/14      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewSetting(ByVal oUser As CModelUser, ByVal oDataSource As IDataSource, ByVal strTableForSaving As String, Optional ByVal strTableForSelecting As String = "", Optional ByVal lngStartRow As Long = -1) As CSetting
    With New CSetting
        .Init ByVal oUser, oDataSource, strTableForSaving, strTableForSelecting, lngStartRow
        Set NewSetting = .Self 'returns the newly created instance
    End With
End Function

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
Public Function NewSettingSystem(ByVal oUpperSetting As CSetting) As CSettingSystem
    With New CSettingSystem
        .Init oUpperSetting
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
Public Function NewSettingDataSource(ByVal oUpperSetting As CSetting) As CSettingDataSource
    With New CSettingDataSource
        .Init oUpperSetting
        Set NewSettingDataSource = .Self 'returns the newly created instance
    End With
End Function