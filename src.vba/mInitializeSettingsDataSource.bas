'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Module    : mInitializeSettingsDataSource
'   Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
'   Created   : 2025/02/05
'   Purpose   : Manage All Data source parameters related variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'***************************************************************************************************************************************************************
'   Public Variables
'***************************************************************************************************************************************************************

Public Const DB_FOLDER_NAME As String = "db"
Public Const DB_NAME As String = "comilogcashdb"

'***************************************************************************************************************************************************************
'   Public Functions & Subroutines
'***************************************************************************************************************************************************************

Public Sub SetMainDatasourceDefaultSetting()
    ' Set DataSource Default attributes
    oMainDatasourceSettings.UserLogin = "defaultDSUser"
    oMainDatasourceSettings.UserPwd = "Default Data Source User password"
    oMainDatasourceSettings.DataSourceClass = databaseSource
    oMainDatasourceSettings.DatabaseClass = access2007
    oMainDatasourceSettings.ServerOrPath = DbPath
    oMainDatasourceSettings.DatabaseOrFileName = DB_NAME
    oMainDatasourceSettings.IntegratedSecurity = True
    
    oMainDatasourceSettings.SaveValues
End Sub

Public Sub MainDatasourceSettingsLoad()
    InitSettingsMainDatasource
    
    ' Try Load saved settings
    oMainDatasourceSettings.LoadValues
    
    If Not oMainDatasourceSettings.IsComplete Then
        SetMainDatasourceDefaultSetting
    End If
End Sub

Public Function DbPath() As String
    DbPath = AppPath & Application.PathSeparator & DB_FOLDER_NAME & Application.PathSeparator
End Function

'***************************************************************************************************************************************************************
'   Private Functions & Subroutines
'***************************************************************************************************************************************************************

