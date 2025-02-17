'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Module    : mInitializeSettings
'   Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
'   Created   : 2025/02/14
'   Purpose   : Manage All App Settings related variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'***************************************************************************************************************************************************************
'   Public Variables
'***************************************************************************************************************************************************************

Public oMainDatasourceSettings As CSettingDataSource
Public oSystemSettings As CSettingSystem


'***************************************************************************************************************************************************************
'   Public Functions & Subroutines
'***************************************************************************************************************************************************************

Public Sub InitSettingsMainDatasource()
    Set oMainDatasourceSettings = NewSettingDataSource(NewSetting(GetLoggedUser(), GetMainSheetDataSource(), "Settings-Data-Source", "Settings-Data-Source", 1))
End Sub

Public Sub InitSettingsSystem()
    Set oSystemSettings = NewSettingSystem(NewSetting(GetLoggedUser(), GetMainSheetDataSource(), "Settings-Data-Source", "Settings-Data-Source", 1))
End Sub

'***************************************************************************************************************************************************************
'   Private Functions & Subroutines
'***************************************************************************************************************************************************************


