'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Module    : mInitializeSettingsSystem
'   Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
'   Created   : 2025/02/12
'   Purpose   : Manage All System parameters related variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'***************************************************************************************************************************************************************
'   Public Variables
'***************************************************************************************************************************************************************

Public oSystemSettings As CSettingSystem

'***************************************************************************************************************************************************************
'   Public Functions & Subroutines
'***************************************************************************************************************************************************************

Private Sub SetSystemDefaultSetting()
    ' Set DataSource Default attributes
    oSystemSettings.UserLogin = "Caisse Virtuelle - COMILOG"
    oSystemSettings.UserPwd = "OutLook"
    
    oSystemSettings.SaveValues
End Sub

Public Sub SystemSettingsLoad()
    Set oSystemSettings = NewSettingSystem(GetLoggedUser(), GetMainSheetDataSource(), "Settings Data Source")
    
    ' Try Load saved settings
    oSystemSettings.LoadValues
    
    If Not oSystemSettings.IsComplete Then
        SetSystemDefaultSetting
    End If
End Sub

Public Function GetSettingSystem() As CSettingSystem
    If oSystemSettings Is Nothing Then
        SystemSettingsLoad
    End If
    
    Set GetSettingSystem = oSystemSettings
End Function

'***************************************************************************************************************************************************************
'   Private Functions & Subroutines
'***************************************************************************************************************************************************************


