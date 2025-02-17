'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Module    : mInitializeDataSources
'   Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
'   Created   : 2025/02/05
'   Purpose   : Manage All Data source related variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'***************************************************************************************************************************************************************
'   Public Variables
'***************************************************************************************************************************************************************

Public oMainDataSource As IDataSource
Public oMainSheetDataSource As CDataSourceSheet
Public oMainAccessDatabase As CDatabaseAccess2007

'***************************************************************************************************************************************************************
'   Public Functions & Subroutines
'***************************************************************************************************************************************************************

Public Function GetMainSheetDataSource(Optional ByVal blnRefreshDatasource As Boolean = False) As IDataSource
    
    If oMainSheetDataSource Is Nothing Or blnRefreshDatasource Then
        Set oMainSheetDataSource = NewDataSource(sheetSource, GetLoggedUser(), AppPath(), "ActiveWorkbook Name")
    End If
    
    Set GetMainSheetDataSource = oMainSheetDataSource
End Function

Public Function GetMainDataSource(Optional ByVal blnRefreshDatasource As Boolean = False) As IDataSource
    
    If oMainDatasourceSettings Is Nothing Then
        MainDatasourceSettingsLoad
    End If
    
    If oMainDataSource Is Nothing Or blnRefreshDatasource Then
        If oMainDatasourceSettings.DataSourceClass = databaseSource Then
            Set oMainDataSource = NewDatabase(oLoggedUser, oMainDatasourceSettings.ServerOrPath, oMainDatasourceSettings.DatabaseOrFileName, oMainDatasourceSettings.DatabaseClass, oMainDatasourceSettings.IntegratedSecurity)
        Else
            Set oMainDataSource = NewDataSource(oMainDatasourceSettings.DataSourceClass, oLoggedUser, oMainDatasourceSettings.ServerOrPath, oMainDatasourceSettings.DatabaseOrFileName, oMainDatasourceSettings.IntegratedSecurity)
        End If
    End If
    
    Set GetMainDataSource = oMainDataSource
End Function



'***************************************************************************************************************************************************************
'   Private Functions & Subroutines
'***************************************************************************************************************************************************************

Private Function GetDatabase(ByVal oUser As CModelUser, ByVal strServerOrPath As String, ByVal strDatabaseOrFileName As String, ByVal oInnerDatabase As IDataSourceDatabase, Optional ByVal blnIntegratedSecurity As Boolean = True) As CDataSourceDatabase
    Dim oMainDatabase As CDataSourceDatabase
    
    
End Function