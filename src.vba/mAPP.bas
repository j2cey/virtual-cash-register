Option Explicit

'Public sAppName As String
Public SettAppName As clsSetting

Public Function GetAppName() As String
    If SettAppName Is Nothing Then
        Set SettAppName = NewSetting("AppName", "Application Name", "APP-Settings", 1, 1, textval, "Comilog-Mobicash", True)
    End If
    
    GetAppName = CStr(SettAppName.Val)
End Function

Public Function SetAppName(newAppName As String) As String
    SettAppName.SaveValue newAppName
End Function

Public Function AppPath() As String
    AppPath = ThisWorkbook.path & ""
End Function

Public Function GetAppLogo() As String
    Dim filePath As String, strFileExists As String
    
    filePath = AppPath & Application.PathSeparator & "ressources" & Application.PathSeparator & "logo" & Application.PathSeparator & "logo.jpg"
    strFileExists = Dir(filePath)
    
    If strFileExists = "" Then
        GetAppLogo = ""
    Else
        GetAppLogo = filePath
    End If
    
End Function

Public Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function