Option Explicit

'Global data access class object
Public gobjDB As clsSQLConnection
Public gobAccessDB As clsDB

Public SettDatabaseType As clsSetting
Public SettDbServerName As clsSetting
Public SettDbName As clsSetting
Public SettDbUserName As clsSetting
Public SettDbUserPassword As clsSetting

Public SettDbSelectLimitStatement As clsSetting
Public SettDbSelectLimitSize As clsSetting
Public SettDbSelectLimitPosition As clsSetting


Public gblDbLocalFolder As String
Public gblIntegratedSecurity As Boolean

Public gblUseQuotas As Boolean
Public gblLibertisQuotas As Integer
Public gblMoovQuotas As Integer

Public Sub InitDBXXX()
    gblDatabaseType = "Access2007" ' "SQLServer"
    gblDbServerName = "192.168.5.113"
    gblDbName = GetDbName
    gblDbLocalFolder = GetDbFolder
    gblIntegratedSecurity = False
    
    gblDbUserID = "sa"
    gblDbPassword = "Synivers@2017"
End Sub

Public Function GetDbType() As String
    If SettDatabaseType Is Nothing Then
        Set SettDatabaseType = NewSetting("DatabaseType", "Database Type", "APP-Settings", 2, 1, textval, "Access2007", False)
    End If
    
    GetDbType = CStr(SettDatabaseType.Val)
End Function

Public Function GetDbServerName() As String
    If SettDbServerName Is Nothing Then
        Set SettDbServerName = NewSetting("DbServerName", "DB Server Name", "APP-Settings", 3, 1, textval, GetDefaultDbServerName, False)
    End If
    If CStr(SettDbServerName.Val) = "" Then
        SettDbServerName.SaveValue GetDefaultDbServerName
    End If
    
    GetDbServerName = CStr(SettDbServerName.Val)
End Function

Public Function GetDefaultDbServerName() As String
    GetDefaultDbServerName = AppPath & Application.PathSeparator & GetDbFolder
End Function

Public Function GetDbName() As String
    If SettDbName Is Nothing Then
        Set SettDbName = NewSetting("DbName", "Database Name", "APP-Settings", 4, 1, textval, GetDefaultDbName, False)
    End If
    If CStr(SettDbName.Val) = "" Then
        SettDbName.SaveValue GetDefaultDbName
    End If
    GetDbName = CStr(SettDbName.Val)
End Function

Private Function GetDefaultDbName() As String
    GetDefaultDbName = "comilogcashdb"
End Function

Public Function GetDbUserName() As String
    If SettDbUserName Is Nothing Then
        Set SettDbUserName = NewSetting("DbUserName", "DB User Name", "APP-Settings", 5, 1, textval, "sa", False)
    End If
    GetDbUserName = CStr(SettDbUserName.Val)
End Function

Public Function GetDbUserPassword() As String
    If SettDbUserPassword Is Nothing Then
        Set SettDbUserPassword = NewSetting("DbUserPassword", "DB User Password", "APP-Settings", 6, 1, textval, "Synivers@2017", False)
    End If
    GetDbUserPassword = CStr(SettDbUserPassword.Val)
End Function

Public Function GetDbSelectLimitStatement() As String
    Dim defaultLimitStatement As String

    defaultLimitStatement = "TOP"

    If SettDbSelectLimitStatement Is Nothing Then
        Set SettDbSelectLimitStatement = NewSetting("DbSelectLimitStatement", "DB Select Limit Statement", "APP-Settings", 22, 1, textval, defaultLimitStatement, True)
    End If
    If CStr(SettDbSelectLimitStatement.Val) = "" Then
        SettDbSelectLimitStatement.SaveValue defaultLimitStatement
    End If

    GetDbSelectLimitStatement = CStr(SettDbSelectLimitStatement.Val)
End Function

Public Function GetDbSelectLimitSize() As Long
    Dim defaultLimitSize As Long

    defaultLimitSize = 0

    If SettDbSelectLimitSize Is Nothing Then
        Set SettDbSelectLimitSize = NewSetting("DbSelectLimitSize", "DB Select Limit Default Size", "APP-Settings", 23, 1, intval, defaultLimitSize, True)
    End If
    If CStr(SettDbSelectLimitSize.Val) = "" Then
        SettDbSelectLimitSize.SaveValue defaultLimitSize
    End If

    GetDbSelectLimitSize = CLng(SettDbSelectLimitSize.Val)
End Function

Public Function GetDbSelectLimitPosition() As String
    Dim defaultLimitPosition As String

    defaultLimitPosition = "AVANT"

    If SettDbSelectLimitPosition Is Nothing Then
        Set SettDbSelectLimitPosition = NewSetting("DbSelectLimitPosition", "DB Select Limit Position", "APP-Settings", 24, 1, textval, defaultLimitPosition, True)
    End If
    If CStr(SettDbSelectLimitPosition.Val) = "" Then
        SettDbSelectLimitPosition.SaveValue defaultLimitPosition
    End If

    GetDbSelectLimitPosition = CStr(SettDbSelectLimitPosition.Val)
End Function



Public Function GetDbFolder() As String
    GetDbFolder = "ressources" & Application.PathSeparator & "db"
End Function

Public Function DbPath() As String
    'ActiveWorkbook.Path & Application.PathSeparator &
    DbPath = AppPath & GetDbFolder & Application.PathSeparator & GetDbName & ".accdb"
End Function

Public Sub PrepareDatabase()
    Dim strDatabaseType As String
    Dim strServerName As String
    Dim strDatabaseName As String
    Dim blnIntegratedSecurity As Boolean
    Dim strUserID As String
    Dim strPassword As String
    Dim strDbFolder As String
    Dim targetDB As String
    
    '*** NOTE: The database properties and user ID, etc. can be read from an
    ' INI file or some other source. For this example, just hard-code
    ' the server and database values and assume that Windows integrated
    ' security is being used (so no UID or Pwd are required).
    ' This procedure assumes that you have a local SQL Server
    ' installation, with a database named "MyDatabase".
    ' Modify this as necessary to conform to your test environment.
    
    strDatabaseType = "Access2007" ' "SQLServer"
    strServerName = "192.168.5.113"
    strDatabaseName = "comilogcashdb"
    blnIntegratedSecurity = False
    
    'Call InitDB
    
    Set gobjDB = New clsSQLConnection
    gobjDB.BuildConnectionString GetDbType, GetDbServerName, _
    GetDbName, gblIntegratedSecurity, GetDbUserName, GetDbUserPassword
    
    ' Initialize access DB
    'Set gobAccessDB = Factory.CreateDBobj(AppPath & Application.PathSeparator & strDbFolder & Application.PathSeparator & gblDbName & ".accdb")
    
End Sub

Public Sub DestroyDatabase()
    If Not (gobjDB Is Nothing) Then
        On Error Resume Next
        gobjDB.CloseDB
        Set gobjDB = Nothing
    End If
End Sub

Public Function checkDbExists() As Boolean
    Dim dbFullPath As String, strFolderExists As String, strFileExists As String
    Dim newValue As String
    
    If GetDbType = "Access2007" Then
        
        strFolderExists = Dir(GetDbServerName, vbDirectory)
        If strFolderExists = "" Then
            MsgBox "Le dossier " & GetDbServerName & " n'existe pas. " & vbCrLf & "Changez le dossier svp."
            newValue = GetFolder
            SettDbServerName.SaveValue newValue
        End If
        
        dbFullPath = GetDbServerName & Application.PathSeparator & GetDbName & ".accdb"
        'MsgBox dbFullPath
        strFileExists = Dir(dbFullPath)
        If strFileExists = "" Then
            MsgBox "Fichier de base de données introuvable sur le chemin: " & vbCrLf & dbFullPath & vbCrLf & "Changez le nom du fichier svp."
            newValue = InputBox("Entrez le Nom du fichier de Bas de Données", GetAppName)
            SettDbName.SaveValue newValue
        End If
        
        checkDbExists = True
    End If
    
    checkDbExists = True
End Function

