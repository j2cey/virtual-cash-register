Option Explicit

Public encrypter As clsEncrypt
Public userToEdit As clsUser
Public roleToEdit As clsUserRole

Public SettPhoneNumLength As clsSetting
Public SettMotifSpecialChars As clsSetting
Public SettOpenCode As clsSetting

Public SettPymntStatusAttenteValidationId As clsSetting
Public SettPymntStatusAttenteValidationLabel As clsSetting
Public SettPymntStatusAttenteExtractionId As clsSetting
Public SettPymntStatusAttenteExtractionLabel As clsSetting
Public SettPymntStatusExtraitId As clsSetting
Public SettPymntStatusExtraitLabel As clsSetting
Public SettPymntStatusFichierEnvoyeId As clsSetting
Public SettPymntStatusFichierEnvoyeLabel As clsSetting
Public SettPymntStatusEffectueId As clsSetting
Public SettPymntStatusEffectueLabel As clsSetting

Public SettPymntNotifyAttenteValidation As clsSetting
Public SettPymntNotifyAttenteExtraction As clsSetting
Public SettPymntNotifyExtraction As clsSetting
Public SettPymntNotifyFichierEnvoye As clsSetting
Public SettPymntNotifyEffectue As clsSetting

Public SettPymntExtractionDontSendFile As clsSetting
Public SettPymntExtractionSendFileToList As clsSetting
Public SettPymntExtractionSendFileToUsers As clsSetting
Public SettPymntExtractionSendFileToAll As clsSetting
Public SettPymntExtractionFileFolder As clsSetting
Public SettPymntExtractionReceiversList As clsSetting

Public SettPymntLastLinesSizeDashboard As clsSetting
Public SettPymntEmployeesSearchLinesSize As clsSetting

Public Permissions As clsPermissions

Public MailObject As clsMail

Public Sub InitGlobal()
    'InitDB
    
    If Permissions Is Nothing Then
        Set Permissions = New clsPermissions
    End If
    
    If encrypter Is Nothing Then
        Set encrypter = New clsEncrypt
    End If
    
    Set SettPhoneNumLength = NewSetting("PhoneNumLength", "Phone Num Length", "APP-Settings", 8, 1, intval, CVar(8), False)
    Set SettMotifSpecialChars = NewSetting("MotifSpecialChars", "Motif Special Chars", "APP-Settings", 9, 1, textval, "é,e;è,e;ê,e;ë,e;ç,c;à,a;â,a", True)
    Set SettOpenCode = NewSetting("OpenCode", "Open Code", "APP-Settings", 7, 1, boolval, 0, False)
    
    Set SettPymntStatusAttenteValidationId = NewSetting("PymntStatusAttenteValidationId", "Attente Validation Id", "APP-Settings", 10, 1, textval, "1", True)
    Set SettPymntStatusAttenteValidationLabel = NewSetting("PymntStatusAttenteValidationLabel", "Attente Validation Label", "APP-Settings", 11, 1, textval, "Attente Validation", True)
    Set SettPymntStatusAttenteExtractionId = NewSetting("PymntStatusAttenteExtractionId", "Attente Extraction Id", "APP-Settings", 12, 1, textval, "2", True)
    Set SettPymntStatusAttenteExtractionLabel = NewSetting("PymntStatusAttenteExtractionLabel", "Attente Extraction Label", "APP-Settings", 13, 1, textval, "Attente Extraction", True)
    Set SettPymntStatusExtraitId = NewSetting("PymntStatusExtraitId", "Extrait Id", "APP-Settings", 14, 1, textval, "3", True)
    Set SettPymntStatusExtraitLabel = NewSetting("PymntStatusExtraitLabel", "Extrait Label", "APP-Settings", 15, 1, textval, "Extrait", True)
    Set SettPymntStatusFichierEnvoyeId = NewSetting("PymntStatusFichierEnvoyeId", "Attente Envoie ID", "APP-Settings", 41, 1, textval, "4", True)
    Set SettPymntStatusFichierEnvoyeLabel = NewSetting("PymntStatusFichierEnvoyeLabel", "Attente Envoie Label", "APP-Settings", 42, 1, textval, "Effectué", True)
    Set SettPymntStatusEffectueId = NewSetting("PymntStatusEffectueId", "Effectué ID", "APP-Settings", 16, 1, textval, "4", True)
    Set SettPymntStatusEffectueLabel = NewSetting("PymntStatusEffectueLabel", "Effectué Label", "APP-Settings", 17, 1, textval, "Effectué", True)
    Set SettPymntLastLinesSizeDashboard = NewSetting("PymntLastLinesSizeDashboard", "Nombre MAX de Lignes pour le Dashboard", "APP-Settings", 25, 1, intval, 50, True)
    Set SettPymntEmployeesSearchLinesSize = NewSetting("PymntEmployeesSearchLinesSize", "Nombre MAX de Lignes pour la Recherche Employes pour Paiement", "APP-Settings", 26, 1, intval, 50, True)
    
    Set SettPymntNotifyAttenteValidation = NewSetting("PymntNotifyAttenteValidation", "Notifier Paiement Attente Validation", "APP-Settings", 29, 1, boolval, False, True)
    Set SettPymntNotifyAttenteExtraction = NewSetting("PymntNotifyAttenteExtraction", "Notifier PaiementAttente Extraction", "APP-Settings", 30, 1, boolval, False, True)
    Set SettPymntNotifyExtraction = NewSetting("PymntNotifyExtrait", "Notifier Paiement Extrait", "APP-Settings", 31, 1, boolval, False, True)
    Set SettPymntNotifyFichierEnvoye = NewSetting("PymntNotifyFichierEnvoye", "Notifier Paiement Attente Envoie", "APP-Settings", 43, 1, boolval, False, True)
    Set SettPymntNotifyEffectue = NewSetting("PymntNotifyEffectue", "Notifier Paiement Effectué", "APP-Settings", 32, 1, boolval, False, True)
    
    Set SettPymntExtractionDontSendFile = NewSetting("PymntExtractionDontSendFile", "Envoie Fichier Paiement - Ne Pas Envoyer de Mail", "APP-Settings", 33, 1, boolval, False, True)
    Set SettPymntExtractionSendFileToList = NewSetting("PymntExtractionSendFileToList", "Envoie Fichier Paiement - Envoyer à la Liste Uniquement", "APP-Settings", 34, 1, boolval, False, True)
    Set SettPymntExtractionSendFileToUsers = NewSetting("PymntExtractionSendFileToUsers", "Envoie Fichier Paiement - Envoyer aux Utilisateurs Habilités", "APP-Settings", 35, 1, boolval, False, True)
    Set SettPymntExtractionSendFileToAll = NewSetting("PymntExtractionSendFileToAll", "Envoie Fichier Paiement - Envoyer à la Liste + Utilisateurs Habilités", "APP-Settings", 36, 1, boolval, False, True)
    Set SettPymntExtractionFileFolder = NewSetting("PymntExtractionFileFolder", "Répertoire Enregistrement Fichier Paiement", "APP-Settings", 40, 1, textval, "", True)
    Set SettPymntExtractionReceiversList = NewSetting("PymntExtractionReceiversList", "Envoie Fichier Paiement - Liste Destinataires", "APP-Settings", 37, 1, textval, "", True)
    SettPymntExtractionReceiversList.ArraySep = "|"
    
    InitPermissions
    
    InitMail
End Sub

'Public Function GetOpenCode() As Boolean
'    GetOpenCode = bOpenCode
'End Function

'Public Function SetOpenCode(newOpenCode As Boolean) As Boolean
'    bOpenCode = newOpenCode
'End Function

Sub CloseAndSaveOpenWorkbooksNew()
    Application.DisplayAlerts = False
    
    If CBool(SettOpenCode.Val) Then
        Application.Visible = True
    Else
        ActiveWorkbook.Close True
    End If
    
    Application.DisplayAlerts = True
End Sub

Sub CloseAndSaveOpenWorkbooks()
    Dim Wkb As Workbook
    
    Application.DisplayAlerts = False
    
    If CBool(SettOpenCode.Val) Then
        Application.Visible = True
    Else
        With Application
            .ScreenUpdating = False
            ' Loop through the workbooks collection
            For Each Wkb In Workbooks
                With Wkb
                    ' if the book is read-only
                    ' don't save but close
                    If Not Wkb.ReadOnly Then
                        .Save
                    End If
                    ' We save this workbook, but we don't close it
                    ' because we will quit Excel at the end,
                    ' Closing here leaves the app running, but no books
                    If .Name <> ThisWorkbook.Name Then
                        .Close
                    End If
                End With
            Next Wkb
            .ScreenUpdating = True
            .Quit 'Quit Excel
        End With
        
        Application.DisplayAlerts = True
    End If
End Sub

Public Function ValidEmail(ByVal strEmailAddress As String) As Boolean
    On Error GoTo Catch
    
    Dim objRegExp As VBScript_RegExp_55.RegExp
    Dim blnIsValidEmail As Boolean
    
    Set objRegExp = New VBScript_RegExp_55.RegExp
    
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    
    blnIsValidEmail = objRegExp.test(strEmailAddress)
    ValidEmail = blnIsValidEmail
    
    Exit Function
    
Catch:
    ValidEmail = False
End Function

Public Function ValidEmail_Old(eMail As String) As Boolean
    Dim MyRegExp As RegExp
    Dim myMatches As MatchCollection
    
    Set MyRegExp = New RegExp
    MyRegExp.Pattern = "^[a-z0-9_.-]+@[a-z0-9.-]{2,}\.[a-z]{2,4}$"
    MyRegExp.IgnoreCase = True
    MyRegExp.Global = False
    Set myMatches = MyRegExp.Execute(eMail)
    
    ValidEmail = (myMatches.Count = 1)
    
    Set myMatches = Nothing
    Set MyRegExp = Nothing
End Function


Public Function GetUsersByPermissions(arrPermissions As Variant, ByRef recordCount As Long)
    Dim rolepermissionsData As Variant, sPermissions As String, sqlRst As String, i As Long
    Dim usersData As Variant, roleList As String
    
    For i = 0 To UBound(arrPermissions)
        sPermissions = sPermissions & ",'" & CStr(arrPermissions(i)) & "'"
    Next i
    sPermissions = Right(sPermissions, Len(sPermissions) - 1)
    
    sqlRst = "SELECT DISTINCT userrole_id FROM rolepermissions WHERE role_permission IN (" & sPermissions & ")"
    'MsgBox sqlRst, vbInformation, GetAppName
    
    Call PrepareDatabase
    rolepermissionsData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
        
    If recordCount = 0 Then
        Exit Function
    End If
    
    For i = 0 To recordCount - 1
        roleList = roleList & "," & CStr(rolepermissionsData(0, i)) & ""
    Next i
    roleList = Right(roleList, Len(roleList) - 1)
    
    sqlRst = "SELECT DISTINCT Id, userlogin, username, usermail FROM users_view WHERE role_id IN (" & roleList & ")"
    'MsgBox sqlRst, vbInformation, GetAppName
    Call PrepareDatabase
    usersData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    If recordCount = 0 Then
        Exit Function
    End If
    
    GetUsersByPermissions = usersData
End Function

Public Function GetUserMailsByPermissions(arrPermissions As Variant)
    Dim usersArr As Variant, userMailsArr As Variant, recordCount As Long, i As Long
    
    usersArr = GetUsersByPermissions(arrPermissions, recordCount)
    
    If IsEmpty(usersArr) Then
        Exit Function
    End If

    For i = 0 To recordCount - 1
        'MsgBox CStr(usersArr(0, i)) & ", " & CStr(usersArr(1, i)) & ", " & CStr(usersArr(2, i)), vbInformation, GetAppName
        userMailsArr = AddToArray(userMailsArr, Array(usersArr(2, i), usersArr(3, i)))
    Next i

    GetUserMailsByPermissions = userMailsArr
End Function

Public Function FormatPhone(ByRef sPhone) As Boolean
    
    If Not IsNumeric(sPhone) Then
        MsgBox "Erreur ! Numéro Téléphone " & sPhone & " Non Numérique !", vbCritical, GetAppName
        FormatPhone = False
        Exit Function
    End If
    
    If Len(sPhone) < CInt(SettPhoneNumLength.Val) Then
        MsgBox "Erreur ! Numéro Téléphone " & sPhone & " trop court", vbCritical, GetAppName
        FormatPhone = False
        Exit Function
    End If
    
    sPhone = Right(sPhone, CInt(SettPhoneNumLength.Val))
    sPhone = "241" & sPhone
    
    FormatPhone = True
End Function


