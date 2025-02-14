Option Explicit

Public bSettingsLoaded As Boolean

Private currAppName As String
Private currDbType As String
Private currDbPath As String
Private currDbServerName As String
Private currDbName As String
Private currDbUserName As String
Private currDbUserPassword As String
Private currOpenCode As Boolean
Private currPhoneNumLength As Integer
Private currMotifSpecialChars As String

Public oSettingPaymentFrm As CUIControl
Public lMotifSpecialCharSelectedIndex As Long

Public oSettingPaymentSendFileFrm As CUIControl

Private currPymntStatusAttenteValidationId As String
Private currPymntStatusAttenteValidationLabel As String
Private currPymntStatusAttenteExtractionId As String
Private currPymntStatusAttenteExtractionLabel As String
Private currPymntStatusExtraitId As String
Private currPymntStatusExtraitLabel As String
Private currPymntStatusFichierEnvoyeId As String
Private currPymntStatusFichierEnvoyeLabel As String
Private currPymntStatusEffectueId As String
Private currPymntStatusEffectueLabel As String

Private currPymntNotifyAttenteValidation As Boolean
Private currPymntNotifyAttenteExtraction As Boolean
Private currPymntNotifyExtrait As Boolean
Private currPymntNotifyFichierEnvoye As Boolean
Private currPymntNotifyEffectue As Boolean

Private currPwdLegnthMin As String
Private currPwdUppercaseMin As String
Private currPwdNumberMin As String
Private currPwdSpecialCharsMin As String
Private currPwdValidity As String
Private currPwdUpdCanBeSame As Boolean

Private currDbSelectLimitStatement As String
Private currDbSelectLimitSize As Long
Private currDbSelectLimitPosition As String

Private currPymntLastLinesSizeDashboard As Long
Private currPymntEmployeesSearchLinesSize As Long

Public oSettingMailParametersFrm As CUIControl
Public lMailProviderParameterSelectedIndex As Long
Private currMailProviderSelected As Long
Private currMailProviderParametersSelected As String

Private currPymntExtractionDontSendFile As Boolean
Private currPymntExtractionSendFileToList As Boolean
Private currPymntExtractionSendFileToUsers As Boolean
Private currPymntExtractionSendFileToAll As Boolean
Private currPymntExtractionFileFolder As String
Private currPymntExtractionReceiversList As String

Public lPymntExtractionReceiverSelectedIndex As Long

Public Sub InitSettingsForm(uFrm As MSForms.UserForm)
    Dim oUctl As CUIControl
    
    uFrm.SettingAppNameTBx.Text = GetAppName
    uFrm.SettingDbTypeTBx.Text = GetDbType
    uFrm.SettingDbPathTBx.Text = GetDbServerName
    uFrm.SettingDbNameTBx.Text = GetDbName
    uFrm.SettingDbUserTBx.Text = GetDbUserName
    uFrm.SettingDbPwdTBx.Text = GetDbUserPassword
    uFrm.OpenCodeCBx.value = CBool(SettOpenCode.Val)
    
    uFrm.PhoneLengthTBx.Text = CStr(SettPhoneNumLength.Val)
    InitMofifSpecialChars uFrm
    
    uFrm.PymntStatusAttenteValidationIdTBx.Text = CStr(SettPymntStatusAttenteValidationId.Val)
    uFrm.PymntStatusAttenteValidationLabelTBx.Text = CStr(SettPymntStatusAttenteValidationLabel.Val)
    uFrm.PymntStatusAttenteExtractionIdTBx.Text = CStr(SettPymntStatusAttenteExtractionId.Val)
    uFrm.PymntStatusAttenteExtractionLabelTBx.Text = CStr(SettPymntStatusAttenteExtractionLabel.Val)
    uFrm.PymntStatusExtraitIdTBx.Text = CStr(SettPymntStatusExtraitId.Val)
    uFrm.PymntStatusExtraitLabelTBx.Text = CStr(SettPymntStatusExtraitLabel.Val)
    uFrm.PymntStatusFichierEnvoyeIdTBx.Text = CStr(SettPymntStatusFichierEnvoyeId.Val)
    uFrm.PymntStatusFichierEnvoyeLabelTBx.Text = CStr(SettPymntStatusFichierEnvoyeLabel.Val)
    uFrm.PymntStatusEffectueIdTBx.Text = CStr(SettPymntStatusEffectueId.Val)
    uFrm.PymntStatusEffectueLabelTBx.Text = CStr(SettPymntStatusEffectueLabel.Val)
    
    uFrm.PymntAttenteValidationChBx.value = CBool(SettPymntNotifyAttenteValidation.Val)
    uFrm.PymntAttenteExtractionNotifyChBx.value = CBool(SettPymntNotifyAttenteExtraction.Val)
    uFrm.PymntExtraitNotifyChBx.value = CBool(SettPymntNotifyExtraction.Val)
    uFrm.PymntFichierEnvoyeNotifyChBx.value = CBool(SettPymntNotifyFichierEnvoye.Val)
    uFrm.PymntEffectueNotifyChBx.value = CBool(SettPymntNotifyEffectue.Val)
    
    Set oSettingPaymentFrm = NewUCTL(uFrm.SettingPaymentFrm, oMainUfrm)
    Set oUctl = oSettingPaymentFrm.AddCtl(uFrm.SettingPaymentFrm.MotifSpecialCharCancelImg).SetSizes(NewSize(15, 15), NewSize(17, 17))
    Set oUctl = oSettingPaymentFrm.AddCtl(uFrm.SettingPaymentFrm.MotifSpecialCharSaveImg).SetSizes(NewSize(15, 15), NewSize(17, 17))
    Set oUctl = oSettingPaymentFrm.AddCtl(uFrm.SettingPaymentFrm.MotifSpecialCharDeleteImg).SetSizes(NewSize(15, 15), NewSize(17, 17))
    
    Set oSettingPaymentSendFileFrm = NewUCTL(uFrm.SettingPaymentSendFileFrm, oMainUfrm)
    Set oUctl = oSettingPaymentSendFileFrm.AddCtl(uFrm.SettingPaymentSendFileFrm.PymntExtractionReceiverSelCancelImg).SetSizes(NewSize(15, 15), NewSize(17, 17))
    Set oUctl = oSettingPaymentSendFileFrm.AddCtl(uFrm.SettingPaymentSendFileFrm.PymntExtractionReceiverSelSaveImg).SetSizes(NewSize(15, 15), NewSize(17, 17))
    Set oUctl = oSettingPaymentSendFileFrm.AddCtl(uFrm.SettingPaymentSendFileFrm.PymntExtractionReceiverSelDeleteImg).SetSizes(NewSize(15, 15), NewSize(17, 17))
    Set oUctl = oSettingPaymentSendFileFrm.AddCtl(uFrm.SettingPaymentSendFileFrm.PymntExtractionFileFolderImg).SetSizes(NewSize(15, 15), NewSize(17, 17))

    UnselectMotifSpecialChar uFrm
    ResetMotifSpecialChar uFrm

    uFrm.PwdLegnthMinTBx.Text = CStr(GetPwdLegnthMin)
    uFrm.PwdUppercaseMinTBx.Text = CStr(GetPwdUppercaseMin)
    uFrm.PwdNumberMinTBx.Text = CStr(GetPwdNumberMin)
    uFrm.PwdSpecialCharsMinTBx.Text = CStr(GetPwdSpecialCharsMin)
    uFrm.PwdValidityTBx.Text = CStr(GetPwdValidity)
    uFrm.PwdUpdCanBeSameChBx.value = CBool(GetPwdUpdCanBeSame)

    InitDbSelectLimitPosition uFrm

    uFrm.SettingDbSelectLimitStatementTBx.Text = CStr(GetDbSelectLimitStatement)
    uFrm.SettingDbSelectLimitSizeTBx.Text = CStr(GetDbSelectLimitSize)
    uFrm.SettingDbSelectLimitPositionCBx.Text = CStr(GetDbSelectLimitPosition)

    uFrm.PymntLastLinesSizeDashboardTBx.Text = CStr(SettPymntLastLinesSizeDashboard.Val)
    uFrm.PymntEmployeesSearchLinesSizeTBx.Text = CStr(SettPymntEmployeesSearchLinesSize.Val)
    
    Set oSettingMailParametersFrm = NewUCTL(uFrm.SettingMailParametersFrm, oMainUfrm)
    Set oUctl = oSettingMailParametersFrm.AddCtl(uFrm.SettingMailParametersFrm.SettingMailProviderParameterSelectedCancelImg).SetSizes(NewSize(15, 15), NewSize(17, 17))
    Set oUctl = oSettingMailParametersFrm.AddCtl(uFrm.SettingMailParametersFrm.SettingMailProviderParameterSelectedSaveImg).SetSizes(NewSize(15, 15), NewSize(17, 17))
    Set oUctl = oSettingMailParametersFrm.AddCtl(uFrm.SettingMailParametersFrm.SettingMailProviderParameterSelectedDeleteImg).SetSizes(NewSize(15, 15), NewSize(17, 17))
    
    InitMailSettingUfrm uFrm

    uFrm.PymntExtractionDontSendFileOBtn.value = CBool(SettPymntExtractionDontSendFile.Val)
    uFrm.PymntExtractionSendFileToListOBtn.value = CBool(SettPymntExtractionSendFileToList.Val)
    uFrm.PymntExtractionSendFileToUsersOBtn.value = CBool(SettPymntExtractionSendFileToUsers.Val)
    uFrm.PymntExtractionSendFileToAllOBtn.value = CBool(SettPymntExtractionSendFileToAll.Val)
    uFrm.PymntExtractionFileFolderTBx.Text = CStr(SettPymntExtractionFileFolder.Val)
    InitPymntExtractionReceiversList uFrm
    
    ChangePymntExtractionSendFile uFrm
    UnselectPymntExtractionReceiver uFrm
End Sub

Public Sub InitSettingsAccess(uFrm As MSForms.UserForm)
    If loggedUser.Can(Array("parametre-nom_application")) Then
        uFrm.SettingsMultiPage.SystemPge.Enabled = True
    Else
        uFrm.SettingsMultiPage.SystemPge.Enabled = False
    End If
    
    If loggedUser.Can(Array("parametre-base_de_donnees")) Then
        uFrm.SettingsMultiPage.DbPge.Enabled = True
    Else
        uFrm.SettingsMultiPage.DbPge.Enabled = False
    End If
    
    If loggedUser.Can(Array("parametre-source_code")) Then
        uFrm.SettingsMultiPage.CodePge.Enabled = True
    Else
        uFrm.SettingsMultiPage.CodePge.Enabled = False
    End If
    
    If loggedUser.Can(Array("parametre-paiement")) Then
        uFrm.SettingsMultiPage.PaymentSettingsPge.Enabled = True
    Else
        uFrm.SettingsMultiPage.PaymentSettingsPge.Enabled = False
    End If

    If loggedUser.Can(Array("parametre-securite")) Then
        uFrm.SettingsMultiPage.SecuritySettingsPge.Enabled = True
    Else
        uFrm.SettingsMultiPage.SecuritySettingsPge.Enabled = False
    End If
End Sub

Public Sub SaveSystParamSettings(uFrm As MSForms.UserForm)
    
    currAppName = uFrm.SettingAppNameTBx.Text
    
    If SaveAPPSettings(True, False, False, False, False, False, False, False) Then
        MsgBox "Paramètres Système enregistés avec Succès", vbInformation, GetAppName
        uFrm.AppTitleLbl.Caption = GetAppName
    Else
        MsgBox "Erreur Enregistrement Paramètres Système", vbInformation, GetAppName
    End If
End Sub

Public Sub SaveDbAccessParamSettings(uFrm As MSForms.UserForm)
    
    currDbType = uFrm.SettingDbTypeTBx.Text
    currDbServerName = uFrm.SettingDbPathTBx.Text
    currDbName = uFrm.SettingDbNameTBx.Text
    currDbUserName = uFrm.SettingDbUserTBx.Text
    currDbUserPassword = uFrm.SettingDbPwdTBx.Text
    
    If SaveAPPSettings(False, True, False, False, False, False, False, False) Then
        MsgBox "Paramètres Accès à la Base de Données enregistés avec Succès", vbInformation, GetAppName
    Else
        MsgBox "Erreur Enregistrement Paramètres Accès à la Base de Données", vbInformation, GetAppName
    End If
End Sub

Public Sub SaveDbSelectParamSettings(uFrm As MSForms.UserForm)
    
    If Not IsNumeric(uFrm.SettingDbSelectLimitSizeTBx.Text) Then
        MsgBox "Taille LIMIT invalide!", vbCritical, GetAppName
        With uFrm.SettingDbSelectLimitSizeTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If

    currDbSelectLimitStatement = uFrm.SettingDbSelectLimitStatementTBx.Text
    currDbSelectLimitSize = CLng(uFrm.SettingDbSelectLimitSizeTBx.Text)
    currDbSelectLimitPosition = uFrm.SettingDbSelectLimitPositionCBx.Text
    
    If SaveAPPSettings(False, False, True, False, False, False, False, False) Then
        MsgBox "Paramètres SELECT de Base de Données enregistés avec Succès", vbInformation, GetAppName
    Else
        MsgBox "Erreur Enregistrement Paramètres SELECT de Base de Données", vbInformation, GetAppName
    End If
End Sub

Public Sub SaveCodeAccessParamSettings(uFrm As MSForms.UserForm)
    
    currOpenCode = uFrm.OpenCodeCBx.value
    
    If SaveAPPSettings(False, False, False, True, False, False, False, False) Then
        MsgBox "Paramètres Accès au Code enregistés avec Succès", vbInformation, GetAppName
    Else
        MsgBox "Erreur Enregistrement Paramètres Accès au Code", vbInformation, GetAppName
    End If
End Sub

Public Sub SaveSettingPayment(uFrm As MSForms.UserForm)

    If Not IsNumeric(uFrm.PhoneLengthTBx.Text) Then
        MsgBox "Taille Numéro Phone invalide!", vbCritical, GetAppName
        With uFrm.PhoneLengthTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If

    If Not IsNumeric(uFrm.PymntLastLinesSizeDashboardTBx.Text) Then
        MsgBox "Nombre de Lignes Dashboard invalide!", vbCritical, GetAppName
        With uFrm.PymntLastLinesSizeDashboardTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If
    
    If Not IsNumeric(uFrm.PymntEmployeesSearchLinesSizeTBx.Text) Then
        MsgBox "Nombre de Lignes Recherche Employés invalide!", vbCritical, GetAppName
        With uFrm.PymntEmployeesSearchLinesSizeTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If

    currPhoneNumLength = CInt(uFrm.PhoneLengthTBx.Text)
    currMotifSpecialChars = GetNewMofifSpecialChars(uFrm)
    
    currPymntStatusAttenteValidationId = uFrm.PymntStatusAttenteValidationIdTBx.Text
    currPymntStatusAttenteValidationLabel = uFrm.PymntStatusAttenteValidationLabelTBx.Text
    currPymntStatusAttenteExtractionId = uFrm.PymntStatusAttenteExtractionIdTBx.Text
    currPymntStatusAttenteExtractionLabel = uFrm.PymntStatusAttenteExtractionLabelTBx.Text
    currPymntStatusExtraitId = uFrm.PymntStatusExtraitIdTBx.Text
    currPymntStatusExtraitLabel = uFrm.PymntStatusExtraitLabelTBx.Text
    currPymntStatusFichierEnvoyeId = uFrm.PymntStatusFichierEnvoyeIdTBx.Text
    currPymntStatusFichierEnvoyeLabel = uFrm.PymntStatusFichierEnvoyeLabelTBx.Text
    currPymntStatusEffectueId = uFrm.PymntStatusEffectueIdTBx.Text
    currPymntStatusEffectueLabel = uFrm.PymntStatusEffectueLabelTBx.Text
    currPymntLastLinesSizeDashboard = CLng(uFrm.PymntLastLinesSizeDashboardTBx.Text)
    currPymntEmployeesSearchLinesSize = CLng(uFrm.PymntEmployeesSearchLinesSizeTBx.Text)

    currPymntNotifyAttenteValidation = CBool(uFrm.PymntAttenteValidationChBx.value)
    currPymntNotifyAttenteExtraction = CBool(uFrm.PymntAttenteExtractionNotifyChBx.value)
    currPymntNotifyExtrait = CBool(uFrm.PymntExtraitNotifyChBx.value)
    currPymntNotifyFichierEnvoye = CBool(uFrm.PymntFichierEnvoyeNotifyChBx.value)
    currPymntNotifyEffectue = CBool(uFrm.PymntEffectueNotifyChBx.value)
    
    If SaveAPPSettings(False, False, False, False, True, False, False, False) Then
        MsgBox "Paramètres Enregistrement Paiement enregistés avec Succès", vbInformation, GetAppName
    Else
        MsgBox "Erreur Enregistrement Paramètres Enregistrement Paiement !", vbInformation, GetAppName
    End If
End Sub

Public Sub SaveSettingPaymentSendFile(uFrm As MSForms.UserForm)
    currPymntExtractionDontSendFile = CBool(uFrm.PymntExtractionDontSendFileOBtn.value)
    currPymntExtractionSendFileToList = CBool(uFrm.PymntExtractionSendFileToListOBtn.value)
    currPymntExtractionSendFileToUsers = CBool(uFrm.PymntExtractionSendFileToUsersOBtn.value)
    currPymntExtractionSendFileToAll = CBool(uFrm.PymntExtractionSendFileToAllOBtn.value)
    currPymntExtractionFileFolder = uFrm.PymntExtractionFileFolderTBx.Text

    currPymntExtractionReceiversList = GetNewPymntExtractionReceivers(uFrm)

    If SaveAPPSettings(False, False, False, False, False, False, False, True) Then
        MsgBox "Paramètres Envoi Fichier Paiement enregistés avec Succès", vbInformation, GetAppName
    Else
        MsgBox "Erreur Enregistrement Paramètres Envoi Fichier Paiement !", vbInformation, GetAppName
    End If
End Sub

Public Sub SaveSettingSecurity(uFrm As MSForms.UserForm)
    
    currPwdLegnthMin = uFrm.PwdLegnthMinTBx.Text
    currPwdUppercaseMin = uFrm.PwdUppercaseMinTBx.Text
    currPwdNumberMin = uFrm.PwdNumberMinTBx.Text
    currPwdSpecialCharsMin = uFrm.PwdSpecialCharsMinTBx.Text
    currPwdValidity = uFrm.PwdValidityTBx.Text
    currPwdUpdCanBeSame = uFrm.PwdUpdCanBeSameChBx.value
    
    If SaveAPPSettings(False, False, False, False, False, True, False, False) Then
        MsgBox "Paramètres Sécurité enregistés avec Succès", vbInformation, GetAppName
    Else
        MsgBox "Erreur Enregistrement Paramètres Sécurité", vbInformation, GetAppName
    End If
End Sub

Public Sub SaveSettingMail(uFrm As MSForms.UserForm)
    
    currMailProviderSelected = CLng(uFrm.SettingMailProviderCBx.ListIndex)
    currMailProviderParametersSelected = GetNewMailProviderParameters(uFrm)
    
    If SaveAPPSettings(False, False, False, False, False, False, True, False) Then
        MsgBox "Paramètres Mail enregistés avec Succès", vbInformation, GetAppName
    Else
        MsgBox "Erreur Enregistrement Paramètres Mail", vbInformation, GetAppName
    End If
End Sub

'Public Sub ReadSettings()
'    Dim ws As Worksheet
'    Set ws = ActiveWorkbook.Sheets("APP-Settings")
    
    'sAppName = CStr(ws.Cells(1, 1))
    
    'gblDatabaseType = CStr(ws.Cells(2, 1))
    'gblDbServerName = CStr(ws.Cells(3, 1))
    'gblDbName = CStr(ws.Cells(4, 1))
    'gblDbUserName = CStr(ws.Cells(5, 1))
    'gblDbUserPassword = CStr(ws.Cells(6, 1))
    
    'bOpenCode = CBool(ws.Cells(7, 1) = "1")
    
    'iPhoneNumLength = CInt(ws.Cells(8, 1))
    'sMotifSpecialChars = CStr(ws.Cells(9, 1))
    
'    bSettingsLoaded = True
    
'End Sub

Private Function SaveAPPSettings(saveSystem As Boolean, saveDBAccess As Boolean, saveDBSelect As Boolean, saveCode As Boolean, savePaymentSetting As Boolean, saveSecuritySetting As Boolean, saveMailSetting As Boolean, SavePaymentSendFile As Boolean) As Boolean
    Dim ws As Worksheet, audit As clsAudit, auditAction As String
    
    Set ws = ActiveWorkbook.Sheets("APP-Settings")
    ws.Activate
    ActiveSheet.Cells.NumberFormat = "@"
    
    If saveSystem Then
        Set audit = loggedUser.StartNewAudit("Modifictaion Paramètres Nom Application, Nom: " & CStr(currAppName))
        
        SetAppName currAppName
        
        audit.EndWithSuccess
    End If
    
    If saveDBAccess Then
        Set audit = loggedUser.StartNewAudit("Modifictaion Paramètres Accès à la Base de Données, Type: " & CStr(currDbType) & ", Serveur/Emplacement: " & CStr(currDbServerName) & ", Nom Base de Données: " & CStr(currDbName) & ", Utilisateur: " & CStr(currDbUserName))
        
        SettDatabaseType.SaveValue currDbType
        SettDbServerName.SaveValue currDbServerName
        SettDbName.SaveValue currDbName
        SettDbUserName.SaveValue currDbUserName
        SettDbUserPassword.SaveValue currDbUserPassword
        
        audit.EndWithSuccess
    End If

    If saveDBSelect Then
        Set audit = loggedUser.StartNewAudit("Modifictaion Paramètres SELECT de Base de Données, Clause LIMIT: " & CStr(currDbSelectLimitStatement) & ", Taille LIMIT: " & CStr(currDbSelectLimitSize) & ", Position LIMIT: " & CStr(currDbSelectLimitPosition))
        
        SettDbSelectLimitStatement.SaveValue currDbSelectLimitStatement
        SettDbSelectLimitSize.SaveValue currDbSelectLimitSize
        SettDbSelectLimitPosition.SaveValue currDbSelectLimitPosition
        
        audit.EndWithSuccess
    End If
    
    If saveCode Then
        Set audit = loggedUser.StartNewAudit("Modifictaion Paramètre Accès au Code Source, Valeur: " & IIf(currOpenCode, "Oui", "Non"))
        
        SettOpenCode.SaveValue currOpenCode
        
        audit.EndWithSuccess
    End If
    
    If savePaymentSetting Then
        
        auditAction = "Modifictaion Paramètres Paiement, Taille Num. Phone: " & CStr(currPhoneNumLength)
        
        auditAction = auditAction & ", Id Attente Validation: " & CStr(currPymntStatusAttenteValidationId)
        auditAction = auditAction & ", Libelé Attente Validation: " & CStr(currPymntStatusAttenteValidationLabel)
        auditAction = auditAction & ", Notifier Paiement Attente Validation: " & IIf(currPymntNotifyAttenteValidation, "Oui", "Non")
        
        auditAction = auditAction & ", Id Attente Extraction: " & CStr(currPymntStatusAttenteExtractionId)
        auditAction = auditAction & ", Libelé Attente Extraction: " & CStr(currPymntStatusAttenteExtractionLabel)
        auditAction = auditAction & ", Notifier Paiement Attente Extraction: " & IIf(currPymntNotifyAttenteExtraction, "Oui", "Non")
        
        auditAction = auditAction & ", Id Extrait: " & CStr(currPymntStatusExtraitId)
        auditAction = auditAction & ", Libelé Extrait: " & CStr(currPymntStatusExtraitLabel)
        auditAction = auditAction & ", Notifier Paiement Extrait: " & IIf(currPymntNotifyExtrait, "Oui", "Non")
        
        auditAction = auditAction & ", Id Fichier Envoyé: " & CStr(currPymntStatusFichierEnvoyeId)
        auditAction = auditAction & ", Libelé Fichier Envoyé: " & CStr(currPymntStatusFichierEnvoyeLabel)
        auditAction = auditAction & ", Notifier Fichier Envoyé: " & IIf(currPymntNotifyFichierEnvoye, "Oui", "Non")
        
        auditAction = auditAction & ", Id Effectué: " & CStr(currPymntStatusEffectueId)
        auditAction = auditAction & ", Libelé Effectué: " & CStr(currPymntStatusEffectueLabel)
        auditAction = auditAction & ", Notifier Paiement Effectué: " & IIf(currPymntNotifyEffectue, "Oui", "Non")

        auditAction = auditAction & ", Nombre Lignes Dashboard: " & CStr(currPymntLastLinesSizeDashboard)
        auditAction = auditAction & ", Nombre Lignes pour la Recherche Employés: " & CStr(currPymntEmployeesSearchLinesSize)

        Set audit = loggedUser.StartNewAudit(auditAction)
        
        SettPhoneNumLength.SaveValue currPhoneNumLength
        SettMotifSpecialChars.SaveValue currMotifSpecialChars
        
        SettPymntStatusAttenteValidationId.SaveValue currPymntStatusAttenteValidationId
        SettPymntStatusAttenteValidationLabel.SaveValue currPymntStatusAttenteValidationLabel
        SettPymntStatusAttenteExtractionId.SaveValue currPymntStatusAttenteExtractionId
        SettPymntStatusAttenteExtractionLabel.SaveValue currPymntStatusAttenteExtractionLabel
        SettPymntStatusExtraitId.SaveValue currPymntStatusExtraitId
        SettPymntStatusExtraitLabel.SaveValue currPymntStatusExtraitLabel
        SettPymntStatusFichierEnvoyeId.SaveValue currPymntStatusFichierEnvoyeId
        SettPymntStatusFichierEnvoyeLabel.SaveValue currPymntStatusFichierEnvoyeLabel
        SettPymntStatusEffectueId.SaveValue currPymntStatusEffectueId
        SettPymntStatusEffectueLabel.SaveValue currPymntStatusEffectueLabel
        SettPymntLastLinesSizeDashboard.SaveValue currPymntLastLinesSizeDashboard
        SettPymntEmployeesSearchLinesSize.SaveValue currPymntEmployeesSearchLinesSize
        
        SettPymntNotifyAttenteValidation.SaveValue currPymntNotifyAttenteValidation
        SettPymntNotifyAttenteExtraction.SaveValue currPymntNotifyAttenteExtraction
        SettPymntNotifyExtraction.SaveValue currPymntNotifyExtrait
        SettPymntNotifyFichierEnvoye.SaveValue currPymntNotifyFichierEnvoye
        SettPymntNotifyEffectue.SaveValue currPymntNotifyEffectue
        
        audit.EndWithSuccess
    End If

    If SavePaymentSendFile Then
        auditAction = "Modifictaion Paramètres Envoi Fichier Paiement, Ne Pas Envoyer de Mail: " & IIf(currPymntExtractionDontSendFile, "Oui", "Non")
        
        auditAction = auditAction & ", Envoyer à la Liste Uniquement: " & IIf(currPymntExtractionSendFileToList, "Oui", "Non")
        auditAction = auditAction & ", Envoyer aux Utilisateurs Habilités: " & IIf(currPymntExtractionSendFileToUsers, "Oui", "Non")
        auditAction = auditAction & ", Envoyer à la Liste + Utilisateurs Habilités: " & IIf(currPymntExtractionSendFileToAll, "Oui", "Non")
        auditAction = auditAction & ", Chemin / Répertoire Fichier Paiement: " & currPymntExtractionFileFolder

        Set audit = loggedUser.StartNewAudit(auditAction)
        
        SettPymntExtractionDontSendFile.SaveValue currPymntExtractionDontSendFile
        SettPymntExtractionSendFileToList.SaveValue currPymntExtractionSendFileToList
        SettPymntExtractionSendFileToUsers.SaveValue currPymntExtractionSendFileToUsers
        SettPymntExtractionSendFileToAll.SaveValue currPymntExtractionSendFileToAll
        SettPymntExtractionFileFolder.SaveValue currPymntExtractionFileFolder

        SettPymntExtractionReceiversList.SaveValue currPymntExtractionReceiversList
        
        audit.EndWithSuccess
    End If
    
    If saveMailSetting Then
        auditAction = "Modifictaion Paramètres Mail"
        
        Set audit = loggedUser.StartNewAudit(auditAction)
        
        SettMailProviderSelected.SaveValue currMailProviderSelected
        
        SettMailProviderParametersSelected.ColNum = GetMailParameterColumn(CInt(currMailProviderSelected))
        SettMailProviderParametersSelected.SaveValue currMailProviderParametersSelected
        
        UpdateMailSettings
        
        audit.EndWithSuccess
    End If

    If saveSecuritySetting Then
        auditAction = "Modifictaion Paramètres Sécurité, Taille Min. Mot de Passe: " & CStr(currPwdLegnthMin)
        auditAction = auditAction & ", Min. Majuscule dans Mot de Passe: " & CStr(currPwdUppercaseMin)
        auditAction = auditAction & ", Min. Chiffre dans Mot de Passe: " & CStr(currPwdNumberMin)
        auditAction = auditAction & ", Min. Caractères spéciaux dans Mot de Passe: " & CStr(currPwdSpecialCharsMin)
        auditAction = auditAction & ", Validite Mot de Passe: " & CStr(currPwdValidity)
        auditAction = auditAction & ", Mot de Passe peut etre reconduit a la modification: " & CStr(currPwdUpdCanBeSame)
        
        Set audit = loggedUser.StartNewAudit(auditAction)
        
        SettPwdLegnthMin.SaveValue currPwdLegnthMin
        SettPwdUppercaseMin.SaveValue currPwdUppercaseMin
        SettPwdNumberMin.SaveValue currPwdNumberMin
        SettPwdSpecialCharsMin.SaveValue currPwdSpecialCharsMin
        SettPwdValidity.SaveValue currPwdValidity
        SettPwdUpdCanBeSame.SaveValue IIf(currPwdUpdCanBeSame, 1, 0)
        
        audit.EndWithSuccess
    End If
    
    SaveAPPSettings = True
End Function


Private Sub InitMofifSpecialChars(uFrm As MSForms.UserForm)
    Dim specialCharsArr() As String, specialCharSingleArr() As String, i As Integer
    
    uFrm.MotifSpecialCharsLBx.Clear
    uFrm.MotifSpecialCharsLBx.ColumnCount = 2
    uFrm.MotifSpecialCharsLBx.ColumnWidths = "80;80"
    
    specialCharsArr = Split(CStr(SettMotifSpecialChars.Val), ";")
    For i = 0 To UBound(specialCharsArr)
        uFrm.MotifSpecialCharsLBx.AddItem
        
        specialCharSingleArr = Split(specialCharsArr(i), ",")
        uFrm.MotifSpecialCharsLBx.List(i, 0) = specialCharSingleArr(0)
        uFrm.MotifSpecialCharsLBx.List(i, 1) = specialCharSingleArr(1)
    Next i
    
End Sub

Private Function GetNewMofifSpecialChars(uFrm As MSForms.UserForm)
    Dim newMotifSpecialChars As String, i As Integer
    
    With uFrm.MotifSpecialCharsLBx
        For i = 0 To .listCount - 1
            If i = 0 Then
                newMotifSpecialChars = .List(i, 0) & "," & .List(i, 1)
            Else
                newMotifSpecialChars = newMotifSpecialChars & ";" & .List(i, 0) & "," & .List(i, 1)
            End If
        Next i
    End With
    
    GetNewMofifSpecialChars = newMotifSpecialChars
End Function


'
'   Motif Special Chars
'
Private Sub UnselectMotifSpecialChar(uFrm As MSForms.UserForm)
    lMotifSpecialCharSelectedIndex = -1
    uFrm.MotifSpecialCharsLBx.ListIndex = -1
End Sub

Public Sub ResetMotifSpecialChar(uFrm As MSForms.UserForm)
    'uFrm.MotifSpecialCharSaveImg.Visible = False
    uFrm.MotifSpecialCharDeleteImg.Visible = False
    
    uFrm.MotifSpecialCharFromTBx.Text = ""
    uFrm.MotifSpecialCharToTBx.Text = ""
End Sub

Public Sub CancelMotifSpecialChar(uFrm As MSForms.UserForm)
    UnselectMotifSpecialChar uFrm
    ResetMotifSpecialChar uFrm
End Sub

Public Sub SelectMotifSpecialChar(uFrm As MSForms.UserForm)
    Dim i As Long
    
    ResetMotifSpecialChar uFrm
    
    With uFrm.MotifSpecialCharsLBx
        For i = 0 To .listCount - 1
          If .Selected(i) = True Then
            
            uFrm.MotifSpecialCharSaveImg.Visible = True
            uFrm.MotifSpecialCharDeleteImg.Visible = True
            
            uFrm.MotifSpecialCharFromTBx.Text = .List(i, 0)
            uFrm.MotifSpecialCharToTBx.Text = .List(i, 1)
            
            lMotifSpecialCharSelectedIndex = i
          End If
        Next i
    End With
End Sub

Public Sub SaveMotifSpecialChar(uFrm As MSForms.UserForm)
    Dim newCharFrom As String, newCharTo As String
    
    newCharFrom = uFrm.MotifSpecialCharFromTBx.Text
    newCharTo = uFrm.MotifSpecialCharToTBx.Text
    
    If lMotifSpecialCharSelectedIndex = -1 Then
        ' Add Special Char
        With uFrm.MotifSpecialCharsLBx
            .AddItem
            .List(.listCount - 1, 0) = newCharFrom
            .List(.listCount - 1, 1) = newCharTo
        End With
    Else
        ' Update Special Char
        uFrm.MotifSpecialCharsLBx.List(lMotifSpecialCharSelectedIndex, 0) = newCharFrom
        uFrm.MotifSpecialCharsLBx.List(lMotifSpecialCharSelectedIndex, 1) = newCharTo
        
    End If
    
    UnselectMotifSpecialChar uFrm
    ResetMotifSpecialChar uFrm
End Sub

Public Sub DeleteMotifSpecialChar(uFrm As MSForms.UserForm)
    If lMotifSpecialCharSelectedIndex <> -1 Then
        ' Remove Special Char
        uFrm.MotifSpecialCharsLBx.RemoveItem (lMotifSpecialCharSelectedIndex)
        
        UnselectMotifSpecialChar uFrm
        ResetMotifSpecialChar uFrm
    End If
End Sub

Public Sub InitDbSelectLimitPosition(uFrm As MSForms.UserForm)
    uFrm.SettingDbSelectLimitPositionCBx.Clear
    uFrm.SettingDbSelectLimitPositionCBx.AddItem "AVANT"
    uFrm.SettingDbSelectLimitPositionCBx.AddItem "APRES"
End Sub

'
' Mail Provider
'

Private Sub InitMailSettingUfrm(uFrm As MSForms.UserForm)
    InitMailProviders uFrm
    SetMailProviderParam uFrm, CInt(SettMailProviderSelected.Val)
End Sub

Private Sub InitMailProviders(uFrm As MSForms.UserForm)
    Dim providersArr() As String, i As Integer
    
    uFrm.SettingMailProviderCBx.Clear
    
    providersArr = SettMailProviders.GetArrayValue
    For i = 0 To UBound(providersArr)
        uFrm.SettingMailProviderCBx.AddItem providersArr(i)
    Next i
    
    uFrm.SettingMailProviderCBx.Text = MailProviderName
    
End Sub

Private Sub SetMailProviderParam(uFrm As MSForms.UserForm, selIndex As Integer)
    Dim specialCharsArr() As String, specialCharSingleArr() As String
    Dim providerParam As clsMailParameter
    
    uFrm.SettingMailProviderParametersLBx.Clear
    uFrm.SettingMailProviderParametersLBx.ColumnCount = 2
    uFrm.SettingMailProviderParametersLBx.ColumnWidths = "100;150"
    
    Set providerParam = GetMailProviderParametersSelected(selIndex)
    
    uFrm.SettingMailProviderDescLbl.Caption = providerParam.ProviderDescription
    
    uFrm.SettingMailProviderParametersLBx.AddItem
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 0) = "Email Expéditeur"
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 1) = providerParam.MailSender
    
    uFrm.SettingMailProviderParametersLBx.AddItem
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 0) = "Nom Expéditeur"
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 1) = providerParam.MailSenderName
    
    uFrm.SettingMailProviderParametersLBx.AddItem
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 0) = "Adresse Serveur Mail"
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 1) = providerParam.MailServerAddress
    
    uFrm.SettingMailProviderParametersLBx.AddItem
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 0) = "Port Serveur Mail"
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 1) = providerParam.MailServerPort
    
    uFrm.SettingMailProviderParametersLBx.AddItem
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 0) = "Nom Utilisateur"
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 1) = providerParam.MailUserName
    
    uFrm.SettingMailProviderParametersLBx.AddItem
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 0) = "Mot de Passe Utilisateur"
    uFrm.SettingMailProviderParametersLBx.List(uFrm.SettingMailProviderParametersLBx.listCount - 1, 1) = providerParam.MailUserPassword
    
    ResetMailProviderParameters uFrm
    
End Sub

Public Sub SelectMailProvider(uFrm As MSForms.UserForm)
    Dim selectedIndex As Integer
    
    selectedIndex = uFrm.SettingMailProviderCBx.ListIndex
    SetMailProviderParam uFrm, selectedIndex
End Sub

Public Sub SelectMailProviderParameter(uFrm As MSForms.UserForm)
    Dim i As Long
    
    ResetMailProviderParameters uFrm
    
    With uFrm.SettingMailProviderParametersLBx
        For i = 0 To .listCount - 1
          If .Selected(i) = True Then
            
            uFrm.SettingMailProviderParameterSelectedSaveImg.Visible = True
            'uFrm.SettingMailProviderParameterSelectedDeleteImg.Visible = True
            
            uFrm.SettingMailProviderParameterSelectedLabelTBx.Text = .List(i, 0)
            uFrm.SettingMailProviderParameterSelectedValueTBx.Text = .List(i, 1)
            
            lMailProviderParameterSelectedIndex = i
          End If
        Next i
    End With
End Sub

Public Sub SaveMailProviderParameter(uFrm As MSForms.UserForm)
    Dim newParameterLabel As String, newParameterValue As String
    
    newParameterLabel = uFrm.SettingMailProviderParameterSelectedLabelTBx.Text
    newParameterValue = uFrm.SettingMailProviderParameterSelectedValueTBx.Text
    
    If lMailProviderParameterSelectedIndex = -1 Then
        ' Add New Parameter
        With uFrm.SettingMailProviderParametersLBx
            .AddItem
            .List(.listCount - 1, 0) = newParameterLabel
            .List(.listCount - 1, 1) = newParameterValue
        End With
    Else
        ' Update Parameter
        uFrm.SettingMailProviderParametersLBx.List(lMailProviderParameterSelectedIndex, 0) = newParameterLabel
        uFrm.SettingMailProviderParametersLBx.List(lMailProviderParameterSelectedIndex, 1) = newParameterValue
        
    End If
    
    UnselectMailProviderParameter uFrm
    ResetMailProviderParameters uFrm
End Sub

Public Sub CancelMailProviderParameter(uFrm As MSForms.UserForm)
    UnselectMailProviderParameter uFrm
    ResetMailProviderParameters uFrm
End Sub

Private Sub UnselectMailProviderParameter(uFrm As MSForms.UserForm)
    lMailProviderParameterSelectedIndex = -1
    uFrm.SettingMailProviderParametersLBx.ListIndex = -1
End Sub

Public Sub ResetMailProviderParameters(uFrm As MSForms.UserForm)
    lMailProviderParameterSelectedIndex = -1
    
    uFrm.SettingMailProviderParameterSelectedSaveImg.Visible = False
    uFrm.SettingMailProviderParameterSelectedDeleteImg.Visible = False
    
    uFrm.SettingMailProviderParameterSelectedLabelTBx.Text = ""
    uFrm.SettingMailProviderParameterSelectedValueTBx.Text = ""
End Sub

Private Function GetNewMailProviderParameters(uFrm As MSForms.UserForm)
    Dim newMailProviderParameters As String, i As Integer
    
    newMailProviderParameters = uFrm.SettingMailProviderDescLbl.Caption
    With uFrm.SettingMailProviderParametersLBx
        For i = 0 To .listCount - 1
            newMailProviderParameters = newMailProviderParameters & CStr(SettMailProviderParametersSelected.ArraySep) & .List(i, 1)
        Next i
    End With
    
    GetNewMailProviderParameters = newMailProviderParameters
End Function

'
' Payment Extraction Send File
'
Private Sub InitPymntExtractionReceiversList(uFrm As MSForms.UserForm)
    Dim receiversArr() As String, currReceiverArr() As String, i As Integer
    
    uFrm.PymntExtractionReceiversListLBx.Clear
    uFrm.PymntExtractionReceiversListLBx.ColumnCount = 3
    uFrm.PymntExtractionReceiversListLBx.ColumnWidths = "150;150;10"
    
    receiversArr = Split(CStr(SettPymntExtractionReceiversList.Val), SettPymntExtractionReceiversList.ArraySep)
    For i = 0 To UBound(receiversArr)
        uFrm.PymntExtractionReceiversListLBx.AddItem
        
        currReceiverArr = Split(receiversArr(i), ";")
        uFrm.PymntExtractionReceiversListLBx.List(i, 0) = currReceiverArr(0)
        uFrm.PymntExtractionReceiversListLBx.List(i, 1) = currReceiverArr(1)
        uFrm.PymntExtractionReceiversListLBx.List(i, 2) = currReceiverArr(2)
    Next i
    
End Sub

Public Sub ChangePymntExtractionSendFile(uFrm As MSForms.UserForm)
    If uFrm.PymntExtractionDontSendFileOBtn.value Then
        ChangePymntExtractionReceiversListStatus uFrm, False
    ElseIf uFrm.PymntExtractionSendFileToListOBtn.value Then
        ChangePymntExtractionReceiversListStatus uFrm, True
    ElseIf uFrm.PymntExtractionSendFileToUsersOBtn.value Then
        ChangePymntExtractionReceiversListStatus uFrm, False
    ElseIf uFrm.PymntExtractionSendFileToAllOBtn.value Then
        ChangePymntExtractionReceiversListStatus uFrm, True
    Else
        ChangePymntExtractionReceiversListStatus uFrm, False
    End If
End Sub

Public Sub CancelPymntExtractionReceiver(uFrm As MSForms.UserForm)
    UnselectPymntExtractionReceiver uFrm
    ResetPymntExtractionReceiver uFrm
End Sub

Private Sub ChangePymntExtractionReceiversListStatus(uFrm As MSForms.UserForm, statusVal As Boolean)
    uFrm.PymntExtractionReceiverNameTBx.Enabled = statusVal
    uFrm.PymntExtractionReceiverEMailTBx.Enabled = statusVal

    uFrm.PymntExtractionReceiverSelCancelImg.Enabled = statusVal
    uFrm.PymntExtractionReceiverSelSaveImg.Enabled = statusVal
    uFrm.PymntExtractionReceiverSelDeleteImg.Enabled = statusVal

    uFrm.PymntExtractionReceiversListLBx.Enabled = statusVal
End Sub

Private Sub UnselectPymntExtractionReceiver(uFrm As MSForms.UserForm)
    lPymntExtractionReceiverSelectedIndex = -1
    uFrm.PymntExtractionReceiversListLBx.ListIndex = -1
End Sub

Public Sub ResetPymntExtractionReceiver(uFrm As MSForms.UserForm)
    'uFrm.PymntExtractionReceiverSelSaveImg.Visible = False
    uFrm.PymntExtractionReceiverSelDeleteImg.Visible = False
    
    uFrm.PymntExtractionReceiverNameTBx.Text = ""
    uFrm.PymntExtractionReceiverEMailTBx.Text = ""
    uFrm.PymntExtractionReceiverActivatedChBx.value = False
End Sub

Public Sub SelectPymntExtractionReceiver(uFrm As MSForms.UserForm)
    Dim i As Long
    
    ResetPymntExtractionReceiver uFrm
    
    With uFrm.PymntExtractionReceiversListLBx
        For i = 0 To .listCount - 1
          If .Selected(i) = True Then
            
            uFrm.PymntExtractionReceiverSelSaveImg.Visible = True
            uFrm.PymntExtractionReceiverSelDeleteImg.Visible = True
            
            uFrm.PymntExtractionReceiverNameTBx.Text = .List(i, 0)
            uFrm.PymntExtractionReceiverEMailTBx.Text = .List(i, 1)
            uFrm.PymntExtractionReceiverActivatedChBx.value = CBool(.List(i, 2))
            
            lPymntExtractionReceiverSelectedIndex = i
          End If
        Next i
    End With
End Sub

Public Sub SavePymntExtractionReceiver(uFrm As MSForms.UserForm)
    Dim newReceiverName As String, newReceiverEMail As String, newActivated As String
    
    If Not ValidEmail(uFrm.PymntExtractionReceiverEMailTBx.Text) Then
        MsgBox "Veuillez renseigner une adresse e-mail valide !", vbCritical, GetAppName
        Exit Sub
    Else
        newReceiverName = uFrm.PymntExtractionReceiverNameTBx.Text
        newReceiverEMail = uFrm.PymntExtractionReceiverEMailTBx.Text
        newActivated = IIf(uFrm.PymntExtractionReceiverActivatedChBx.value, "1", "0")

        If lPymntExtractionReceiverSelectedIndex = -1 Then
            ' Add Receiver
            With uFrm.PymntExtractionReceiversListLBx
                .AddItem
                .List(.listCount - 1, 0) = newReceiverName
                .List(.listCount - 1, 1) = newReceiverEMail
                .List(.listCount - 1, 2) = newActivated
            End With
        Else
            ' Update Receiver
            uFrm.PymntExtractionReceiversListLBx.List(lPymntExtractionReceiverSelectedIndex, 0) = newReceiverName
            uFrm.PymntExtractionReceiversListLBx.List(lPymntExtractionReceiverSelectedIndex, 1) = newReceiverEMail
            uFrm.PymntExtractionReceiversListLBx.List(lPymntExtractionReceiverSelectedIndex, 2) = newActivated
            
        End If
        
        UnselectPymntExtractionReceiver uFrm
        ResetPymntExtractionReceiver uFrm
    End If
End Sub

Public Sub DeletePymntExtractionReceiver(uFrm As MSForms.UserForm)
    If lPymntExtractionReceiverSelectedIndex <> -1 Then
        ' Remove Receiver
        uFrm.PymntExtractionReceiversListLBx.RemoveItem (lPymntExtractionReceiverSelectedIndex)
        
        UnselectPymntExtractionReceiver uFrm
        ResetPymntExtractionReceiver uFrm
    End If
End Sub

Private Function GetNewPymntExtractionReceivers(uFrm As MSForms.UserForm)
    Dim newPymntExtractionReceivers As String, i As Integer
    
    With uFrm.PymntExtractionReceiversListLBx
        For i = 0 To .listCount - 1
            If i = 0 Then
                newPymntExtractionReceivers = .List(i, 0) & ";" & .List(i, 1) & ";" & .List(i, 2)
            Else
                newPymntExtractionReceivers = newPymntExtractionReceivers & SettPymntExtractionReceiversList.ArraySep & .List(i, 0) & ";" & .List(i, 1) & ";" & .List(i, 2)
            End If
        Next i
    End With
    
    GetNewPymntExtractionReceivers = newPymntExtractionReceivers
End Function

