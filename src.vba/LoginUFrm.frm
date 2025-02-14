
Option Explicit

Private Sub UserForm_Initialize()
    InitLoginUFrm Me
    
    'Dim mailRecipients As Variant
    
    'mailRecipients = AddToArray(mailRecipients, "j.ngomnze@moov-africa.ga")
    'mailRecipients = AddToArray(mailRecipients, "jud10parfait@gmail.com")
    'MailObject.MailSubject = "Test With Classes"
    'MailObject.SendMail "Hi Test", "Here is the Body Mail", "j.ngomnze@moov-africa.ga", "jud10parfait@gmail.com"
    
    'Dim fileAttach As String
    'fileAttach = AppPath & Application.PathSeparator & "mobicashpayment_.csv"
    'MailObject.SendMailToMany "Hi Test", "Here is the Body Mail", GetUserMailsByPermissions(Array("parametre-base_de_donnees", "parametre-source_code")), fileAttach
    
    Application.Visible = False
End Sub

Private Sub CancelCBtn_Click()
    
    Unload Me
    CloseAndSaveOpenWorkbooks
End Sub

Private Sub ValidateCBtn_Click()
    LaunchTryLogin
End Sub

Private Sub PwdTBx_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.value = vbKeyReturn Then
       LaunchTryLogin
    End If
End Sub

Private Sub LaunchTryLogin()
    
    If checkDbExists Then
        loggedUser.TryLogIn Me.LoginTBx.Text, Me.PwdTBx.Text
        
        If loggedUser.IsLogged Then
            
            If loggedUser.IsPwdExpired Then
                EditPwd
            End If
            
            If Not loggedUser.IsPwdExpired Then
                loggedUser.SaveLastLogin
                Unload Me
                SetLoggedUser
                MainUFrm_Show
            End If
        End If
    End If
    
End Sub

Private Sub EditPwd()
    Dim answer As Integer
    
    answer = MsgBox("Votre Mot de Passe a expiré. Vous devez le Modifier", vbQuestion + vbYesNo + vbDefaultButton2, GetAppName)
    
    If answer = vbYes Then
        EditLoggedUserPwd Me
    End If
End Sub