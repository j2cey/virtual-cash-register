
Private oUfrm As clsUFRM


Private Sub UserForm_Initialize()
    Set oUfrm = NewUForm(Me)
    
    oUfrm.AddCtl Me.EditPwdCancelImg
    oUfrm.AddCtl Me.EditPwdSaveImg
    
    If userToEdit Is Nothing Then
        MsgBox "Aucun Utilisateur selectionné", vbCritical, GetAppName
        Unload Me
    Else
        Me.EditPwdUserNameLbl.Caption = userToEdit.Name
    End If
End Sub

Private Sub EditPwdCancelImg_Click()
    Unload Me
End Sub

Private Sub EditPwdSaveImg_Click()
    TrySavePwd
End Sub

Private Sub UserRePwdTBx_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyReturn Then
       TrySavePwd
    End If
End Sub

Private Sub TrySavePwd()
    If Me.UserPwdTBx.Text = "" Then
        MsgBox "Veuillez renseigner le Mot de Passe", vbCritical, GetAppName
        With Me.UserPwdTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If
    
    If Me.UserRePwdTBx.Text = "" Then
        MsgBox "Veuillez confirmer le Mot de Passe", vbCritical, GetAppName
        With Me.UserRePwdTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If
    
    If (Me.UserPwdTBx.Text <> Me.UserRePwdTBx.Text) Then
        MsgBox "Les Mots de Passe doivent être identiques", vbCritical, GetAppName
        Exit Sub
    End If
    
    If Not IsPwdComplexityMatched(Me.UserPwdTBx.Text, userToEdit) Then
        Exit Sub
    End If
    
    If Not userToEdit Is Nothing Then
        If userToEdit.SavePwd(UserPwdTBx) Then
            MsgBox "Mot de Passe modifié avec succès", vbInformation, GetAppName
            Unload Me
        Else
            MsgBox "Une erreur innatendue s'est produite veuillez contacter l'Administrateur !", vbCritical, GetAppName
        End If
    End If
End Sub