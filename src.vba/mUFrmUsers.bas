Public userBodyFrm As CUIControl, oUserSearchFrm As CUIControl, oUserSaveCardBodyFrm As CUIControl, oUserSearchForm As clsSearchForm, oUserListFrm As CUIControl
Public sUserSelectedId As String


Public Sub InitUsers(uFrm As MSForms.UserForm, oUserSaveForm As clsSaveForm)
    uFrm.UsersMultiPage.Value = 0
    
    Set oUserSaveForm.SaveBtn = Nothing
    
    ' Main Body
    Set userBodyFrm = NewUCTL(uFrm.UsersBodyCardFrm, oMainUfrm)
    userBodyFrm.AddCtl uFrm.UserNewImg
    userBodyFrm.AddCtl uFrm.UserrolesImg
    
    ' Search Frame
    Set oUserSearchFrm = NewUCTL(uFrm.UserSearchFrm, oMainUfrm)
    oUserSearchFrm.AddCtl uFrm.UserSearchImg
    oUserSearchFrm.AddCtl uFrm.UserSearchCancelImg
    
    ' List Page
    Set oUserListFrm = NewUCTL(uFrm.UserListFrm, oMainUfrm)
    oUserListFrm.AddCtl uFrm.UserEditImg
    oUserListFrm.AddCtl uFrm.UserDeleteImg
    oUserListFrm.AddCtl uFrm.UserEditPwdImg
    
    ' Save Page
    Set oUserSaveCardBodyFrm = NewUCTL(uFrm.UserSaveCardBodyFrm, oMainUfrm)
    oUserSaveCardBodyFrm.AddCtl uFrm.UserSaveImg
    oUserSaveCardBodyFrm.AddCtl uFrm.UserCancelImg
    
    Call InitSearchForm(uFrm)
    
    Call InitSaveForm(uFrm, oUserSaveForm)
    
    Call InitUserRoleCbx(uFrm)
End Sub

Public Sub InitUsersAccess(uFrm As MSForms.UserForm)
    
    ' Users
    SetVisibility uFrm.UserNewImg, True, Array("utilisateur-ajouter")
    SetVisibility uFrm.UserNewLbl, True, Array("utilisateur-ajouter")
    
    SetVisibility uFrm.UserSaveImg, True, Array("utilisateur-ajouter", "utilisateur-modifer")
    SetVisibility uFrm.UserSaveLbl, True, Array("utilisateur-ajouter", "utilisateur-modifer")
    
    ' Roles
    SetVisibility uFrm.UserrolesImg, True, Array("profile-lister", "profile-ajouter", "profile-modifer", "profile-supprimer", "profile-modifier_permissions")
    SetVisibility uFrm.UserrolesLbl, True, Array("profile-lister", "profile-ajouter", "profile-modifer", "profile-supprimer", "profile-modifier_permissions")
    
    SetVisibility uFrm.UserroleSaveImg, True, Array("profile-ajouter", "profile-modifer")
    SetVisibility uFrm.UserroleSaveLbl, True, Array("profile-ajouter", "profile-modifer")
    
    ' Self Edit Password
    If loggedUser.IsLogged Then
        uFrm.LoggedUserEditPwdLbl.Visible = True
    Else
        uFrm.LoggedUserEditPwdLbl.Visible = False
    End If
End Sub

Public Sub SelectUser(uFrm As MSForms.UserForm)
    Dim i As Long, can_count As Integer
    
    With uFrm.UsersListLBx
    For i = 0 To .listCount - 1
      If .Selected(i) = True Then
        
        can_count = 0
        
        If SetVisibility(uFrm.UserEditImg, True, Array("utilisateur-modifer")) Then
            can_count = can_count + 1
        End If
        
        If SetVisibility(uFrm.UserDeleteImg, True, Array("utilisateur-supprimer")) Then
            can_count = can_count + 1
        End If
        
        If loggedUser.Can(Array("utilisateur-modifier_mot_de_passe")) Then
            uFrm.UserEditPwdImg.Visible = True
            uFrm.UserEditPwdLbl.Visible = True
            can_count = can_count + 1
        End If
        
        If can_count > 0 Then
            sUserSelectedId = .List(i, 0)
            Set userToEdit = New clsUser
            userToEdit.Id = CLng(sUserSelectedId)
            userToEdit.Name = CStr(.List(i, 2))
        End If
      End If
    Next i
  End With
End Sub

Public Sub UnselectUser(uFrm As MSForms.UserForm)
    uFrm.UserEditPwdImg.Visible = False
    uFrm.UserEditPwdLbl.Visible = False
End Sub

Public Sub SaveUser(uFrm As MSForms.UserForm, oUserSaveForm As clsSaveForm)
    Dim audit As clsAudit, sAuditAction As String, checkUserStr As String
    
    sAuditAction = uFrm.UserNameTBx.Text & ", Login: " & uFrm.UserLoginTBx.Text & ", E-Mail: " & uFrm.UserMailTBx.Text & ", Profile: " & uFrm.UserRoleCBx.Text
    
    If Not ValidEmail(uFrm.UserMailTBx.Text) Then
        Set audit = loggedUser.StartNewAudit("Enregistrement Utilisateur" & sAuditAction)
        audit.EndWithFailure
        MsgBox "Veuillez renseigner une adresse e-mail valide !", vbCritical, GetAppName
        Exit Sub
    Else
        ' Check unique login
        checkUserStr = "SELECT COUNT(*) FROM users WHERE userlogin = '" & uFrm.UserLoginTBx.Text & "'"
        checkUserStr = checkUserStr & IIf(oUserSaveForm.Status = Add, "", " AND Id NOT IN (" & sUserSelectedId & ")")
        If GetCount(checkUserStr) > 0 Then
            audit.EndWithFailure
            MsgBox "Un Utilisateur possède déjà ce Login !" & vbCrLf & "Veuillez changer de Login.", vbCritical, GetAppName
            Exit Sub
        End If
        
        If oUserSaveForm.Status = Add Then
            If Not IsPwdComplexityMatched(uFrm.UserPwdTBx.Text) Then
                audit.EndWithFailure
                Exit Sub
            End If
            
            If (uFrm.UserPwdTBx.Text <> uFrm.UserRePwdTBx.Text) Then
                oUserSaveForm.SetFieldCtlErrorStyle uFrm.UserRePwdTBx.Name
                
                audit.EndWithFailure
                MsgBox "Les Mots de Passe doivent être identiques", vbCritical, GetAppName
                Exit Sub
            End If
            
            sAuditAction = "Création Utilisateur" & sAuditAction
            Set audit = loggedUser.StartNewAudit(sAuditAction)
            
            If oUserSaveForm.Save Then
                Dim NewUser As clsUser
                
                Set NewUser = New clsUser
                NewUser.Id = oUserSaveForm.LastResult
                NewUser.SavePwd uFrm.UserPwdTBx.Text, False
                
                ResetUserForms uFrm, oUserSaveForm
                
                audit.EndWithSuccess
            Else
                audit.EndWithFailure
            End If
        Else
            sAuditAction = "Modifictaion  Utilisateur(" & sUserSelectedId & ") " & sAuditAction
            Set audit = loggedUser.StartNewAudit(sAuditAction)
            
            If oUserSaveForm.Save Then
                audit.EndWithSuccess
                ResetUserForms uFrm, oUserSaveForm
            Else
                audit.EndWithFailure
            End If
        End If
    End If
End Sub

Public Sub EditUser(uFrm As MSForms.UserForm, oUserSaveForm As clsSaveForm)
    oUserSaveForm.ExecEdit sUserSelectedId
    uFrm.UsersMultiPage.Value = 1
End Sub

Public Sub FormatUserSaveForm(uFrm As MSForms.UserForm, oUserSaveForm As clsSaveForm, bValue As Boolean)
    oUserSaveForm.SetUctlRequiered uFrm.UserPwdTBx.Name, bValue
    oUserSaveForm.SetUctlRequiered uFrm.UserRePwdTBx.Name, bValue
    
    uFrm.UserPwdTBx.Visible = bValue
    uFrm.UserRePwdTBx.Visible = bValue
    
    uFrm.UserPwdLbl.Visible = bValue
    uFrm.UserRePwdLbl.Visible = bValue
End Sub

Public Sub EditUserPwd(uFrm As MSForms.UserForm)
    EditPwdUFrm.Show
End Sub

Public Sub EditLoggedUserPwd(uFrm As MSForms.UserForm)
    Set userToEdit = loggedUser
    'userToEdit.id = CLng(sUserSelectedId)
    'userToEdit.Name = CStr(.list(i, 2))
    EditPwdUFrm.Show
End Sub

Public Sub DeleteUser(uFrm As MSForms.UserForm, oUserSaveForm As clsSaveForm)
    Dim answer As Integer, audit As clsAudit
    
    answer = MsgBox("Supprimer cet Utilisateur ?", vbQuestion + vbYesNo + vbDefaultButton2, GetAppName)
    
    If answer = vbYes Then
        Set audit = loggedUser.StartNewAudit("Suppression Utilisateur (" & sUserSelectedId & ") " & uFrm.UserNameTBx.Text)
        
        oUserSaveForm.ExecDelete sUserSelectedId
        ResetUserForms uFrm, oUserSaveForm
        
        audit.EndWithSuccess
    End If
End Sub

Private Sub ResetUserForms(uFrm As MSForms.UserForm, oUserSaveForm As clsSaveForm)
    oUserSearchForm.ResetForm
    uFrm.UsersMultiPage.Value = 0
    UnselectUser uFrm
End Sub

Private Sub InitSearchForm(uFrm As MSForms.UserForm)
    ' ***   Users Search Form
    Dim userEditBtn As CUIControl, userDeleteBtn As CUIControl, clr As String, searchUctl As CUIControl
    
    ' ErrorColor
    clr = uFrm.SearchUserLoginTBx.BackColor
    clr = uFrm.SearchUserNameTBx.BackColor
    
    Set oUserSearchForm = NewSearchForm("users_view", uFrm.UserSearchImg, uFrm.UsersListLBx, uFrm.UserSearchCancelImg)
    oUserSearchForm.SetResultTitle uFrm.UsersListLbl, "Utilisateur", "Utilisateurs"
    
    Set searchUctl = oUserSearchForm.AddFieldCtl(uFrm.SearchUserLoginTBx, "userlogin", "Login", True, True). _
    SetClearContentButton(uFrm.UserSearchFrm.SearchUserLoginCancelImg)
    Set searchUctl = oUserSearchFrm.AddCtl(uFrm.UserSearchFrm.SearchUserLoginCancelImg).SetSizes(NewSize(10, 12), NewSize(12, 14))
    
    Set searchUctl = oUserSearchForm.AddFieldCtl(uFrm.SearchUserNameTBx, "username", "Nom", True, True). _
    SetClearContentButton(uFrm.SearchUserNameCancelImg)
    Set searchUctl = oUserSearchFrm.AddCtl(uFrm.UserSearchFrm.SearchUserNameCancelImg).SetSizes(NewSize(10, 12), NewSize(12, 14))
    
    oUserSearchForm.AddFieldCtl uFrm.SearchUserRoleCBx, "role_title", "Profile", True, True
    
    Set userEditBtn = oUserSearchForm.AddFieldCtl(uFrm.UserEditImg, "", "", False, False)
    userEditBtn.AddAction Default, setVisibilityState, False
    userEditBtn.AddAction Active, setVisibilityState, True
    Set userDeleteBtn = oUserSearchForm.AddFieldCtl(uFrm.UserDeleteImg, "", "", False, False)
    userDeleteBtn.AddAction Default, setVisibilityState, False
    userDeleteBtn.AddAction Active, setVisibilityState, True
End Sub

Private Sub InitSaveForm(uFrm As MSForms.UserForm, oUserSaveForm As clsSaveForm)
    ' ***   Users Search Form
    Dim userEditBtn As CUIControl, userDeleteBtn As CUIControl
    
    'Set oUserSaveForm = NewSaveForm("users", uFrm.UserSaveTitleLbl, uFrm.UserSaveImg, uFrm.UserCancelImg, uFrm.UserNewImg)
    ' Save Titles
    oUserSaveForm.AddSaveTitle None, ""
    oUserSaveForm.AddSaveTitle Add, "Ajouter Utilisateur"
    oUserSaveForm.AddSaveTitle Update, "Modifier Utilisateur"
    oUserSaveForm.AddSaveTitle Delete, "Suppression Utilisateur"
    
    ' Save Controls
    oUserSaveForm.AddFieldCtl uFrm.UserLoginTBx, "userlogin", "Login", True, True, True, uFrm.UserLoginLbl
    oUserSaveForm.AddFieldCtl uFrm.UserNameTBx, "username", "Nom", True, True, True, uFrm.UserNameLbl
    oUserSaveForm.AddFieldCtl uFrm.UserMailTBx, "usermail", "E-Mail", True, True, True, uFrm.UserMailLbl
    oUserSaveForm.AddFieldCtl uFrm.UserRoleCBx, "role_id", "Role", True, True, True, uFrm.UserRoleLbl
    oUserSaveForm.AddFieldCtl uFrm.UserPwdTBx, "userpwd", "Mot de Passe", True, True, True, uFrm.UserPwdLbl
    oUserSaveForm.AddFieldCtl uFrm.UserRePwdTBx, "userpwd", "Mot de Passe Valide", True, False, True, uFrm.UserRePwdLbl
    
End Sub

Public Sub InitUserRoleCbx(uFrm As MSForms.UserForm)
    Dim oSql As clsSqlQuery, oSelectFields As Collection
    
    Set oSelectFields = New Collection
    oSelectFields.Add NewDbField("role_title", "Titre", 0, True, True, True, True, True), "role_title"
    Set oSql = NewSqlQuery("userroles")
    
    'uFrm.SearchUserRoleCBx.ColumnCount = 2
    'uFrm.SearchUserRoleCBx.ColumnWidths = ";0"
    uFrm.UserRoleCBx.ColumnCount = 2
    uFrm.UserRoleCBx.ColumnWidths = ";0"
    
    oSql.SelectToListByCriterion NewUCTL(uFrm.SearchUserRoleCBx), oSelectFields, Nothing, False
    oSql.SelectToListByCriterion NewUCTL(uFrm.UserRoleCBx), oSelectFields, Nothing, True
    
End Sub
