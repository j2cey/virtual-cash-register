Option Explicit

Public oLoginUfrm As CUFRM

Sub LoginUFrm_Show()
    LoginUFrm.Show 'vbModeless
End Sub

Public Sub InitLoginUFrm(uFrm As MSForms.UserForm)
    
    Set oLoginUfrm = NewUForm(uFrm, "", "", "", "", "", "")
    oLoginUfrm.HideBar
    
    Call InitGlobal
    
    SetLoggedUser
    
End Sub


Public Sub SetLoggedUser(Optional uFrm As MSForms.UserForm)
    
    If oLoggedUser Is Nothing Then
        Set oLoggedUser = New CModelUser
    End If
    
    If IsMissing(uFrm) Or (uFrm Is Nothing) Then
        MainUFrm.LoggedUserNameLbl.Caption = oLoggedUser.Name
    Else
        uFrm.LoggedUserNameLbl.Caption = oLoggedUser.Name
    End If
    If oLoggedUser.Role Is Nothing Then
        If IsMissing(uFrm) Or (uFrm Is Nothing) Then
            MainUFrm.LoggedUserRoleLbl.Caption = ""
        Else
            uFrm.LoggedUserRoleLbl.Caption = ""
        End If
    Else
        If IsMissing(uFrm) Or (uFrm Is Nothing) Then
            MainUFrm.LoggedUserRoleLbl.Caption = oLoggedUser.Role.Title
        Else
            uFrm.LoggedUserRoleLbl.Caption = loggedUser.Role.Title
        End If
    End If
    
    InitAllAccess uFrm
    
End Sub

Public Sub InitAllAccess(Optional uFrm As MSForms.UserForm)
    If IsMissing(uFrm) Or (uFrm Is Nothing) Then
        InitEmployeesAccess MainUFrm
        InitPaymentsAccess MainUFrm
        InitUsersAccess MainUFrm
        InitSettingsAccess MainUFrm
        InitAuditAccess MainUFrm
    Else
        InitEmployeesAccess uFrm
        InitPaymentsAccess uFrm
        InitUsersAccess uFrm
        InitSettingsAccess uFrm
        InitAuditAccess MainUFrm
    End If
End Sub