Option Explicit

Public Sub AccessMenuSettings(uFrm As MSForms.UserForm)
    If loggedUser.Can(Array("parametre-nom_application", "parametre-base_de_donnees", "parametre-source_code", "parametre-paiement")) Then
        InitSettingsForm uFrm
        SetActivePage uFrm, 4, "Settings", uFrm.SettingsMultiPage
        uFrm.SettingsMultiPage.Value = 0
    Else
        MsgBox "Vous n'êtes pas autorisé à accéder à cette Rubrique !", vbCritical, GetAppName
    End If
End Sub

Public Sub AccessMenuUsers(uFrm As MSForms.UserForm)
    If loggedUser.Can(Array("utilisateur-lister", "utilisateur-ajouter")) Then
        SetActivePage uFrm, 3, "Utilisateurs", uFrm.UsersMultiPage
    Else
        MsgBox "Vous n'êtes pas autorisé à accéder à cette Rubrique !", vbCritical, GetAppName
    End If
End Sub

Public Sub AccessMenuEmployees(uFrm As MSForms.UserForm)
    If loggedUser.Can(Array("employe-lister", "employe-ajouter")) Then
        SetActivePage uFrm, 1, "Employés", uFrm.EmployeesMultiPage
    Else
        MsgBox "Vous n'êtes pas autorisé à accéder à cette Rubrique !", vbCritical, GetAppName
    End If
End Sub

Public Sub AccessMenuPayments(uFrm As MSForms.UserForm)
    If loggedUser.Can(Array("paiement-lister", "paiement-ajouter")) Then
        SetActivePage uFrm, 2, "Paiements", uFrm.PaymentsMultiPage
    Else
        MsgBox "Vous n'êtes pas autorisé à accéder à cette Rubrique !", vbCritical, GetAppName
    End If
End Sub

Public Sub AccessMenuAudit(uFrm As MSForms.UserForm)
    If loggedUser.Can(Array("audit-lister")) Then
        AuditUFrm.Show
    Else
        MsgBox "Vous n'êtes pas autorisé à accéder à cette Rubrique !", vbCritical, GetAppName
    End If
End Sub