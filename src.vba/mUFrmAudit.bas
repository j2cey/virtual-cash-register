Option Explicit

Public oAuditUFrm As CUFRM, oAuditSearchForm As clsSearchForm

Private oAuditSearchFrm As CUIControl
Public oAuditSaveForm As clsSaveForm, auditSelectedIndex As Long, auditSelectedId As Long

Public Sub InitAudit(uFrm As MSForms.UserForm)
    Set oAuditUFrm = NewUForm(uFrm, "", "", "", "", "", "")
    
    Call oAuditUFrm.AddCtl(uFrm.CloseImg).SetSizes(NewSize(14, 14), NewSize(16, 16))
    
    ' Search Frame
    Set oAuditSearchFrm = NewUCTL(uFrm.AuditSearchFrm, oAuditUFrm)
    oAuditSearchFrm.AddCtl uFrm.AuditSearchImg
    oAuditSearchFrm.AddCtl uFrm.AuditSearchCancelImg

    InitSearchForm uFrm
    
    InitSaveForm uFrm
End Sub

Public Sub InitAuditAccess(uFrm As MSForms.UserForm)
    Call SetVisibility(uFrm.MenuAudit, True, Array("audit-lister"))
End Sub

Private Sub InitSearchForm(uFrm As MSForms.UserForm)
    ' ***   Item Search Form
    Dim searchUctl As CUIControl
    
    Set oAuditSearchForm = NewSearchForm("audittrail", uFrm.AuditSearchImg, uFrm.AuditListLBx, uFrm.AuditSearchCancelImg)
    Set oAuditSearchForm.SearchBtn = Nothing
    oAuditSearchForm.SetResultTitle uFrm.AuditListLbl, "Action", "Actions"
    
    oAuditSearchForm.LimitLines = 100
    
    Set searchUctl = oAuditSearchForm.AddFieldCtl(uFrm.SearchUsernameTBx, "username", "Utilisateur", True, True). _
    SetClearContentButton(uFrm.SearchUsernameCancelImg)
    Set searchUctl = oAuditSearchFrm.AddCtl(uFrm.AuditSearchFrm.SearchUsernameCancelImg).SetSizes(NewSize(8, 10), NewSize(10, 12))
    
    Set searchUctl = oAuditSearchForm.AddFieldCtl(uFrm.SearchActionTBx, "audit_action", "Action", True, True). _
    SetClearContentButton(uFrm.SearchActionCancelImg)
    Set searchUctl = oAuditSearchFrm.AddCtl(uFrm.AuditSearchFrm.SearchActionCancelImg).SetSizes(NewSize(8, 10), NewSize(10, 12))

    Set searchUctl = oAuditSearchForm.AddFieldCtl(uFrm.SearchStartedAtTBx, "started_at", "Date", True, True). _
    SetClearContentButton(uFrm.SearchStartedAtCancelImg)
    Set searchUctl = oAuditSearchFrm.AddCtl(uFrm.AuditSearchFrm.SearchStartedAtCancelImg).SetSizes(NewSize(8, 10), NewSize(10, 12))
End Sub

Private Sub InitSaveForm(uFrm As MSForms.UserForm)
    Dim editBtn As CUIControl, deleteBtn As CUIControl
    
    Set oAuditSaveForm = NewSaveForm("audittrail", uFrm.AuditSaveTitleLbl, uFrm.AuditSaveImg, uFrm.AuditCancelImg)
    Set oAuditSaveForm.SaveBtn = Nothing
    
    ' Save Titles
    oAuditSaveForm.AddSaveTitle None, "Gestion " & subListSingularTitle
    oAuditSaveForm.AddSaveTitle Add, "Ajouter " & subListSingularTitle
    oAuditSaveForm.AddSaveTitle Update, "Modifier " & subListSingularTitle
    oAuditSaveForm.AddSaveTitle Delete, "Suppression " & subListSingularTitle
    
    ' Save Controls
    oAuditSaveForm.AddFieldCtl uFrm.AuditUsernameTBx, "username", "Utilisateur", True, True, True, uFrm.AuditUsernameLbl
    oAuditSaveForm.AddFieldCtl uFrm.AuditActionTBx, "audit_action", "Action", True, True, True, uFrm.AuditActionLbl
    oAuditSaveForm.AddFieldCtl uFrm.AuditStartedAtTBx, "started_at", "Début", True, True, True, uFrm.AuditStartedAtLbl
    oAuditSaveForm.AddFieldCtl uFrm.AuditEndedAtTBx, "ended_at", "Fin", True, True, True, uFrm.AuditEndedAtLbl
    oAuditSaveForm.AddFieldCtl uFrm.AuditActionResultTBx, "action_result", "Résultat", True, True, True, uFrm.AuditActionResultLbl

    oAuditSaveForm.AddFieldCtl uFrm.AuditHostNameTBx, "host_name", "Nom Ordinateur", True, True, True, uFrm.AuditHostNameLbl
    oAuditSaveForm.AddFieldCtl uFrm.AuditHostIpTBx, "host_ip", "Adresse IP Ordinateur", True, True, True, uFrm.AuditHostIpLbl
    oAuditSaveForm.AddFieldCtl uFrm.AuditHostUserTBx, "host_user", "Utilisateur Ordinateur", True, True, True, uFrm.AuditHostUserLbl
    oAuditSaveForm.AddFieldCtl uFrm.AuditHostModelTBx, "host_model", "Modèle Ordinateur", True, True, True, uFrm.AuditHostModelLbl
    oAuditSaveForm.AddFieldCtl uFrm.AuditHostOsTBx, "host_os", "Système Exploitation", True, True, True, uFrm.AuditHostOsLbl
    oAuditSaveForm.AddFieldCtl uFrm.AuditHostOsVersionTBx, "host_osversion", "Version OS", True, True, True, uFrm.AuditHostOsVersionLbl
End Sub

Public Sub SearchAudit(uFrm As MSForms.UserForm)
    oAuditSearchForm.Search
    oAuditSaveForm.ResetForm None
End Sub

Public Sub SelectAudit(uFrm As MSForms.UserForm)
    If SelectFromSingleListBox(uFrm.AuditListLBx, auditSelectedIndex, auditSelectedId) Then
        EditAudit uFrm
    End If
End Sub

Public Sub EditAudit(uFrm As MSForms.UserForm)
    oAuditSaveForm.ExecEdit auditSelectedId
End Sub

Public Sub ResetAuditForms(uFrm As MSForms.UserForm)
    oAuditSearchForm.ResetForm
    oAuditSaveForm.ResetForm None
End Sub

