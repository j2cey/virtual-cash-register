Public empBodyFrm As CUIControl, oEmployeeSearchFrm As CUIControl, oEmployeeSaveCardBodyFrm As CUIControl, oEmployeeSearchForm As clsSearchForm, oEmployeeListFrm As CUIControl
Public oEmployeeSaveForm As clsSaveForm, sEmployeeSelectedId As String

Public Sub InitEmployees(uFrm As MSForms.UserForm)
    uFrm.EmployeesMultiPage.value = 0
    
    ' Main Body
    Set empBodyFrm = NewUCTL(uFrm.EmployeesBodyCardFrm, oMainUfrm)
    empBodyFrm.AddCtl uFrm.EmployeeNewImg
    
    ' Search Frame
    Set oEmployeeSearchFrm = NewUCTL(uFrm.EmployeeSearchFrm, oMainUfrm)
    oEmployeeSearchFrm.AddCtl uFrm.EmployeeSearchImg
    oEmployeeSearchFrm.AddCtl uFrm.EmployeeSearchCancelImg
    
    ' List Page
    Set oEmployeeListFrm = NewUCTL(uFrm.EmployeeListFrm, oMainUfrm)
    oEmployeeListFrm.AddCtl uFrm.EmployeeEditImg
    oEmployeeListFrm.AddCtl uFrm.EmployeeDeleteImg
    
    ' Save Page
    Set oEmployeeSaveCardBodyFrm = NewUCTL(uFrm.EmployeeSaveCardBodyFrm, oMainUfrm)
    oEmployeeSaveCardBodyFrm.AddCtl uFrm.EmployeeSaveImg
    oEmployeeSaveCardBodyFrm.AddCtl uFrm.EmployeeCancelImg
    
    Call oEmployeeSaveCardBodyFrm.AddCtl(uFrm.ManageEmployeeSiteImg).SetSizes(NewSize(14, 14), NewSize(16, 16))
    Call oEmployeeSaveCardBodyFrm.AddCtl(uFrm.ManageEmployeeDirectionImg).SetSizes(NewSize(14, 14), NewSize(16, 16))
    Call oEmployeeSaveCardBodyFrm.AddCtl(uFrm.ManageEmployeeDepartementImg).SetSizes(NewSize(14, 14), NewSize(16, 16))
    Call oEmployeeSaveCardBodyFrm.AddCtl(uFrm.ManageEmployeePosteImg).SetSizes(NewSize(14, 14), NewSize(16, 16))

    Call InitEmployeeSiteCBx(uFrm)
    Call InitEmployeeDirectionCBx(uFrm)
    Call InitEmployeeDepartementCBx(uFrm)
    Call InitEmployeePosteCBx(uFrm)
    
    Call InitSearchForm(uFrm)
    
    Call InitSaveForm(uFrm)
End Sub

Public Sub InitEmployeesAccess(uFrm As MSForms.UserForm)
    SetVisibility uFrm.EmployeeNewImg, True, Array("employe-ajouter")
    SetVisibility uFrm.EmployeeNewLbl, True, Array("employe-ajouter")
    
    SetVisibility uFrm.EmployeeSaveImg, True, Array("employe-ajouter", "paiement-modifer")
    SetVisibility uFrm.EmployeeSaveLbl, True, Array("employe-ajouter", "paiement-modifer")

    SetVisibility uFrm.ManageEmployeeSiteImg, True, Array("site_employe-lister", "site_employe-ajouter", "site_employe-modifier", "site_employe-supprimer")
    SetVisibility uFrm.ManageEmployeeDirectionImg, True, Array("direction_employe-lister", "direction_employe-ajouter", "direction_employe-modifier", "direction_employe-supprimer")
    SetVisibility uFrm.ManageEmployeeDepartementImg, True, Array("departement_employe-lister", "departement_employe-ajouter", "departement_employe-modifier", "departement_employe-supprimer")
    SetVisibility uFrm.ManageEmployeePosteImg, True, Array("poste_employe-lister", "poste_employe-ajouter", "poste_employe-modifier", "poste_employe-supprimer")
End Sub

Public Sub SelectEmployee(uFrm As MSForms.UserForm)
    Dim i As Long
    
    With uFrm.EmployeesListLBx
    For i = 0 To .listCount - 1
      If .Selected(i) = True Then
        If loggedUser.Can(Array("employe-modifer")) Then
            uFrm.EmployeeEditImg.Visible = True
        End If
        If loggedUser.Can(Array("paiement-supprimer")) Then
            uFrm.EmployeeDeleteImg.Visible = True
        End If
        sEmployeeSelectedId = .List(i, 0)
      End If
    Next i
  End With
End Sub

Public Sub SaveEmployee(uFrm As MSForms.UserForm)
    Dim audit As clsAudit, sAuditAction As String, checkEmployeeStr As String, sPhoneFormatted As String
    
    If oEmployeeSaveForm.Status = Add Then
        sAuditAction = "Création Nouvel Employé "
    Else
        sAuditAction = "Modification Employé "
    End If
    sAuditAction = sAuditAction & uFrm.EmployeeLastNameTBx.Text & " " & uFrm.EmployeeFirstNameTBx.Text & ", Mle: " & uFrm.EmployeeMatriculeTBx.Text
    Set audit = loggedUser.StartNewAudit(sAuditAction)
    
    ' Check Employee Matricule
    checkEmployeeStr = "SELECT COUNT(*) FROM employees WHERE matricule = '" & SqlStringVar(uFrm.EmployeeMatriculeTBx.Text) & "'"
    checkEmployeeStr = checkEmployeeStr & IIf(oEmployeeSaveForm.Status = Add, "", " AND Id NOT IN (" & sEmployeeSelectedId & ")")
    
    If GetCount(checkEmployeeStr) > 0 Then
        MsgBox "Un Employé possède déjà ce Matricule !" & vbCrLf & "Veuillez changer de Matricule.", vbCritical, GetAppName
        audit.EndWithFailure
        Exit Sub
    End If
    
    ' Check Employee Phone
    sPhoneFormatted = uFrm.EmployeeTelephoneTBx.Text
    If Not FormatPhone(sPhoneFormatted) Then
        audit.EndWithFailure
        Exit Sub
    End If
    
    If oEmployeeSaveForm.Save Then
        ResetEmployeeForms uFrm
        audit.EndWithSuccess
    Else
        audit.EndWithFailure
    End If
End Sub

Public Sub EditEmployee(uFrm As MSForms.UserForm)
    oEmployeeSaveForm.ExecEdit sEmployeeSelectedId
    uFrm.EmployeesMultiPage.value = 1
End Sub

Public Sub DeleteEmployee(uFrm As MSForms.UserForm)
    Dim answer As Integer, audit As clsAudit, sAuditAction As String
    
    answer = MsgBox("Supprimer cet Employé ?", vbQuestion + vbYesNo + vbDefaultButton2, GetAppName)
    
    If answer = vbYes Then
        sAuditAction = "Suppression Employé " & uFrm.EmployeeLastNameTBx.Text & " " & uFrm.EmployeeFirstNameTBx.Text & ", Mle: " & uFrm.EmployeeMatriculeTBx.Text
        Set audit = loggedUser.StartNewAudit(sAuditAction)
        
        oEmployeeSaveForm.ExecDelete sEmployeeSelectedId
        ResetEmployeeForms uFrm
        
        audit.EndWithSuccess
    End If
End Sub

Public Sub ResetEmployeeForms(uFrm As MSForms.UserForm)
    oEmployeeSearchForm.ResetForm
    uFrm.EmployeesMultiPage.value = 0
End Sub

Private Sub InitSearchForm(uFrm As MSForms.UserForm)
    ' ***   Employees Search Form
    Dim empEditBtn As CUIControl, empDeleteBtn As CUIControl, searchUctl As CUIControl
    
    Set oEmployeeSearchForm = NewSearchForm("employees", uFrm.EmployeeSearchImg, uFrm.EmployeesListLBx, uFrm.EmployeeSearchCancelImg)
    oEmployeeSearchForm.SetResultTitle uFrm.EmployeesListLbl, "Employé", "Employés"
    
    Set searchUctl = oEmployeeSearchForm.AddFieldCtl(uFrm.SearchMatriculeTBx, "matricule", "Matricule", True, True). _
    SetClearContentButton(uFrm.SearchMatriculeCancelImg)
    Set searchUctl = oEmployeeSearchFrm.AddCtl(uFrm.EmployeeSearchFrm.SearchMatriculeCancelImg).SetSizes(NewSize(10, 12), NewSize(12, 14))
    
    Set searchUctl = oEmployeeSearchForm.AddFieldCtl(uFrm.SearchLastNameTBx, "lastname", "Nom", True, True). _
    SetClearContentButton(uFrm.SearchLastNameCancelImg)
    Set searchUctl = oEmployeeSearchFrm.AddCtl(uFrm.EmployeeSearchFrm.SearchLastNameCancelImg).SetSizes(NewSize(10, 12), NewSize(12, 14))
    
    Set searchUctl = oEmployeeSearchForm.AddFieldCtl(uFrm.SearchFirstNameTBx, "firstname", "Prénom", True, True). _
    SetClearContentButton(uFrm.SearchFirstNameCancelImg)
    Set searchUctl = oEmployeeSearchFrm.AddCtl(uFrm.EmployeeSearchFrm.SearchFirstNameCancelImg).SetSizes(NewSize(10, 12), NewSize(12, 14))
    
    Set empEditBtn = oEmployeeSearchForm.AddFieldCtl(uFrm.EmployeeEditImg, "", "", False, False)
    empEditBtn.AddAction Default, setVisibilityState, False
    empEditBtn.AddAction Active, setVisibilityState, True, Array("employe-modifer")
    Set empDeleteBtn = oEmployeeSearchForm.AddFieldCtl(uFrm.EmployeeDeleteImg, "", "", False, False)
    empDeleteBtn.AddAction Default, setVisibilityState, False
    empDeleteBtn.AddAction Active, setVisibilityState, True, Array("employe-supprimer")
End Sub

Private Sub InitSaveForm(uFrm As MSForms.UserForm)
    ' ***   Employees Search Form
    Dim empEditBtn As CUIControl, empDeleteBtn As CUIControl
    
    Set oEmployeeSaveForm = NewSaveForm("employees", uFrm.EmployeeSaveTitleLbl, uFrm.EmployeeSaveImg, uFrm.EmployeeCancelImg, uFrm.EmployeeNewImg)
    Set oEmployeeSaveForm.SaveBtn = Nothing
    
    ' Save Titles
    oEmployeeSaveForm.AddSaveTitle None, ""
    oEmployeeSaveForm.AddSaveTitle Add, "Ajouter Employé"
    oEmployeeSaveForm.AddSaveTitle Update, "Modifier Employé"
    oEmployeeSaveForm.AddSaveTitle Delete, "Suppression Employé"
    
    ' Save Controls
    oEmployeeSaveForm.AddFieldCtl uFrm.EmployeeMatriculeTBx, "matricule", "Matricule", True, True, True, uFrm.EmployeeMatriculeLbl
    oEmployeeSaveForm.AddFieldCtl uFrm.EmployeeLastNameTBx, "lastname", "Nom", True, True, True, uFrm.EmployeeLastNameLbl
    oEmployeeSaveForm.AddFieldCtl uFrm.EmployeeFirstNameTBx, "firstname", "Prénom", True, True, True, uFrm.EmployeeFirstNameLbl
    oEmployeeSaveForm.AddFieldCtl uFrm.EmployeeSiteCBx, "site", "Site", True, True, True, uFrm.EmployeeSiteLbl
    oEmployeeSaveForm.AddFieldCtl uFrm.EmployeeDirectionCBx, "direction", "Dirction", True, True, True, uFrm.EmployeeDirectionLbl
    oEmployeeSaveForm.AddFieldCtl uFrm.EmployeeDepartementCBx, "departement", "Département", True, True, True, uFrm.EmployeeDepartementLbl
    oEmployeeSaveForm.AddFieldCtl uFrm.EmployeePosteCBx, "poste", "Poste", True, True, True, uFrm.EmployeePosteLbl
    oEmployeeSaveForm.AddFieldCtl uFrm.EmployeeImputationTBx, "imputation", "Imputation", True, True, True, uFrm.EmployeeImputationLbl
    oEmployeeSaveForm.AddFieldCtl uFrm.EmployeeTelephoneTBx, "telephone", "Téléphone", True, True, True, uFrm.EmployeeTelephoneLbl
    
End Sub

Private Sub InitEmployeeSiteCBx(uFrm As MSForms.UserForm)
    Dim oSql As clsSqlQuery, oSelectFields As Collection
    
    Set oSelectFields = New Collection
    oSelectFields.Add NewDbField("title", "Titre", 0, True, True, True, True, True), "title"
    Set oSql = NewSqlQuery("employeesites")
    oSql.SelectToListByCriterion NewUCTL(uFrm.EmployeeSiteCBx), oSelectFields, Nothing, False
End Sub

Public Sub ManageEmployeeSite(uFrm As MSForms.UserForm)
    ManageSubList "employeesites", "Site", "Sites", "site_employe-ajouter", "site_employe-modifier", "site_employe-supprimer"
    InitEmployeeSiteCBx uFrm
End Sub

Private Sub InitEmployeeDirectionCBx(uFrm As MSForms.UserForm)
    Dim oSql As clsSqlQuery, oSelectFields As Collection
    
    Set oSelectFields = New Collection
    oSelectFields.Add NewDbField("title", "Titre", 0, True, True, True, True, True), "title"
    Set oSql = NewSqlQuery("employeedirections")
    oSql.SelectToListByCriterion NewUCTL(uFrm.EmployeeDirectionCBx), oSelectFields, Nothing, False
End Sub

Public Sub ManageEmployeeDirection(uFrm As MSForms.UserForm)
    ManageSubList "employeedirections", "Direction", "Directions", "direction_employe-ajouter", "direction_employe-modifier", "direction_employe-supprimer"
    InitEmployeeDirectionCBx uFrm
End Sub

Private Sub InitEmployeeDepartementCBx(uFrm As MSForms.UserForm)
    Dim oSql As clsSqlQuery, oSelectFields As Collection
    
    Set oSelectFields = New Collection
    oSelectFields.Add NewDbField("title", "Titre", 0, True, True, True, True, True), "title"
    Set oSql = NewSqlQuery("employeedepartements")
    oSql.SelectToListByCriterion NewUCTL(uFrm.EmployeeDepartementCBx), oSelectFields, Nothing, False
End Sub

Public Sub ManageEmployeeDepartement(uFrm As MSForms.UserForm)
    ManageSubList "employeedepartements", "Département", "Départements", "departement_employe-ajouter", "departement_employe-modifier", "departement_employe-supprimer"
    InitEmployeeDepartementCBx uFrm
End Sub

Private Sub InitEmployeePosteCBx(uFrm As MSForms.UserForm)
    Dim oSql As clsSqlQuery, oSelectFields As Collection
    
    Set oSelectFields = New Collection
    oSelectFields.Add NewDbField("title", "Titre", 0, True, True, True, True, True), "title"
    Set oSql = NewSqlQuery("employeepostes")
    oSql.SelectToListByCriterion NewUCTL(uFrm.EmployeePosteCBx), oSelectFields, Nothing, False
End Sub

Public Sub ManageEmployeePoste(uFrm As MSForms.UserForm)
    ManageSubList "employeepostes", "Poste", "Postes", "poste_employe-ajouter", "poste_employe-modifier", "poste_employe-supprimer"
    InitEmployeePosteCBx uFrm
End Sub

