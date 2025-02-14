Option Explicit

Public subListTable As String, subListSingularTitle As String, subListPluralTitle As String, subListAddPermission As String, subListUpdatePermission As String, subListDeletePermission As String

Public oManageSubListUFrm As CUFRM, oItemSearchForm As clsSearchForm

Private oItemSearchFrm As CUIControl
Public oItemSaveForm As clsSaveForm, itemSelectedIndex As Long, itemSelectedId As Long

Public Sub InitManageSubList(uFrm As MSForms.UserForm)
    uFrm.ItemSaveTitleLbl.Caption = "Gestion " & subListSingularTitle
    uFrm.ItemSearchFrm.Caption = "Recherche " & subListSingularTitle
    
    Set oManageSubListUFrm = NewUForm(uFrm, "", "", "", "", "", "")
    
    Call oManageSubListUFrm.AddCtl(uFrm.CloseImg).SetSizes(NewSize(14, 14), NewSize(16, 16))
    oManageSubListUFrm.AddCtl uFrm.ItemDeleteImg
    oManageSubListUFrm.AddCtl uFrm.ItemCancelImg
    oManageSubListUFrm.AddCtl uFrm.ItemSaveImg
    
    ' Search Frame
    Set oItemSearchFrm = NewUCTL(uFrm.ItemSearchFrm, oManageSubListUFrm)
    oItemSearchFrm.AddCtl uFrm.ItemSearchImg
    oItemSearchFrm.AddCtl uFrm.ItemSearchCancelImg

    InitSearchForm uFrm
    
    InitSaveForm uFrm
End Sub

Private Sub InitSearchForm(uFrm As MSForms.UserForm)
    ' ***   Item Search Form
    Dim deleteBtn As CUIControl, searchUctl As CUIControl
    
    Set oItemSearchForm = NewSearchForm(subListTable, uFrm.ItemSearchImg, uFrm.ItemsListLBx, uFrm.ItemSearchCancelImg)
    oItemSearchForm.SetResultTitle uFrm.ItemsListLbl, subListSingularTitle, subListPluralTitle
    
    Set searchUctl = oItemSearchForm.AddFieldCtl(uFrm.SearchTitleTBx, "title", "Titre", True, True). _
    SetClearContentButton(uFrm.SearchTitleCancelImg)
    Set searchUctl = oItemSearchFrm.AddCtl(uFrm.ItemSearchFrm.SearchTitleCancelImg).SetSizes(NewSize(8, 10), NewSize(10, 12))
    
    Set searchUctl = oItemSearchForm.AddFieldCtl(uFrm.SearchDescriptionTBx, "description", "Description", True, True). _
    SetClearContentButton(uFrm.SearchDescriptionCancelImg)
    Set searchUctl = oItemSearchFrm.AddCtl(uFrm.ItemSearchFrm.SearchDescriptionCancelImg).SetSizes(NewSize(8, 10), NewSize(10, 12))
    
    Set deleteBtn = oItemSearchForm.AddFieldCtl(uFrm.ItemDeleteImg, "", "", False, False)
    deleteBtn.AddAction Default, setVisibilityState, False
    deleteBtn.AddAction Active, setVisibilityState, True, Array(subListDeletePermission)
End Sub

Private Sub InitSaveForm(uFrm As MSForms.UserForm)
    Dim editBtn As CUIControl, deleteBtn As CUIControl
    
    Set oItemSaveForm = NewSaveForm(subListTable, uFrm.ItemSaveTitleLbl, uFrm.ItemSaveImg, uFrm.ItemCancelImg)
    Set oItemSaveForm.SaveBtn = Nothing
    
    ' Save Titles
    oItemSaveForm.AddSaveTitle None, "Gestion " & subListSingularTitle
    oItemSaveForm.AddSaveTitle Add, "Ajouter " & subListSingularTitle
    oItemSaveForm.AddSaveTitle Update, "Modifier " & subListSingularTitle
    oItemSaveForm.AddSaveTitle Delete, "Suppression " & subListSingularTitle
    
    ' Save Controls
    oItemSaveForm.AddFieldCtl uFrm.ItemTitleTBx, "title", "Titre", True, True, True, uFrm.ItemTitleLbl
    oItemSaveForm.AddFieldCtl uFrm.ItemDescriptionTBx, "description", "Description", True, True, True, uFrm.ItemDescriptionLbl
End Sub

Public Sub SearchItem(uFrm As MSForms.UserForm)
    oItemSaveForm.ResetForm None
End Sub

Public Sub SelectItem(uFrm As MSForms.UserForm)
    If SelectFromSingleListBox(uFrm.ItemsListLBx, itemSelectedIndex, itemSelectedId) Then
        EditItem uFrm
        
        If loggedUser.Can(Array(subListUpdatePermission)) Then
            uFrm.ItemSaveImg.Visible = True
        End If
        If loggedUser.Can(Array(subListDeletePermission)) Then
            uFrm.ItemDeleteImg.Visible = True
        End If
    End If
End Sub

Public Sub SaveItem(uFrm As MSForms.UserForm)
    Dim audit As clsAudit, sAuditAction As String, checkSublistStr As String
    
    ' Check unique Title
    checkSublistStr = "SELECT COUNT(*) FROM " & subListTable & " WHERE title = '" & uFrm.ItemTitleTBx.Text & "'"
    checkSublistStr = checkSublistStr & IIf(oItemSaveForm.Status = Add, "", " AND Id NOT IN (" & itemSelectedId & ")")
    If GetCount(checkSublistStr) > 0 Then
        MsgBox "Ce Titre existe déjà !" & vbCrLf & "Veuillez changer de Titre.", vbCritical, GetAppName
        Exit Sub
    End If
    
    If oItemSaveForm.Status = Add Then
        sAuditAction = "Création " & subListSingularTitle
    Else
        sAuditAction = "Modification " & subListSingularTitle
    End If
    sAuditAction = sAuditAction & " " & uFrm.ItemTitleTBx.Text & " " & ", Description: " & uFrm.ItemDescriptionTBx.Text
    Set audit = loggedUser.StartNewAudit(sAuditAction)
    
    If oItemSaveForm.Save Then
        ResetItemForms uFrm, None
        audit.EndWithSuccess
    Else
        audit.EndWithFailure
    End If
End Sub

Public Sub EditItem(uFrm As MSForms.UserForm)
    oItemSaveForm.ExecEdit itemSelectedId
End Sub

Public Sub DeleteItem(uFrm As MSForms.UserForm)
    Dim answer As Integer, audit As clsAudit, sAuditAction As String
    
    answer = MsgBox("Supprimer cet Element ?", vbQuestion + vbYesNo + vbDefaultButton2, GetAppName)
    
    If answer = vbYes Then
        sAuditAction = "Suppression " & subListSingularTitle & uFrm.ItemTitleTBx.Text
        Set audit = loggedUser.StartNewAudit(sAuditAction)
        
        oItemSaveForm.ExecDelete itemSelectedId
        ResetItemForms uFrm, None
        
        audit.EndWithSuccess
    End If
End Sub

Public Sub ResetItemForms(uFrm As MSForms.UserForm, Optional vSaveStatus As saveFormStatusType)
    oItemSearchForm.ResetForm
    oItemSaveForm.ResetForm vSaveStatus
    
    If vSaveStatus = Add And loggedUser.Can(Array(subListAddPermission)) Then
        uFrm.ItemSaveImg.Visible = True
    Else
        uFrm.ItemSaveImg.Visible = False
    End If
End Sub

