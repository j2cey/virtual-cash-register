Option Explicit

Public oMainUfrm As CUFRM

Sub MainUFrm_Show()
    MainUFrm.Show 'vbModeless
End Sub

Public Sub InitMainUFrm(uFrm As MSForms.UserForm)
    'Dim mainMultiPage As MSForms.MultiPage
    
    'Set mainMultiPage = uFrm.mainMultiPage
    InitDashbord uFrm
    uFrm.MainMultiPage.Value = 0
    
    'mainMultiPage.Style = fmTabStyleNone
    'mainMultiPage.value = 0
    
    uFrm.AppTitleLbl.Caption = GetAppName
    Set oMainUfrm = NewUForm(uFrm, "", "", "", "", "", "")
    
    SetLoggedUser uFrm
    
    'oMainUfrm.HideBar
End Sub

Public Sub SetActivePage(uFrm As MSForms.UserForm, iPageIndex As Long, sPageTitle As String, Optional pageMultiPage As MSForms.MultiPage)
    uFrm.MainMultiPage.Value = iPageIndex
    uFrm.PageTitleLbl.Caption = sPageTitle
    
    If Not IsMissing(pageMultiPage) Then
        If Not pageMultiPage Is Nothing Then
            pageMultiPage.Value = 0
        End If
    End If
End Sub

Public Sub ManageSubList(sTable As String, sSingularTitle As String, sPluralTitle As String, sAddPermission As String, sUpdatePermission As String, sDeletePermission As String)
    subListTable = sTable
    subListSingularTitle = sSingularTitle
    subListPluralTitle = sPluralTitle
    
    subListAddPermission = sAddPermission
    subListUpdatePermission = sUpdatePermission
    subListDeletePermission = sDeletePermission

    ManageSubListUFrm.Show 'vbModeless
End Sub
