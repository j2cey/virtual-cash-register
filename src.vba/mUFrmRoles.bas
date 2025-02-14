Public roleBodyFrm As CUIControl, oRoleSearchFrm As CUIControl, oRoleSearchForm As clsSearchForm
Public oUserroleSaveFrm As CUIControl
Public sRoleSelectedId As String
Public iRolePermissionSelectedIndex As Integer, iPermissionSelectedIndex As Integer


Public Sub InitRoles(uFrm As MSForms.UserForm)
    ' Main Body
    Set roleBodyFrm = NewUCTL(uFrm.UserroleCardBodyFrm, oMainUfrm)
    roleBodyFrm.AddCtl uFrm.UserroleDeleteImg
    
    ' Search Frame
    Set oRoleSearchFrm = NewUCTL(uFrm.UserroleSearchFrm, oMainUfrm)
    oRoleSearchFrm.AddCtl uFrm.UserroleSearchImg
    oRoleSearchFrm.AddCtl uFrm.UserroleSearchCancelImg
    
    ' Save Form
    Set oUserroleSaveFrm = NewUCTL(uFrm.UserroleSaveFrm, oMainUfrm)
    oUserroleSaveFrm.AddCtl uFrm.AddPermissionToRoleImg
    oUserroleSaveFrm.AddCtl uFrm.AddAllPermissionsToRoleImg
    
    oUserroleSaveFrm.AddCtl uFrm.RemovePermissionFromRoleImg
    oUserroleSaveFrm.AddCtl uFrm.RemoveAllPermissionsFromRoleImg
    
    oUserroleSaveFrm.AddCtl uFrm.UserroleCancelImg
    oUserroleSaveFrm.AddCtl uFrm.UserroleSaveImg
    
    Permissions.LoadPermissionsToList uFrm.PermissionsLBx
    
    SetAddState uFrm
    ResetRoleForm uFrm
    InitRoleCbx uFrm
    
    Call InitSearchForm(uFrm)
    
End Sub

Private Sub InitSearchForm(uFrm As MSForms.UserForm)
    ' ***   Users Search Form
    Dim roleDeleteBtn As CUIControl, clr As String, searchUctl As CUIControl
    
    Set oRoleSearchForm = NewSearchForm("userroles", uFrm.UserroleSearchImg, uFrm.UserrolesListLBx, uFrm.UserroleSearchCancelImg)
    oRoleSearchForm.SetResultTitle uFrm.RolesListLbl, "Role", "Roles"
    
    oRoleSearchForm.AddFieldCtl uFrm.SearchRoleCBx, "role_title", "Role", True, True
    oRoleSearchForm.AddFieldCtl uFrm.RoleDescriptionTBx, "description", "Description", True, False, False, False, False
    
    Set roleDeleteBtn = oRoleSearchForm.AddFieldCtl(uFrm.UserroleDeleteImg, "", "", False, False)
    roleDeleteBtn.AddAction Default, setVisibilityState, False
    roleDeleteBtn.AddAction Active, setVisibilityState, True
End Sub

Public Sub SelectRole(uFrm As MSForms.UserForm)
    Dim i As Long, rolesAccess() As String
    
    With uFrm.UserrolesListLBx
    For i = 0 To .listCount - 1
      If .Selected(i) = True Then
        ' selected-role
        sRoleSelectedId = .List(i, 0)
        
        ' profile-to-edit
        Set roleToEdit = NewUserRole(CLng(sRoleSelectedId), CStr(.List(i, 1)), CStr(.List(i, 2)))
        
        ' profile-delete
        SetVisibility uFrm.UserroleDeleteImg, True, Array("profile-supprimer")
        
        ' fill edit-profile-form
        FillRoleEditForm uFrm
      End If
    Next i
  End With
End Sub

Public Sub UnSelectRole(uFrm As MSForms.UserForm)
    SetAddState uFrm
    ResetRoleForm uFrm
    sRoleSelectedId = ""
    
    oRoleSearchForm.Search True
End Sub

Public Sub UnSelectRolePermission(uFrm As MSForms.UserForm)
    uFrm.RolePermissionsLBx.ListIndex = -1
    iRolePermissionSelectedIndex = -1
    uFrm.RemovePermissionFromRoleImg.Visible = False
    uFrm.RemoveAllPermissionsFromRoleImg.Visible = False
End Sub

Public Sub SelectRolePermission(uFrm As MSForms.UserForm)
    Dim i As Long, rolesAccess() As String
    
    With uFrm.RolePermissionsLBx
    For i = 0 To .listCount - 1
      If .Selected(i) = True Then
        ' selected-profile-permission
        iRolePermissionSelectedIndex = i
        
        ' button
        UnSelectPermission uFrm
        
        SetVisibility uFrm.RemovePermissionFromRoleImg, True, Array("profile-modifer")
        SetVisibility uFrm.RemoveAllPermissionsFromRoleImg, True, Array("profile-modifer")
        
      End If
    Next i
  End With
End Sub

Private Sub UnSelectPermission(uFrm As MSForms.UserForm)
    uFrm.PermissionsLBx.ListIndex = -1
    iPermissionSelectedIndex = -1
    uFrm.AddPermissionToRoleImg.Visible = False
    uFrm.AddAllPermissionsToRoleImg.Visible = False
End Sub

Public Sub SelectPermission(uFrm As MSForms.UserForm)
    Dim i As Long, rolesAccess() As String
    
    With uFrm.PermissionsLBx
    For i = 0 To .listCount - 1
      If .Selected(i) = True Then
        ' selected-profile-permission
        iPermissionSelectedIndex = i
        
        ' buttons
        UnSelectRolePermission uFrm
        SetVisibility uFrm.AddPermissionToRoleImg, True, Array("profile-modifer")
        SetVisibility uFrm.AddAllPermissionsToRoleImg, True, Array("profile-modifer")
        
      End If
    Next i
  End With
End Sub

Public Sub AddPermissionToRole(uFrm As MSForms.UserForm)
    Dim sId As String, sGroup As String, sLevel As String, sPermission As String
    
    sId = CStr(uFrm.PermissionsLBx.List(iPermissionSelectedIndex, 0))
    sGroup = CStr(uFrm.PermissionsLBx.List(iPermissionSelectedIndex, 1))
    sLevel = CStr(uFrm.PermissionsLBx.List(iPermissionSelectedIndex, 2))
    sPermission = CStr(uFrm.PermissionsLBx.List(iPermissionSelectedIndex, 3))
    
    AddPermissionToRolePermissionList uFrm, sId, sGroup, sLevel, sPermission
    
    UnSelectPermission uFrm
End Sub

Public Sub AddAllPermissionsToRole(uFrm As MSForms.UserForm)
    Dim permissionIndex As Integer, i As Integer
    
    permissionIndex = -1
    With uFrm.PermissionsLBx
        For i = 0 To .listCount - 1
            AddPermissionToRolePermissionList uFrm, CStr(.List(i, 0)), CStr(.List(i, 1)), CStr(.List(i, 2)), CStr(.List(i, 3))
        Next i
    End With
    
    UnSelectPermission uFrm
End Sub

Private Sub AddPermissionToRolePermissionList(uFrm As MSForms.UserForm, Id As String, group As String, level As String, permission As String)
    Dim permissionIndex As Integer, i As Integer
    
    permissionIndex = -1
    With uFrm.RolePermissionsLBx
        For i = 0 To .listCount - 1
          If .List(i, 3) = permission Then
            permissionIndex = i
            Exit For
          End If
        Next i
    End With
    
    If permissionIndex = -1 Then
        ' lbx.ColumnWidths = "0;0;0;100" ' Id,Group,Level,Permission
        uFrm.RolePermissionsLBx.AddItem
        uFrm.RolePermissionsLBx.List(uFrm.RolePermissionsLBx.listCount - 1, 0) = Id
        uFrm.RolePermissionsLBx.List(uFrm.RolePermissionsLBx.listCount - 1, 1) = group
        uFrm.RolePermissionsLBx.List(uFrm.RolePermissionsLBx.listCount - 1, 2) = level
        uFrm.RolePermissionsLBx.List(uFrm.RolePermissionsLBx.listCount - 1, 3) = permission
    End If
End Sub

Public Sub RemovePermissionToRole(uFrm As MSForms.UserForm)
    uFrm.RolePermissionsLBx.RemoveItem (iRolePermissionSelectedIndex)
    UnSelectRolePermission uFrm
End Sub

Public Sub RemoveAllPermissionsToRole(uFrm As MSForms.UserForm)
    ResetRolePermissionsList uFrm
    UnSelectRolePermission uFrm
End Sub

Public Sub SaveRole(uFrm As MSForms.UserForm)
    Dim checkRoleStr As String
    
    ' Check unique Title
    checkRoleStr = "SELECT COUNT(*) FROM userroles WHERE role_title = '" & uFrm.RoleTitleTBx.Text & "'"
    checkRoleStr = checkRoleStr & IIf(sRoleSelectedId = "", "", " AND Id NOT IN (" & sRoleSelectedId & ")")
    If GetCount(checkRoleStr) > 0 Then
        MsgBox "Un Profile possède déjà ce Titre !" & vbCrLf & "Veuillez changer de Titre.", vbCritical, GetAppName
        Exit Sub
    End If
    
    If ValidateRole(uFrm) Then
        If sRoleSelectedId = "" Then
            AddRole uFrm
        Else
            UpdateRole uFrm
        End If
    End If
End Sub

Private Function ValidateRole(uFrm As MSForms.UserForm) As Boolean
    ' Validate Role Title
    If uFrm.RoleTitleTBx.Text = "" Then
        MsgBox "Veuillez renseigner un Titre", vbCritical, GetAppName
        With uFrm.RoleTitleTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        
        ValidateRole = False
        Exit Function
    End If
    
    ' Validate Role Description
    If uFrm.RoleDescriptionTBx.Text = "" Then
        MsgBox "Veuillez renseigner une Description", vbCritical, GetAppName
        With uFrm.RoleDescriptionTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        
        ValidateRole = False
        Exit Function
    End If
    
    ' Validate Role Permissions
    If uFrm.RolePermissionsLBx.listCount = 0 Then
        MsgBox "Veuillez attribuer au moins une Permission au Profile", vbCritical, GetAppName
        
        ValidateRole = False
        Exit Function
    End If
    
    ValidateRole = True
End Function

Private Sub AddRole(uFrm As MSForms.UserForm)
    Dim queryStr As String, newRoleId As Variant, createProgr As clsProgression, audit As clsAudit
    
    Set createProgr = StartNewProgression("Création Nouveau Profile", 1)
    Set audit = loggedUser.StartNewAudit("Ajout du Profile " & uFrm.RoleTitleTBx.Text & ", Description: " & uFrm.RoleDescriptionTBx.Text)
    ' Insert Role
    queryStr = "INSERT INTO userroles (role_title,description) VALUES ('" & SqlStringVar(uFrm.RoleTitleTBx.Text) & "', '" & SqlStringVar(uFrm.RoleDescriptionTBx.Text) & "')"
    
    Call PrepareDatabase
    If gobjDB.ExecuteActionQuery(queryStr, newRoleId) Then
        ' Insert Permissions role
        If InsertPermissionsRole(uFrm, CLng(newRoleId), createProgr) Then
            MsgBox "Profile Créé avec Succès", vbInformation, GetAppName
            UnSelectRole uFrm
            
            InitRoleCbx uFrm
            modUFrmUsers.InitUserRoleCbx uFrm
            
            audit.EndWithSuccess
        Else
            audit.EndWithFailure
        End If
    Else
        audit.EndWithFailure
        MsgBox "Erreur Insertion Role dans la Base de Données", vbCritical, GetAppName
    End If
    
    createProgr.AddDone 1, True
End Sub

Private Sub UpdateRole(uFrm As MSForms.UserForm)
    Dim queryStr As String, updProgr As clsProgression, audit As clsAudit
    
    Set updProgr = StartNewProgression("Modification Profile", 1)
    Set audit = loggedUser.StartNewAudit("Modification du Profile (" & sRoleSelectedId & ") " & uFrm.RoleTitleTBx.Text & ", Description: " & uFrm.RoleDescriptionTBx.Text)
    ' Update Role
    queryStr = "UPDATE userroles SET role_title = '" & SqlStringVar(uFrm.RoleTitleTBx.Text) & "', description = '" & SqlStringVar(uFrm.RoleDescriptionTBx.Text) & "' WHERE Id = " & sRoleSelectedId & ""
    
    Call PrepareDatabase
    If gobjDB.ExecuteActionQuery(queryStr) Then
        
        If InsertPermissionsRole(uFrm, CLng(sRoleSelectedId), updProgr) Then
            MsgBox "Profile Modifié avec Succès", vbInformation, GetAppName
            UnSelectRole uFrm
            modUFrmUsers.InitUserRoleCbx uFrm
            
            audit.EndWithSuccess
        Else
            audit.EndWithFailure
            MsgBox "Erreur Mise à Jour Profile dans la Base de Données", vbCritical, GetAppName
        End If
        
    Else
        audit.EndWithFailure
        MsgBox "Erreur Mise à Jour Paiement dans la Base de Données", vbCritical, GetAppName
    End If
    
    updProgr.AddDone 1, True
End Sub

Private Function InsertPermissionsRole(uFrm As MSForms.UserForm, lRoleId As Long, Optional insertProgr As clsProgression) As Boolean
    Dim queryStr As String, sPermissionsList As String, i As Integer, newRolepermissionId As Integer, recordCount As Long, rolepermissionsData As Variant
    Dim audit As clsAudit
    
    StartNewSubProgressionFromParent "Ajout Permissions au Profile", 1, insertProgr
    
    sPermissionsList = ""
    With uFrm.RolePermissionsLBx
        For i = 0 To .listCount - 1
            If sPermissionsList = "" Then
                sPermissionsList = "("
            Else
                sPermissionsList = sPermissionsList & ","
            End If
            sPermissionsList = sPermissionsList & "'" & SqlStringVar(CStr(.List(i, 3))) & "'"
        Next i
    End With
    sPermissionsList = sPermissionsList & ")"
    
    Call PrepareDatabase
    queryStr = "DELETE FROM rolepermissions WHERE userrole_id = " & CStr(lRoleId) & " AND role_permission NOT IN " & sPermissionsList
    If gobjDB.ExecuteActionQuery(queryStr) Then
        
        With uFrm.RolePermissionsLBx
            For i = 0 To .listCount - 1
                
                AddToDoLastSubFromParent 1, insertProgr
                
                queryStr = "SELECT * FROM rolepermissions WHERE userrole_id = " & CStr(lRoleId) & " AND role_permission = '" & SqlStringVar(CStr(.List(i, 3))) & "'"
                rolepermissionsData = gobjDB.GetRecordsetToArray(queryStr, recordCount)
                If recordCount = 0 Then
                    Dim role_permission As String
                    
                    role_permission = CStr(.List(i, 3))
                    Set audit = loggedUser.StartNewAudit("Ajout de la Permission " & role_permission & " au Profile (" & CStr(lRoleId) & ") " & uFrm.RoleTitleTBx.Text)
                    
                    queryStr = "INSERT INTO rolepermissions (userrole_id,role_permission,permission_group,permission_level) VALUES (" & CStr(lRoleId) & ", '" & SqlStringVar(role_permission) & "','" & SqlStringVar(CStr(.List(i, 1))) & "'," & CStr(.List(i, 2)) & ")"
                    gobjDB.ExecuteActionQuery queryStr, newRolepermissionId
                    
                    audit.EndWithSuccess
                End If
                
                AddDoneLastSubFromParent 1, True, insertProgr
            Next i
        End With
        
        AddDoneLastSubFromParent 1, True, insertProgr
        InsertPermissionsRole = True
    Else
        MsgBox "Erreur Supression des Permissions retirées !", vbCritical, GetAppName
        
        createProgr.AddDone 1, True
        InsertPermissionsRole = False
        
        AddDoneLastSubFromParent 1, True, insertProgr
        Exit Function
    End If
End Function

Public Sub DeleteRole(uFrm As MSForms.UserForm)
    Dim answer As Integer
    Dim queryStr As String, delProgr As clsProgression, audit As clsAudit
    
    answer = MsgBox("Supprimer ce Profile ?", vbQuestion + vbYesNo + vbDefaultButton2, GetAppName)
    
    If answer = vbYes Then
        ' Delete Role
        Set delProgr = StartNewProgression("Modification Profile", 1)
        
        Set audit = loggedUser.StartNewAudit("Suppression du Profile (" & sRoleSelectedId & ") " & uFrm.RoleTitleTBx.Text & ", Description: " & uFrm.RoleDescriptionTBx.Text)
        queryStr = "DELETE FROM userroles WHERE Id = " & sRoleSelectedId & ""
        Call PrepareDatabase
        If gobjDB.ExecuteActionQuery(queryStr) Then
            MsgBox "Profile supprimé avec Succès", vbInformation, GetAppName
            
            audit.EndWithSuccess
            UnSelectRole uFrm
            
            InitRoleCbx uFrm
            modUFrmUsers.InitUserRoleCbx uFrm
        Else
            audit.EndWithFailure
            MsgBox "Erreur Supression Profile !", vbCritical, GetAppName
        End If
        
        delProgr.AddDone 1, True
    End If
    
End Sub

Private Sub SetAddState(uFrm As MSForms.UserForm)
    uFrm.UserroleDetailsFrm.Caption = "Créer Nouveau Profile"
    
    uFrm.UserroleDeleteImg.Visible = False
    
    SetVisibility uFrm.UserroleSaveImg, True, Array("profile-ajouter", "profile-modifer")
    SetVisibility uFrm.UserroleSaveLbl, True, Array("profile-ajouter", "profile-modifer")
End Sub

Private Sub FillRoleEditForm(uFrm As MSForms.UserForm)
    
    uFrm.UserroleDetailsFrm.Caption = "Modifier Profile"
    ResetRoleForm uFrm
    
    uFrm.RoleTitleTBx.Text = roleToEdit.Title
    uFrm.RoleDescriptionTBx.Text = roleToEdit.Description
    
    roleToEdit.Permissions.LoadPermissionsToList uFrm.RolePermissionsLBx
    roleToEdit.LoadUsersInRoleToList uFrm.UsersInRoleLBx
    
    SetVisibility uFrm.UserroleSaveImg, True, Array("profile-ajouter", "profile-modifer")
    SetVisibility uFrm.UserroleSaveLbl, True, Array("profile-ajouter", "profile-modifer")
End Sub

Private Sub ResetRoleForm(uFrm As MSForms.UserForm)
    uFrm.RoleTitleTBx.Text = ""
    uFrm.RoleDescriptionTBx.Text = ""
    
    ResetRolePermissionsList uFrm
    ResetUsersInRoleList uFrm
    
    UnSelectPermission uFrm
    UnSelectRolePermission uFrm
End Sub

Private Sub ResetRolePermissionsList(uFrm As MSForms.UserForm)
    uFrm.RolePermissionsLBx.Clear
    uFrm.RolePermissionsLBx.ColumnCount = 4
    uFrm.RolePermissionsLBx.ColumnWidths = "0;0;0;100" ' Id,Groupe,Level,role_permission
End Sub

Private Sub ResetUsersInRoleList(uFrm As MSForms.UserForm)
    uFrm.UsersInRoleLBx.Clear
    uFrm.UsersInRoleLBx.ColumnCount = 2
    uFrm.UsersInRoleLBx.ColumnWidths = "0;100" ' Id,username
End Sub

Private Sub InitRoleCbx(uFrm As MSForms.UserForm)
    Dim oSql As clsSqlQuery, oSelectFields As Collection
    
    Set oSelectFields = New Collection
    oSelectFields.Add NewDbField("role_title", "Titre", 0, True, True, True, True, True), "role_title"
    Set oSql = NewSqlQuery("userroles")
    
    'uFrm.SearchUserRoleCBx.ColumnCount = 2
    'uFrm.SearchUserRoleCBx.ColumnWidths = ";0"
    'uFrm.UserRoleCBx.ColumnCount = 2
    'uFrm.UserRoleCBx.ColumnWidths = ";0"
    
    oSql.SelectToListByCriterion NewUCTL(uFrm.SearchRoleCBx), oSelectFields, Nothing, False
    
End Sub