Option Explicit

Private m_id As Long
Private m_title As String
Private m_description As String

Private m_permissions As CPermissions
Private m_usersInRole As Collection

'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CUserRole
    Set Self = Me
End Property

' Id
Public Property Let Id(lId As Long)
    m_id = lId
End Property

Public Property Get Id() As Long
    Id = m_id
End Property

' Title
Public Property Let Title(sTitle As String)
    m_title = sTitle
End Property

Public Property Get Title() As String
    Title = m_title
End Property

' Description
Public Property Let Description(sDescription As String)
    m_description = sDescription
End Property

Public Property Get Description() As String
    Description = m_description
End Property

' Permissions
Public Property Set Permissions(oPermissions As CPermissions)
    Set m_permissions = oPermissions
End Property

Public Property Get Permissions() As CPermissions
    Set Permissions = m_permissions
End Property

' UsersInRole
Public Property Set UsersInRole(oUsersInRole As Collection)
    Set m_usersInRole = oUsersInRole
End Property

Public Property Get UsersInRole() As Collection
    Set UsersInRole = m_usersInRole
End Property



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(lId As Long, sTitle As String, Optional sDescription As String)
    Me.Id = lId
    Me.Title = sTitle
    
    If Not IsMissing(sDescription) Then
        Me.Description = sDescription
    End If
    
    LoadRolePermissions
    LoadUsersInRole
End Sub

Public Sub LoadUsersInRoleToList(lbx As MSForms.ListBox)
    Dim itemColl As Variant, currUser As CUser
    
    ResetUsersInRoleList lbx
    
    For Each itemColl In Me.UsersInRole
        Set currUser = itemColl
        
        lbx.AddItem
        lbx.List(lbx.listCount - 1, 0) = CStr(currUser.Id)
        lbx.List(lbx.listCount - 1, 1) = CStr(currUser.Name)
    Next itemColl
End Sub

Public Function HasPermission(sPermission As String) As Boolean
    HasPermission = Me.Permissions.IsPermissionInList(sPermission)
End Function


'*****************************************************************************************************************************
'    Private functions and subs
'*****************************************************************************************************************************

Private Sub ResetUsersInRoleList(lbx As MSForms.ListBox)
    
    lbx.Clear
    lbx.ColumnCount = 2
    lbx.ColumnWidths = "0;100" ' Id,username
    
End Sub

Private Sub LoadRolePermissions()
    Dim permissionsData As Variant, sqlRst As String, recordCount As Long, i As Long
    
    sqlRst = "SELECT * FROM rolepermissions WHERE userrole_id = " & CStr(Me.Id) & ""
    
    Call PrepareDatabase
    permissionsData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    If recordCount = 0 Then
        'MsgBox "Aucune Permission pour le role " & Me.Title & " dans la Base de Données !", vbCritical, GetAppName
        Exit Sub
    End If
    
    For i = 0 To recordCount - 1
        Me.Permissions.AddPermission CStr(permissionsData(3, i)), CStr(permissionsData(2, i)), CInt(permissionsData(4, i))
    Next i
    
End Sub

Private Sub LoadUsersInRole()
    Dim usersData As Variant, sqlRst As String, recordCount As Long, i As Long
    
    sqlRst = "SELECT * FROM users_view WHERE role_id = " & CStr(Me.Id) & ""
    
    Call PrepareDatabase
    usersData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    If recordCount = 0 Then
        'MsgBox "Aucun Utilisateur pour le role " & Me.Title & " dans la Base de Données !", vbCritical, GetAppName
        Exit Sub
    End If
    
    For i = 0 To recordCount - 1
        Me.UsersInRole.Add NewUser(CLng(usersData(0, i)), CStr(usersData(1, i)), CStr(usersData(2, i)))
    Next i
    
End Sub


'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    Set m_permissions = New CPermissions
    Set m_usersInRole = New Collection
End Sub