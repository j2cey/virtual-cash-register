'Build 000
'*****************************************************************************************************************************
'
' Authors:  Jude Parfait NGOM NZE, jud10parfait@gmail.com
'
' (c)2024, all rights reserved to the authors
'
'*****************************************************************************************************************************
'-------------------------------------------------------------------------
' Module    : CPermissions
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 11-06-2024
' Purpose   : Manage All the Permissions of the System
'-------------------------------------------------------------------------
Option Explicit


Private m_oGroups As Collection                 ' all Permissions Groups, Collection of Group Collections



'*****************************************************************************************************************************
'    Public events
'*****************************************************************************************************************************



'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CPermissions
    Set Self = Me
End Property

' Groups
Public Property Set Groups(oGroups As Collection)
    m_oGroups = oGroups
End Property

Public Property Get Groups() As Collection
    Set Groups = m_oGroups
End Property



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************


Public Sub AddPermission(sGroupKey As String, sPermission As String, iPermissionLevel As Integer)
    Dim group As Collection
    
    Set group = GetOrCreateGroup(sGroupKey)
    
    If CollKeyExists(group, sPermission) Then
        'MsgBox sPermission & " permission already exists in Group " & sGroupKey & " !", vbCritical, GetAppName
    Else
        group.Add NewPermissionNode(sGroupKey, sPermission, iPermissionLevel), sPermission
    End If
End Sub

Public Sub LoadPermissionsToList(lbx As MSForms.ListBox)
    Dim itemColl As Variant, group As Collection
    
    ResetPermissionsList lbx
    
    For Each itemColl In Me.Groups
        Set group = itemColl
        
        AddGroupPermissionsToList group, lbx
    Next itemColl
End Sub

Public Function IsPermissionInList(sPermission As String) As Boolean
    Dim itemColl As Variant, group As Collection
    
    For Each itemColl In Me.Groups
        Set group = itemColl
        If CollKeyExists(group, sPermission) Then
            IsPermissionInList = True
            Exit Function
        End If
    Next itemColl
    
    IsPermissionInList = False
End Function


'*****************************************************************************************************************************
'    Private functions and subs
'*****************************************************************************************************************************

Private Function GetOrCreateGroup(sGroupKey As String) As Collection
    Dim group As Collection, itemColl As Variant, permNode As clsPermissionNode
    
    If CollKeyExists(Me.Groups, sGroupKey) Then
        Set group = Me.Groups(sGroupKey)
    Else
        Set group = New Collection
        Me.Groups.Add group, sGroupKey
    End If
    
    Set GetOrCreateGroup = group
End Function

Private Sub AddGroupPermissionsToList(group As Collection, lbx As MSForms.ListBox)
    Dim itemColl As Variant, perm As clsPermissionNode
    
    For Each itemColl In group
        Set perm = itemColl
        
        lbx.AddItem
        lbx.List(lbx.listCount - 1, 0) = ""
        lbx.List(lbx.listCount - 1, 1) = perm.GroupKey
        lbx.List(lbx.listCount - 1, 2) = CStr(perm.PermissionLevel)
        lbx.List(lbx.listCount - 1, 3) = perm.permission
    Next itemColl
    
End Sub

Private Sub ResetPermissionsList(lbx As MSForms.ListBox)
    lbx.Clear
    lbx.ColumnCount = 4
    lbx.ColumnWidths = "0;0;0;100" ' Id,Group,Level,Permission
    
End Sub


'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************


Private Sub Class_Initialize()
    Set m_oGroups = New Collection
End Sub

