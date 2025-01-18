'Build 000
'*****************************************************************************************************************************
'
' Authors:  Jude Parfait NGOM NZE, jud10parfait@gmail.com
'
' (c)2024, all rights reserved to the authors
'
'*****************************************************************************************************************************
'-------------------------------------------------------------------------
' Module    : CPermissionNode
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 11-06-2024
' Purpose   : Manage A Permission Node
'-------------------------------------------------------------------------
Option Explicit


Private m_sGroupKey As String                   ' the permission group key
Private m_sPermission As String                 ' the permission name
Private m_iPermissionLevel As Integer           ' the permission level in the group



'*****************************************************************************************************************************
'    Public events
'*****************************************************************************************************************************



'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CPermissionNode
    Set Self = Me
End Property

' GroupKey
Public Property Let GroupKey(sGroupKey As String)
    m_sGroupKey = sGroupKey
End Property

Public Property Get GroupKey() As String
    GroupKey = m_sGroupKey
End Property

' Permission
Public Property Let permission(sPermission As String)
    m_sPermission = sPermission
End Property

Public Property Get permission() As String
    permission = m_sPermission
End Property

' PermissionLevel
Public Property Let PermissionLevel(iPermissionLevel As Integer)
    m_iPermissionLevel = iPermissionLevel
End Property

Public Property Get PermissionLevel() As Integer
    PermissionLevel = m_iPermissionLevel
End Property



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(sGroupKey As String, sPermission As String, iPermissionLevel As Integer)
    Me.GroupKey = sGroupKey
    Me.permission = sPermission
    Me.PermissionLevel = iPermissionLevel
End Sub


'*****************************************************************************************************************************
'    Public functions and subs




'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************


Private Sub Class_Initialize()
    
End Sub
