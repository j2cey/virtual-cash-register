'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Module    : mInitializeUsers
'   Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
'   Created   : 2025/02/05
'   Purpose   : Manage All (logged) Users related variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'***************************************************************************************************************************************************************
'   Public Variables
'***************************************************************************************************************************************************************

Public Const USERS_TABLE_FOR_SAVING_DEFAULT As String = "users"
Public Const USERS_TABLE_FOR_SELECTING_DEFAULT As String = "users_view"

Public oLoggedUser As CModelUser



'***************************************************************************************************************************************************************
'   Public Functions & Subroutines
'***************************************************************************************************************************************************************

Public Function GetLoggedUser() As CModelUser
    If oLoggedUser Is Nothing Then
        Set oLoggedUser = NewUser()
    Else
        Set GetLoggedUser = oLoggedUser
    End If
End Function

Public Function TryLogInLocalUser() As Boolean
    
End Function


