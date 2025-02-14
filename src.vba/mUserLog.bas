Option Explicit

Public Function UserIsLogged(roles() As String) As Boolean
    Dim i As Integer
    
    If oLoggedUser Is Nothing Then
        UserIsLogged = False
        Exit Function
    End If
    
    If oLoggedUser.Id < 1 Then
        UserIsLogged = False
        Exit Function
    End If
    
    If Not IsArrayStringInitialised(roles) Then
        UserIsLogged = True
        Exit Function
    End If
    
    If oLoggedUser.Role Is Nothing Then
        UserIsLogged = False
        Exit Function
    End If
    
    For i = 0 To UBound(roles)
        If oLoggedUser.Role.Id = CLng(roles(i)) Then
            UserIsLogged = True
            Exit Function
        End If
    Next i
    
    UserIsLogged = False
End Function

Public Sub TestVariantArray(strName As Variant)
    Dim i As Integer
    
    For i = 0 To UBound(strName)
        MsgBox strName(i)
    Next i
End Sub