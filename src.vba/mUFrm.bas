Option Explicit


Public Function SetVisibility(ctl As MSForms.Control, bVal As Boolean, Optional arrPermissions As Variant) As Boolean
    If Not IsMissing(arrPermissions) And (Not IsNull(arrPermissions)) And (Not IsEmpty(arrPermissions)) Then
        If Not oLoggedUser.Can(arrPermissions) Then
            SetVisibility = False
            Exit Function
        End If
    End If
    
    SetVisibility = True
    ctl.Visible = bVal
End Function