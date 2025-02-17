Option Explicit


Private Sub UserForm_Initialize()
    
    Dim userListBL As CBusinessLogic, oUser As CModelUser
    
    Set userListBL = NewBusinessLogic(NewDataAccess(GetLoggedUser(), GetMainDataSource, "users", "users_view"), GetLoggedUser())
    
    Set oUser = NewUserFromBD(1)
    oUser.LoadValues
    
    userListBL.AddField NewFieldValueString(), "userlogin", "User Login"
    userListBL.AddField NewFieldValueString(), "username", "User Name"
    userListBL.AddField NewFieldValueString(), "usermail", "User E-Mail"
    
End Sub

Private Sub UserForm_Terminate()
    
End Sub