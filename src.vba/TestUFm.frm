Option Explicit


Private Sub UserForm_Initialize()
    
    Dim userListBL As CBusinessLogic
    
    Set userListBL = NewBusinessLogic(GetLoggedUser(), NewDataAccess(GetLoggedUser(), GetMainDataSource, "users_view"))
    
    userListBL.AddField NewFieldValueString(), "userlogin", "User Login"
    userListBL.AddField NewFieldValueString(), "username", "User Name"
    userListBL.AddField NewFieldValueString(), "usermail", "User E-Mail"
    
End Sub

Private Sub UserForm_Terminate()
    
End Sub