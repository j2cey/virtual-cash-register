Option Explicit

Private Sub UserForm_Initialize()
    Dim accessdb As CDatabaseDS
    
    Set accessdb = NewDatabase(NewUser(), "Server Or Path", "DB Name", NewAccess2007())
    
    MsgBox Now_System() & ", MS: " & GetTodayMilliseconds() & ", CreateGUID: " & CreateGUID()
    
End Sub