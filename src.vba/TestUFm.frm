Option Explicit

Dim mainaccessdb As CDatabaseDS
Dim dataacess As CRecordableDA
Dim currentUser As CUser

Private Sub UserForm_Initialize()
    
    Set currentUser = NewUser()
    Set mainaccessdb = NewDatabase(currentUser, "D:\WorkPersoData\PersoData\J2CEY_DATA\WRK\GT\BUSINESS_APPS\CAISSEVIRTUELLECOMILOG\app\db", "comilogcashdb", NewAccess2007())
    Set dataacess = NewRecordable(currentUser, mainaccessdb, "users")
    
    dataacess.FieldList.AddField NewField(NewFieldValueString(), "userlogin", "Login")
    
    Dim oLoginField As CField
    
    Set oLoginField = dataacess.GetValue("userlogin", True)
    
    MsgBox Now_System() & ", MS: " & GetTodayMilliseconds() & ", CreateGUID: " & CreateGUID()
    
End Sub

Private Sub UserForm_Terminate()
    mainaccessdb.CloseDatabase
End Sub