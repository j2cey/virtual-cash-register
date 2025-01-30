Option Explicit

Dim mainaccessdb As CDatabaseDS
Dim dataacess As CRecordableDA
Dim currentUser As CUser

Private Sub UserForm_Initialize()
    
    Set currentUser = NewUser()
    Set mainaccessdb = NewDatabase(currentUser, "D:\WorkPersoData\PersoData\J2CEY_DATA\WRK\GT\BUSINESS_APPS\CAISSEVIRTUELLECOMILOG\app\db", "comilogcashdb", NewAccess2007())
    Set dataacess = NewRecordableDA(currentUser, mainaccessdb, "users")
    
    dataacess.List
    
    MsgBox Now_System() & ", MS: " & GetTodayMilliseconds() & ", CreateGUID: " & CreateGUID()
    
End Sub

Private Sub UserForm_Terminate()
    mainaccessdb.CloseDatabase
End Sub