Option Explicit

Dim mainaccessdb As CDatabaseDS

Private Sub UserForm_Initialize()
    
    Set mainaccessdb = NewDatabase(NewUser(), "D:\WorkPersoData\PersoData\J2CEY_DATA\WRK\GT\BUSINESS_APPS\CAISSEVIRTUELLECOMILOG\app\db", "comilogcashdb", NewAccess2007())
    
    mainaccessdb.OpenDatabase
    
    MsgBox Now_System() & ", MS: " & GetTodayMilliseconds() & ", CreateGUID: " & CreateGUID()
    
End Sub

Private Sub UserForm_Terminate()
    mainaccessdb.CloseDatabase
End Sub