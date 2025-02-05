Option Explicit

Dim mainaccessdb As CDatabaseDS
Dim dataacess As CRecordableDA
Dim currentUser As CUser

Private Sub UserForm_Initialize()
    
    Set currentUser = NewUser()
    Set mainaccessdb = NewDatabase(currentUser, "D:\WorkPersoData\PersoData\J2CEY_DATA\WRK\GT\BUSINESS_APPS\CAISSEVIRTUELLECOMILOG\app\db", "comilogcashdb", NewAccess2007())
    Set dataacess = NewRecordable(currentUser, mainaccessdb, "users")
    
    dataacess.Record.FieldList.AddField NewField(NewFieldValueString(), "userlogin", "Login").SetSelectable(True)
    dataacess.Record.FieldList.AddField NewField(NewFieldValueString(), "username", "User Name").SetSelectable(True)
    
    Dim result As CResult, oRec As CRecord, oRecList As CRecordList
    
    'Set result = dataacess.GetValue("username", True)
    'Set oRec = dataacess.GetRecord(True)
    Set oRecList = dataacess.GetRecordList(True)
    
    MsgBox Now_System() & ", MS: " & GetTodayMilliseconds() & ", CreateGUID: " & CreateGUID()
    
End Sub

Private Sub UserForm_Terminate()
    mainaccessdb.CloseDatabase
End Sub