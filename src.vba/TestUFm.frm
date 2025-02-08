Option Explicit

Dim mainaccessdb As CDataSourceDatabase
Dim dataacess As CDataAccess

Private Sub UserForm_Initialize()
    
    Set mainaccessdb = NewDatabase(GetLoggedUser, "D:\WorkPersoData\PersoData\J2CEY_DATA\WRK\GT\BUSINESS_APPS\CAISSEVIRTUELLECOMILOG\app\db", "comilogcashdb", access2007)
    Set dataacess = NewDataAccess(GetLoggedUser, mainaccessdb, "users")
    
    dataacess.Record.FieldList.AddField NewField(NewFieldValueString(), "userlogin", "Login").SetSelectable(True)
    dataacess.Record.FieldList.AddField NewField(NewFieldValueString(), "username", "User Name").SetSelectable(True)
    
    Dim result As CResult, oRec As CRecord, oRecList As CRecordList
    
    'Set result = dataacess.GetValue("username", True)
    'Set oRec = dataacess.GetRecord(result,True)
    Set oRecList = dataacess.GetRecordList(result, True)
    
    GetMainDataSource
    
    MsgBox Now_System() & ", MS: " & GetTodayMilliseconds() & ", CreateGUID: " & CreateGUID()
    
End Sub

Private Sub UserForm_Terminate()
    mainaccessdb.CloseDatabase
End Sub