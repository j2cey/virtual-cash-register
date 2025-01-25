Option Explicit

Private Sub UserForm_Initialize()
    Dim rec As CRecord
    
    Set rec = NewRecord()
    
    MsgBox Now_System() & ", MS: " & GetTodayMilliseconds() & ", CreateGUID: " & CreateGUID()
    
End Sub