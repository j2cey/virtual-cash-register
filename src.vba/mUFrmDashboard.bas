Option Explicit


Public Sub InitDashbord(uFrm As MSForms.UserForm)
    Dim dashboardProgr As clsProgression
    
    Set dashboardProgr = StartNewProgression("Chargement Dashboard", 1)
    
    dashboardProgr.StartNewSubProgression "Chargement Statistiques Employés", 1
    uFrm.StatEmployeesCountLbl.Caption = CStr(GetEmployeesCount)
    dashboardProgr.AddDoneLastSub 1, True
    
    dashboardProgr.StartNewSubProgression "Chargement Paiements en Attente Validation", 1
    uFrm.StatPaymentWaitingValidationLbl.Caption = CStr(GetPaymentWaitingValidationCount)
    dashboardProgr.AddDoneLastSub 1, True
    
    dashboardProgr.StartNewSubProgression "Chargement Paiements en Attente Extraction", 1
    uFrm.StatPaymentWaitingExtractionLbl.Caption = CStr(GetPaymentWaitingExtractionCount)
    dashboardProgr.AddDoneLastSub 1, True
    
    dashboardProgr.StartNewSubProgression "Chargement Statistiques Utilisateurs", 1
    uFrm.StatUsersCountLbl.Caption = CStr(GetUsersCount)
    dashboardProgr.AddDoneLastSub 1, True
    
    dashboardProgr.StartNewSubProgression "Chargement Liste Derniers Paiements", 1
    GetPaymentStatList uFrm
    dashboardProgr.AddDoneLastSub 1, True
    
    dashboardProgr.AddDone 1, True
    
End Sub

Private Function GetEmployeesCount() As Long
    Dim employeesData As Variant, sqlRst As String, recordCount As Long, i As Long

    sqlRst = "SELECT COUNT(*) FROM employees"
    
    Call PrepareDatabase
    employeesData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    GetEmployeesCount = CLng(employeesData(0, 0))
End Function

Private Function GetUsersCount() As Long
    Dim usersData As Variant, sqlRst As String, recordCount As Long, i As Long

    sqlRst = "SELECT COUNT(*) FROM users"
    
    Call PrepareDatabase
    usersData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    GetUsersCount = CLng(usersData(0, 0))
End Function

Private Function GetPaymentWaitingValidationCount() As Long
    Dim paymentsData As Variant, sqlRst As String, recordCount As Long, i As Long

    sqlRst = "SELECT COUNT(*) FROM payments_view WHERE status_id = " & CStr(SettPymntStatusAttenteValidationId.Val) & ""
    
    Call PrepareDatabase
    paymentsData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    GetPaymentWaitingValidationCount = CLng(paymentsData(0, 0))
End Function

Private Function GetPaymentWaitingExtractionCount() As Long
    Dim paymentsData As Variant, sqlRst As String, recordCount As Long, i As Long

    sqlRst = "SELECT COUNT(*) FROM payments_view WHERE status_id = " & CStr(SettPymntStatusAttenteExtractionId.Val) & ""
    
    Call PrepareDatabase
    paymentsData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    GetPaymentWaitingExtractionCount = CLng(paymentsData(0, 0))
End Function

Public Sub GetPaymentStatList(uFrm As MSForms.UserForm)
    Dim paymentData As Variant, sqlRst As String, recordCount As Long, i As Long
    
    ResetPaymentStatList uFrm
    
    sqlRst = SqlSelectSetLIMIT("SELECT Id,title,created_at,amount,status_title FROM payments_view", CLng(SettPymntLastLinesSizeDashboard.Val))
    
    Call PrepareDatabase
    paymentData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    If recordCount = 0 Then
        Exit Sub
    End If
    
    For i = 0 To recordCount - 1
        AddToPaymentStatList uFrm, CStr(paymentData(0, i)), CStr(paymentData(1, i)), CStr(paymentData(2, i)), CStr(paymentData(3, i)), CStr(paymentData(4, i))
    Next i
End Sub

Private Sub AddToPaymentStatList(uFrm As MSForms.UserForm, sId, sTitre, sDate, sMontant, sStatut As String)
    Dim lastAddedIdx As Long, i As Long, answer As Integer, totalAmount As Long
    
    uFrm.PaymentStatListLBx.AddItem
    
    lastAddedIdx = uFrm.PaymentStatListLBx.listCount - 1
    
    uFrm.PaymentStatListLBx.List(lastAddedIdx, 0) = sId
    uFrm.PaymentStatListLBx.List(lastAddedIdx, 1) = sTitre
    uFrm.PaymentStatListLBx.List(lastAddedIdx, 2) = sDate
    uFrm.PaymentStatListLBx.List(lastAddedIdx, 3) = sMontant
    uFrm.PaymentStatListLBx.List(lastAddedIdx, 4) = sStatut
End Sub

Private Sub ResetPaymentStatList(uFrm As MSForms.UserForm)
    Dim lbx As MSForms.ListBox, lColumnCount As Long, sColumnWidths As String
    
    lColumnCount = 5
    sColumnWidths = "0;100;100;70;100" 'Id,Titre,Date,Montant,Statut
    
    uFrm.PaymentStatListLBx.Locked = True
    uFrm.PaymentStatListLBx.Clear
    uFrm.PaymentStatListLBx.ColumnCount = lColumnCount
    uFrm.PaymentStatListLBx.ColumnWidths = sColumnWidths
    
End Sub