

Private Sub UserForm_Initialize()
    InitAudit Me

    'Call RemoveAll(Me)
    Call CursorHand(Me)
End Sub

Private Sub SearchPaymentCreateAtImg_Click()
    Me.SearchStartedAtTBx.Text = CalendarForm.GetDate
End Sub

Private Sub AuditSearchImg_Click()
    SearchAudit Me
End Sub

Private Sub AuditSearchCancelImg_Click()
    ResetAuditForms Me
End Sub

Private Sub AuditListLBx_Click()
    SelectAudit Me
End Sub

Private Sub CloseImg_Click()
    Unload Me
End Sub