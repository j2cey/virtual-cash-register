Public WithEvents btnLabel          As MSForms.Label
Public WithEvents btnImage          As MSForms.Image
Public WithEvents btnCheckBox       As MSForms.CheckBox
Public WithEvents btnButton         As MSForms.CommandButton

Private Sub btnButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call LittleHand
End Sub

Private Sub btnCheckBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call LittleHand
End Sub

Private Sub btnImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call LittleHand
End Sub

Private Sub btnLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call LittleHand
End Sub