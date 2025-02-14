
Private Sub UserForm_Initialize()
    InitManageSubList Me

    'Call RemoveAll(Me)
    Call CursorHand(Me)
End Sub

Private Sub ItemSearchImg_Click()
    SearchItem Me
End Sub

Private Sub ItemSearchCancelImg_Click()
    ResetItemForms Me, Add
End Sub

Private Sub ItemsListLBx_Click()
    SelectItem Me
End Sub

Private Sub ItemCancelImg_Click()
    ResetItemForms Me, Add
End Sub

Private Sub ItemSaveImg_Click()
    SaveItem Me
End Sub

Private Sub ItemDeleteImg_Click()
    DeleteItem Me
End Sub

Private Sub CloseImg_Click()
    Unload Me
End Sub