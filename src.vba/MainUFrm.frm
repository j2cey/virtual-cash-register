
Option Explicit

Public WithEvents oUserSaveForm As clsSaveForm
Dim pAllowClose As Boolean

Private Sub EmergencyExitFrm_Click()
    Dim mrgcy_pwd As String
    
    'mrgcy_pwd = InputBox("Emergency PWD for Code Access ?", GetAppName, "")
    mrgcy_pwd = InputBoxDK("Emergency PWD for Code Access ?", GetAppName)
    
    If mrgcy_pwd = "ComiCash123" Then
        Unload Me
        Application.Visible = True
    End If
End Sub

Private Sub UserForm_Initialize()

    Call InitGlobal
    
    Call InitMainUFrm(Me)
    
    Call InitEmployees(Me)
    
    Set oUserSaveForm = NewSaveForm("users", Me.UserSaveTitleLbl, Me.UserSaveImg, Me.UserCancelImg, Me.UserNewImg)
    Call InitUsers(Me, oUserSaveForm)
    Call InitRoles(Me)
    
    Call InitPayments(Me)
    
    Call SetVisibility(Me.MenuAudit, True, Array("audit-lister"))
    
    'Remove Border and Title Bar
    'oUfrm.HideBar
    'oUfrm.SetNoTitleBar
    
    Call RemoveAll(Me)
    Call CursorHand(Me)
    
    Me.MainMultiPage.Style = fmTabStyleNone
    
    Application.Visible = False
End Sub

Private Sub MenuDashboard_Click()
    InitDashbord Me
    SetActivePage Me, 0, "Dashboard"
End Sub
Private Sub MenuDashboardLabel_Click()
    InitDashbord Me
    SetActivePage Me, 0, "Dashboard"
End Sub

Private Sub MenuEmployeesLabel_Click()
    AccessMenuEmployees Me
End Sub
Private Sub MenuEmployees_Click()
    AccessMenuEmployees Me
End Sub

Private Sub MenuPaymentsLabel_Click()
    AccessMenuPayments Me
End Sub
Private Sub MenuPayments_Click()
    AccessMenuPayments Me
End Sub

Private Sub MenuUsersLabel_Click()
    AccessMenuUsers Me
End Sub
Private Sub MenuUsers_Click()
    AccessMenuUsers Me
End Sub

Private Sub MenuSettingsLabel_Click()
    AccessMenuSettings Me
End Sub

Private Sub MenuSettings_Click()
    AccessMenuSettings Me
End Sub

Private Sub MenuSettingsImg_Click()
    AccessMenuSettings Me
End Sub

Private Sub MenuLogoutLabel_Click()
    Logout
End Sub
Private Sub MenuLogout_Click()
    Logout
End Sub

Private Sub MenuAudit_Click()
    AccessMenuAudit Me
End Sub

Private Sub MenuAuditImg_Click()
    AccessMenuAudit Me
End Sub

Private Sub MenuAuditLbl_Click()
    AccessMenuAudit Me
End Sub

Private Sub Logout()
    loggedUser.Logout
    Unload Me
    CloseAndSaveOpenWorkbooks
End Sub

'
'   EMPLOYEES
'
Private Sub EmployeeNewImg_Click()
    EmployeesMultiPage.Value = 1
End Sub

Private Sub EmployeeSearchImg_Click()
    EmployeesMultiPage.Value = 0
End Sub

Private Sub EmployeesListLBx_Click()
    SelectEmployee Me
End Sub

Private Sub EmployeeSaveImg_Click()
    SaveEmployee Me
End Sub

Private Sub EmployeeEditImg_Click()
    EditEmployee Me
End Sub

Private Sub EmployeeDeleteImg_Click()
    DeleteEmployee Me
End Sub

Private Sub EmployeeCancelImg_Click()
    EmployeesMultiPage.Value = 0
End Sub

Private Sub ManageEmployeeSiteImg_Click()
    ManageEmployeeSite Me
End Sub

Private Sub ManageEmployeeDirectionImg_Click()
    ManageEmployeeDirection Me
End Sub

Private Sub ManageEmployeeDepartementImg_Click()
    ManageEmployeeDepartement Me
End Sub

Private Sub ManageEmployeePosteImg_Click()
    ManageEmployeePoste Me
End Sub

'
'   USERS
'
Private Sub UserNewImg_Click()
    FormatUserSaveForm Me, oUserSaveForm, True
    
    UsersMultiPage.Value = 1
End Sub

Private Sub UserSaveImg_Click()
    SaveUser Me, oUserSaveForm
End Sub

Private Sub UserSearchCancelImg_Click()
    UnselectUser Me
End Sub

Private Sub UserSearchImg_Click()
    UnselectUser Me
    UsersMultiPage.Value = 0
End Sub

Private Sub UsersListLBx_Click()
    SelectUser Me
End Sub

Private Sub UserEditImg_Click()
    FormatUserSaveForm Me, oUserSaveForm, False
    
    EditUser Me, oUserSaveForm
End Sub

Private Sub UserDeleteImg_Click()
    DeleteUser Me, oUserSaveForm
End Sub

Private Sub UserCancelImg_Click()
    UsersMultiPage.Value = 0
End Sub

Private Sub UserEditPwdImg_Click()
    EditUserPwd Me
End Sub

Private Sub oUserSaveForm_FormSaved()
    
End Sub

Private Sub LoggedUserEditPwdLbl_Click()
    EditLoggedUserPwd Me
End Sub

'
'   USER ROLES
'
Private Sub UserrolesImg_Click()
    UsersMultiPage.Value = 2
End Sub

Private Sub UserroleSearchImg_Click()
    UnSelectRole Me
End Sub

Private Sub UserroleSearchCancelImg_Click()
    UnSelectRole Me
End Sub

Private Sub UserrolesListLBx_Click()
    SelectRole Me
End Sub

Private Sub RolePermissionsLBx_Click()
    SelectRolePermission Me
End Sub

Private Sub PermissionsLBx_Click()
    SelectPermission Me
End Sub

Private Sub AddPermissionToRoleImg_Click()
    AddPermissionToRole Me
End Sub

Private Sub AddAllPermissionsToRoleImg_Click()
    AddAllPermissionsToRole Me
End Sub

Private Sub RemovePermissionFromRoleImg_Click()
    RemovePermissionToRole Me
End Sub

Private Sub RemoveAllPermissionsFromRoleImg_Click()
    RemoveAllPermissionsToRole Me
End Sub

Private Sub UserroleSaveImg_Click()
    SaveRole Me
End Sub

Private Sub UserroleCancelImg_Click()
    UnSelectRole Me
End Sub

Private Sub UserroleDeleteImg_Click()
    DeleteRole Me
End Sub

'
'   PAYMENTS
'
Private Sub PaymentNewImg_Click()
    LaunchCreatePayment Me
End Sub

Private Sub SearchPaymentCreateAtImg_Click()
    Me.SearchPaymentCreateAtTBx.Text = CalendarForm.GetDate
End Sub

Private Sub PaymentSearchEmployeeListLBx_Click()
    'SelectPaymentEmployee Me
End Sub

Private Sub PaymentSearchEmployeeListLBx_Change()
    SelectPaymentEmployee Me
End Sub

Private Sub PaymentSearchEmployeeListLBx_Enter()
    'SelectPaymentEmployee Me
End Sub

Private Sub PaymentAddImg_Click()
    AddPaymentDetail Me
End Sub

Private Sub PaymentSearchEmployeeImg_Click()
    ClearPaymentEmployeeDetails Me
End Sub

Private Sub PaymentSearchEmployeeCancelImg_Click()
    ClearPaymentEmployeeDetails Me
End Sub

Private Sub PaymentResetImg_Click()
    PaymentReset Me
End Sub

Private Sub PaymentDetailsLBx_Click()
    SelectPaymentDetail Me
End Sub

Private Sub PaymentRemoveImg_Click()
    RemovePaymentDetail Me
End Sub

Private Sub PaymentSaveImg_Click()
    SavePayment Me
End Sub

Private Sub PaymentListLBx_Click()
    SelectPayment Me
End Sub

Private Sub PaymentSearchImg_Click()
    SwitchToPaymentSearch Me
End Sub

Private Sub PaymentSearchCancelImg_Click()
    SwitchToPaymentSearch Me
End Sub

Private Sub PaymentAmountTBx_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyReturn Then
       AddPaymentDetail Me
    End If
End Sub

Private Sub PaymentTitleTBx_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyReturn Then
       SavePayment Me
    End If
End Sub

Private Sub PaymentEditImg_Click()
    EditPayment Me
End Sub

Private Sub PaymentShowImg_Click()
    ShowPayment Me
End Sub

Private Sub PaymentValidateImg_Click()
    ValidatePayment Me
End Sub

Private Sub PaymentExtractImg_Click()
    ExtractPayment Me
End Sub

Private Sub PaymentSendFileImg_Click()
    SendPayment Me
End Sub

Private Sub PaymentExecuteImg_Click()
    ExecutePayment Me
End Sub

Private Sub PaymentDeleteImg_Click()
    DeletePayment Me
End Sub

Private Sub ManagePaymentMotifImg_Click()
    ManagePaymentMotif Me
End Sub

'
'   SETTINGS
'
Private Sub SettingSystParamSaveCBtn_Click()
    SaveSystParamSettings Me
End Sub

Private Sub SettingDbAccessParamSaveCBtn_Click()
    SaveDbAccessParamSettings Me
End Sub

Private Sub SettingCodeAccessParamSaveCBtn_Click()
    SaveCodeAccessParamSettings Me
End Sub

Private Sub SettingPaymentSaveCBtn_Click()
    SaveSettingPayment Me
End Sub

Private Sub MotifSpecialCharsLBx_Click()
    SelectMotifSpecialChar Me
End Sub

Private Sub MotifSpecialCharCancelImg_Click()
    CancelMotifSpecialChar Me
End Sub

Private Sub MotifSpecialCharSaveImg_Click()
    SaveMotifSpecialChar Me
End Sub

Private Sub MotifSpecialCharDeleteImg_Click()
    DeleteMotifSpecialChar Me
End Sub

Private Sub PwdLegnthMinSCb_Change()
    PwdLegnthMinTBx.Text = PwdLegnthMinSCb.Value
End Sub

Private Sub PwdUppercaseMinSCb_Change()
    PwdUppercaseMinTBx.Text = PwdUppercaseMinSCb.Value
End Sub

Private Sub PwdNumberMinSCb_Change()
    PwdNumberMinTBx.Text = PwdNumberMinSCb.Value
End Sub

Private Sub PwdSpecialCharsMinSCb_Change()
    PwdSpecialCharsMinTBx.Text = PwdSpecialCharsMinSCb.Value
End Sub

Private Sub SettingSecuritySaveCBtn_Click()
    SaveSettingSecurity Me
End Sub

Private Sub SettingDbSelectParamSaveCBtn_Click()
    SaveDbSelectParamSettings Me
End Sub

'
' MAIL SETTINGS
'
Private Sub SettingMailProviderCBx_Change()
    SelectMailProvider Me
End Sub

Private Sub SettingMailProviderParametersLBx_Click()
    SelectMailProviderParameter Me
End Sub

Private Sub SettingMailProviderParameterSelectedCancelImg_Click()
    CancelMailProviderParameter Me
End Sub

Private Sub SettingMailProviderParameterSelectedSaveImg_Click()
    SaveMailProviderParameter Me
End Sub

Private Sub SettingMailSaveCBtn_Click()
    SaveSettingMail Me
End Sub

'
' Payment Extraction Save / Send File
'
Private Sub PymntExtractionDontSendFileOBtn_Click()
    ChangePymntExtractionSendFile Me
End Sub

Private Sub PymntExtractionSendFileToListOBtn_Click()
    ChangePymntExtractionSendFile Me
End Sub

Private Sub PymntExtractionSendFileToUsersOBtn_Click()
    ChangePymntExtractionSendFile Me
End Sub

Private Sub PymntExtractionSendFileToAllOBtn_Click()
    ChangePymntExtractionSendFile Me
End Sub

Private Sub PymntExtractionReceiversListLBx_Click()
    SelectPymntExtractionReceiver Me
End Sub

Private Sub PymntExtractionReceiverSelCancelImg_Click()
    CancelPymntExtractionReceiver Me
End Sub

Private Sub PymntExtractionReceiverSelSaveImg_Click()
    SavePymntExtractionReceiver Me
End Sub

Private Sub PymntExtractionReceiverSelDeleteImg_Click()
    DeletePymntExtractionReceiver Me
End Sub

Private Sub SettingPaymentSendFileSaveCBtn_Click()
    SaveSettingPaymentSendFile Me
End Sub

Private Sub PymntExtractionFileFolderImg_Click()
    PymntExtractionFileFolderTBx.Text = GetFolder
End Sub