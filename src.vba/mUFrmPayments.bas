Public paymentBodyFrm As CUIControl, oPaymentSearchFrm As CUIControl, oPaymentSearchEmployeeMatriculeTBx As CUIControl, oPaymentSaveFrm As CUIControl
Public oPaymentSearchEmployeeForm As clsSearchForm, oPaymentSearchEmployeeFrm As CUIControl
Public sPaymentEmployeeSelectedId As String, lPaymentDetailSelectedId As Long
Public oPaymentSearchForm As clsSearchForm, oPaymentListFrm As CUIControl, sPaymentSelectedId As String, oPaymentSelected As clsPayment
Public oPaymentSaveForm As clsSaveForm
Public paymentEmployeeSelectedCount As Long
Public oPaymentDetailFrm As CUIControl

Public Sub InitPayments(uFrm As MSForms.UserForm)
    uFrm.PaymentsMultiPage.Value = 0
    
    ' Main Body
    Set paymentBodyFrm = NewUCTL(uFrm.PaymentsBodyCardFrm, oMainUfrm)
    paymentBodyFrm.AddCtl uFrm.PaymentNewImg
    
    ' Search Frame
    Set oPaymentSearchFrm = NewUCTL(uFrm.PaymentSearchFrm, oMainUfrm)
    oPaymentSearchFrm.AddCtl uFrm.PaymentSearchImg
    oPaymentSearchFrm.AddCtl uFrm.PaymentSearchCancelImg
    
    ' List Page
    Set oPaymentListFrm = NewUCTL(uFrm.PaymentListFrm, oMainUfrm)
    oPaymentListFrm.AddCtl uFrm.PaymentEditImg
    oPaymentListFrm.AddCtl uFrm.PaymentDeleteImg
    oPaymentListFrm.AddCtl uFrm.PaymentShowImg
    
    oPaymentListFrm.AddCtl uFrm.PaymentValidateImg
    oPaymentListFrm.AddCtl uFrm.PaymentExtractImg
    oPaymentListFrm.AddCtl uFrm.PaymentExecuteImg
    oPaymentListFrm.AddCtl uFrm.PaymentSendFileImg
    
    ' Search Employees Frame
    Set oPaymentSearchEmployeeFrm = NewUCTL(uFrm.PaymentSearchEmployeeFrm, oMainUfrm)
    oPaymentSearchEmployeeFrm.AddCtl uFrm.PaymentSearchEmployeeImg
    oPaymentSearchEmployeeFrm.AddCtl uFrm.PaymentSearchEmployeeCancelImg
    
    Set oPaymentDetailFrm = NewUCTL(uFrm.PaymentDetailFrm, oMainUfrm)
    oPaymentDetailFrm.AddCtl uFrm.PaymentAddImg
    Call oPaymentDetailFrm.AddCtl(uFrm.ManagePaymentMotifImg).SetSizes(NewSize(14, 14), NewSize(16, 16))
    
    ' Save Page
    Set oPaymentSaveFrm = NewUCTL(uFrm.PaymentSaveFrm, oMainUfrm)
    oPaymentSaveFrm.AddCtl uFrm.PaymentRemoveImg
    oPaymentSaveFrm.AddCtl uFrm.PaymentSaveImg
    oPaymentSaveFrm.AddCtl uFrm.PaymentResetImg
    
    ' PaymentSaveTitleLbl
    
    Call InitSearchForm(uFrm)
    Call InitSearchEmployeeForm(uFrm)
    Call InitSaveForm(uFrm)
    
    Call InitPaymentMotifsCbx(uFrm)
    
    Call ResetPaymentDetails(uFrm)
    Call InitPaymentStatus(uFrm)
    
    Call InitFormat(uFrm)
End Sub

Public Sub InitPaymentsAccess(uFrm As MSForms.UserForm)
    SetVisibility uFrm.PaymentNewImg, True, Array("paiement-ajouter")
    SetVisibility uFrm.PaymentNewLbl, True, Array("paiement-ajouter")
    
    SetVisibility uFrm.PaymentSaveImg, True, Array("paiement-ajouter", "paiement-modifer")
    SetVisibility uFrm.PaymentSaveLbl, True, Array("paiement-ajouter", "paiement-modifer")
    
    SetVisibility uFrm.ManagePaymentMotifImg, True, Array("motif_paiement-lister", "motif_paiement-ajouter", "motif_paiement-modifier", "motif_paiement-supprimer")
End Sub

Private Sub InitFormat(uFrm As MSForms.UserForm)
    uFrm.SearchPaymentCreateAtTBx.BackColor = &H80000005
    uFrm.SearchPaymentCreateAtTBx.Locked = True
    
    uFrm.PaymentTotalAmountTBx.BackColor = &HC0C0C0
    uFrm.PaymentTotalAmountTBx.Locked = True
    
    uFrm.PaymentCreatedAtTBx.BackColor = &HC0C0C0
    uFrm.PaymentCreatedAtTBx.Locked = True
    SetCreatedAtFormat uFrm, False
    
    InitPaymentListFormat uFrm
End Sub

Private Sub SetCreatedAtFormat(uFrm As MSForms.UserForm, bValue As Boolean)
    uFrm.PaymentCreatedAtTBx.Visible = bValue
    uFrm.PaymentCreatedAtLbl.Visible = bValue
End Sub

Private Sub SetPaymentValidateFormat(uFrm As MSForms.UserForm, bFormat As Boolean, Optional arrPermissions As Variant)
    If bFormat Then
        SetVisibility uFrm.PaymentValidateImg, bFormat, arrPermissions
        SetVisibility uFrm.PaymentValidateLbl, bFormat, arrPermissions
        
        SetVisibility uFrm.PaymentEditImg, bFormat, Array("paiement-modifer")
        SetVisibility uFrm.PaymentDeleteImg, bFormat, Array("paiement-supprimer")
        SetVisibility uFrm.PaymentShowImg, bFormat, Array("paiement-voir_details")
    Else
        SetVisibility uFrm.PaymentValidateImg, bFormat
        SetVisibility uFrm.PaymentValidateLbl, bFormat
        
        SetVisibility uFrm.PaymentEditImg, bFormat
        SetVisibility uFrm.PaymentDeleteImg, bFormat
        SetVisibility uFrm.PaymentShowImg, bFormat
    End If
End Sub

Private Sub SetPaymentExtractFormat(uFrm As MSForms.UserForm, bFormat As Boolean)
    Dim arrPerms As Variant
    
    If bFormat Then
        arrPerms = Array("paiement-extraire")
    Else
        arrPerms = Empty
    End If
    
    SetVisibility uFrm.PaymentExtractImg, bFormat, arrPerms
    SetVisibility uFrm.PaymentExtractLbl, bFormat, arrPerms
End Sub

Private Sub SetPaymentExecuteFormat(uFrm As MSForms.UserForm, bFormat As Boolean)
    Dim arrPerms As Variant
    
    If bFormat Then
        arrPerms = Array("paiement-marquer_execute")
    Else
        arrPerms = Empty
    End If
    
    SetVisibility uFrm.PaymentExecuteImg, bFormat, arrPerms
    SetVisibility uFrm.PaymentExecuteLbl, bFormat, arrPerms
End Sub

Private Sub SetPaymentSentFormat(uFrm As MSForms.UserForm, bFormat As Boolean)
    Dim arrPerms As Variant
    
    If bFormat Then
        arrPerms = Array("paiement-envoyer_fichier")
    Else
        arrPerms = Empty
    End If
    
    SetVisibility uFrm.PaymentSendFileImg, bFormat, arrPerms
    SetVisibility uFrm.PaymentSendFileLbl, bFormat, arrPerms
End Sub

Private Sub InitPaymentListFormat(uFrm As MSForms.UserForm)
    SetPaymentValidateFormat uFrm, False
    SetPaymentExtractFormat uFrm, False
    SetPaymentExecuteFormat uFrm, False
    SetPaymentSentFormat uFrm, False
End Sub

Public Sub SwitchToPaymentSearch(uFrm As MSForms.UserForm)
    InitPaymentListFormat uFrm
    uFrm.PaymentsMultiPage.Value = 0
End Sub

Public Sub SelectPayment(uFrm As MSForms.UserForm)
    Dim i As Long
    
    InitPaymentListFormat uFrm
    'PaymentSelected = ClearArray(PaymentSelected)
    
    With uFrm.PaymentListLBx
    For i = 0 To .listCount - 1
      If .Selected(i) = True Then
        
        If .List(i, 3) = CStr(SettPymntStatusAttenteValidationLabel.Val) Then
            SetPaymentValidateFormat uFrm, True, Array("paiement-autoriser")
        ElseIf .List(i, 3) = CStr(SettPymntStatusAttenteExtractionLabel.Val) Then
            SetPaymentExtractFormat uFrm, True
        ElseIf .List(i, 3) = CStr(SettPymntStatusExtraitLabel.Val) Then
            SetPaymentSentFormat uFrm, True
        ElseIf .List(i, 3) = CStr(SettPymntStatusFichierEnvoyeLabel.Val) Then
            SetPaymentExecuteFormat uFrm, True
        End If
        
        SetVisibility uFrm.PaymentShowImg, True, Array("paiement-voir_details")
        
        sPaymentSelectedId = .List(i, 0)
        Set PaymentSelected = NewPayment(CLng(.List(i, 0)))
      End If
    Next i
  End With
End Sub

Public Sub LaunchCreatePayment(uFrm As MSForms.UserForm)
    uFrm.PaymentsMultiPage.Value = 1
    SetCreatedAtFormat uFrm, False
    ResetPaymentForm uFrm
    
    SetPaymentShowFormat uFrm, True
End Sub

Public Sub SelectPaymentEmployee(uFrm As MSForms.UserForm)
    Dim i As Long
    
    paymentEmployeeSelectedCount = 0
    uFrm.PaymentAddImg.Visible = False
    
    With uFrm.PaymentSearchEmployeeListLBx
        For i = 0 To .listCount - 1
          If .Selected(i) = True Then
            If paymentEmployeeSelectedCount = 0 Then
                SetVisibility uFrm.PaymentAddImg, True, Array("paiement-ajouter")
            End If
            'sPaymentEmployeeSelectedId = .list(i, 0)
            paymentEmployeeSelectedCount = paymentEmployeeSelectedCount + 1
          End If
        Next i
    End With
    
End Sub

Public Sub AddPaymentDetail(uFrm As MSForms.UserForm)
    If uFrm.PaymentMotifCBx.ListIndex < 0 Then
        MsgBox "Veuillez sélectionner un Motif", vbCritical, GetAppName
        With uFrm.PaymentMotifCBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If
    If uFrm.PaymentAmountTBx.Text = "" Then
        MsgBox "Veuillez renseigner un Montant", vbCritical, GetAppName
        With uFrm.PaymentAmountTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If
    If Not IsNumeric(uFrm.PaymentAmountTBx.Text) Then
        MsgBox "Montant invalide!", vbCritical, GetAppName
        With uFrm.PaymentAmountTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If
    
    If Not uFrm.PaymentAmountTBx.Text > 0 Then
        MsgBox "Veuillez renseigner un Montant supérieur à 0!", vbCritical, GetAppName
        With uFrm.PaymentAmountTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If
    
    Dim x As Long
    Dim nb As Long
    
    nb = 0
    For x = 0 To uFrm.PaymentSearchEmployeeListLBx.listCount - 1
        If uFrm.PaymentSearchEmployeeListLBx.Selected(x) = True Then
            AddToPaymentDetailsList uFrm, nb, "0", CStr(uFrm.PaymentSearchEmployeeListLBx.List(x, 1)), CStr(uFrm.PaymentSearchEmployeeListLBx.List(x, 2)), CStr(uFrm.PaymentSearchEmployeeListLBx.List(x, 3)), CStr(uFrm.PaymentSearchEmployeeListLBx.List(x, 4)), CStr(uFrm.PaymentAmountTBx.Text), CStr(uFrm.PaymentMotifCBx.Text)
            nb = nb + 1
        End If
    Next x
    
    FormatPaymentActions uFrm
    
    If nb = 0 Then
        MsgBox "Veuillez sélectionner un Employé !", vbCritical, GetAppName
    End If
End Sub

Private Sub AddToPaymentDetailsList(uFrm As MSForms.UserForm, lrow As Long, sId As String, sMatricule As String, sNom As String, sPrenom As String, sTelephone As String, sMontant As String, sMotif As String)
    Dim lastAddedIdx As Long, i As Long, answer As Integer, totalAmount As Long
    
    If Not FormatPhone(sTelephone) Then
        Exit Sub
    End If
    
    If Not FormatMotif(sMotif) Then
        Exit Sub
    End If
    
    ' Check Matricule already exists
    With uFrm.PaymentDetailsLBx
        For i = 0 To .listCount - 1
          If .List(i, 1) = sMatricule Then
            answer = MsgBox("L'Employé " & sNom & " " & sPrenom & " existe déjà dans la Liste. " & vbCrLf & "L'ajouter quand même ?", vbQuestion + vbYesNo + vbDefaultButton2, GetAppName)
            If answer = vbNo Then
                Exit Sub
            End If
          End If
        Next i
    End With
    
    uFrm.PaymentDetailsLBx.AddItem
    
    lastAddedIdx = uFrm.PaymentDetailsLBx.listCount - 1
    
    uFrm.PaymentDetailsLBx.List(lastAddedIdx, 0) = sId
    uFrm.PaymentDetailsLBx.List(lastAddedIdx, 1) = sMatricule
    uFrm.PaymentDetailsLBx.List(lastAddedIdx, 2) = sNom
    uFrm.PaymentDetailsLBx.List(lastAddedIdx, 3) = sPrenom
    uFrm.PaymentDetailsLBx.List(lastAddedIdx, 4) = sTelephone
    uFrm.PaymentDetailsLBx.List(lastAddedIdx, 5) = sMontant
    uFrm.PaymentDetailsLBx.List(lastAddedIdx, 6) = sMotif
    
    UpdateTotalAmount uFrm, CLng(sMontant), True
End Sub

Private Sub UpdateTotalAmount(uFrm As MSForms.UserForm, currAmount As Long, bAdd As Boolean)
    Dim totalAmount As Long
    
    If uFrm.PaymentTotalAmountTBx.Text = "" Then
        totalAmount = 0
    Else
        totalAmount = CLng(uFrm.PaymentTotalAmountTBx.Text)
    End If
    
    If bAdd Then
        uFrm.PaymentTotalAmountTBx.Text = CStr(totalAmount + currAmount)
    Else
        If totalAmount = 0 Or totalAmount < currAmount Then
            uFrm.PaymentTotalAmountTBx.Text = "0"
        Else
            uFrm.PaymentTotalAmountTBx.Text = CStr(totalAmount - currAmount)
        End If
    End If
End Sub

Public Sub ClearPaymentEmployeeDetails(uFrm As MSForms.UserForm)
    uFrm.PaymentMotifCBx.Value = ""
    uFrm.PaymentMotifCBx.ListIndex = -1
    uFrm.PaymentAmountTBx.Text = ""
End Sub

Public Sub PaymentReset(uFrm As MSForms.UserForm)
    If oPaymentSaveForm.Status = Show Then
        uFrm.PaymentsMultiPage.Value = 0
        Exit Sub
    End If
    
    Call ResetPaymentDetails(uFrm)
End Sub

Public Sub SelectPaymentDetail(uFrm As MSForms.UserForm)
    Dim i As Long
    
    If oPaymentSaveForm.Status = Show Then
        'uFrm.PaymentDetailsLBx.ListIndex = -1
        Exit Sub
    End If
    
    With uFrm.PaymentDetailsLBx
        For i = 0 To .listCount - 1
          If .Selected(i) = True Then
            lPaymentDetailSelectedId = i
          End If
        Next i
    End With
    
    FormatPaymentActions uFrm
End Sub

Private Sub InitSearchForm(uFrm As MSForms.UserForm)
    ' ***   Payments Search Form
    Dim paymtEditBtn As CUIControl, paymtDeleteBtn As CUIControl, searchUctl As CUIControl
    
    Set oPaymentSearchForm = NewSearchForm("payments_view", uFrm.PaymentSearchImg, uFrm.PaymentListLBx, uFrm.PaymentSearchCancelImg)
    oPaymentSearchForm.SetResultTitle uFrm.PaymentListLbl, "Paiement", "Paiements"
    
    Set searchUctl = oPaymentSearchForm.AddFieldCtl(uFrm.SearchPaymentTitleTBx, "title", "Titre", True, True). _
    SetClearContentButton(uFrm.SearchPaymentTitleCancelImg)
    Set searchUctl = oPaymentSearchFrm.AddCtl(uFrm.PaymentSearchFrm.SearchPaymentTitleCancelImg).SetSizes(NewSize(10, 12), NewSize(12, 14))
    
    Set searchUctl = oPaymentSearchForm.AddFieldCtl(uFrm.SearchPaymentCreateAtTBx, "created_at", "Date Création", True, True). _
    SetClearContentButton(uFrm.SearchPaymentCreateAtCancelImg)
    Set searchUctl = oPaymentSearchFrm.AddCtl(uFrm.PaymentSearchFrm.SearchPaymentCreateAtCancelImg).SetSizes(NewSize(10, 12), NewSize(12, 14))
    
    Set searchUctl = oPaymentSearchForm.AddFieldCtl(uFrm.SearchPaymentStatusCBx, "status_title", "Statut", True, True)
    
    Set empEditBtn = oPaymentSearchForm.AddFieldCtl(uFrm.PaymentEditImg, "", "", False, False)
    empEditBtn.AddAction Default, setVisibilityState, False
    empEditBtn.AddAction Active, setVisibilityState, True
    Set empDeleteBtn = oPaymentSearchForm.AddFieldCtl(uFrm.PaymentDeleteImg, "", "", False, False)
    empDeleteBtn.AddAction Default, setVisibilityState, False
    empDeleteBtn.AddAction Active, setVisibilityState, True
End Sub

Private Sub InitSearchEmployeeForm(uFrm As MSForms.UserForm)
    ' ***   Payments Search Form
    Dim paymentAddBtn As CUIControl, searchUctl As CUIControl
    
    clr = uFrm.SearchUserLoginTBx.BackColor
    
    Set oPaymentSearchEmployeeForm = NewSearchForm("employees", uFrm.PaymentSearchEmployeeImg, uFrm.PaymentSearchEmployeeListLBx, uFrm.PaymentSearchEmployeeCancelImg)
    oPaymentSearchEmployeeForm.SetResultTitle uFrm.PaymentSearchEmployeeListLbl, "Employé", "Employés"
    oPaymentSearchEmployeeForm.LimitLines = CLng(SettPymntEmployeesSearchLinesSize.Val)
    
    Set searchUctl = oPaymentSearchEmployeeForm.AddFieldCtl(uFrm.PaymentSearchEmployeeMatriculeTBx, "matricule", "Matricule", True, True). _
    SetClearContentButton(uFrm.PaymentSearchEmployeeMatriculeCancelImg)
    Set searchUctl = oPaymentSearchEmployeeFrm.AddCtl(uFrm.PaymentSearchEmployeeFrm.PaymentSearchEmployeeMatriculeCancelImg).SetSizes(NewSize(10, 12), NewSize(12, 14))
    
    Set searchUctl = oPaymentSearchEmployeeForm.AddFieldCtl(uFrm.PaymentSearchEmployeeLastNameTBx, "lastname", "Nom", True, True). _
    SetClearContentButton(uFrm.PaymentSearchEmployeeLastNameCancelImg)
    Set searchUctl = oPaymentSearchEmployeeFrm.AddCtl(uFrm.PaymentSearchEmployeeFrm.PaymentSearchEmployeeLastNameCancelImg).SetSizes(NewSize(10, 12), NewSize(12, 14))
    
    Set searchUctl = oPaymentSearchEmployeeForm.AddFieldCtl(uFrm.PaymentSearchEmployeeFirstNameTBx, "firstname", "Prénom", True, True). _
    SetClearContentButton(uFrm.PaymentSearchEmployeeFirstNameCancelImg)
    Set searchUctl = oPaymentSearchEmployeeFrm.AddCtl(uFrm.PaymentSearchEmployeeFrm.PaymentSearchEmployeeFirstNameCancelImg).SetSizes(NewSize(10, 12), NewSize(12, 14))
    
    oPaymentSearchEmployeeForm.AddFieldCtl uFrm.PaymentSearchEmployeePhoneTBx, "telephone", "Téléphone", True, False
    
    Set paymentAddBtn = oPaymentSearchEmployeeForm.AddFieldCtl(uFrm.PaymentAddImg, "", "", False, False)
    paymentAddBtn.AddAction Default, setVisibilityState, False
    paymentAddBtn.AddAction Active, setVisibilityState, True
End Sub

Private Sub InitSaveForm(uFrm As MSForms.UserForm)
    ' ***   Employees Search Form
    Dim paymentEditBtn As CUIControl, paymentDeleteBtn As CUIControl
    
    Set oPaymentSaveForm = NewSaveForm("payments", uFrm.PaymentSaveTitleLbl, Nothing, Nothing, uFrm.PaymentNewImg)
    ' Save Titles
    oPaymentSaveForm.AddSaveTitle None, ""
    oPaymentSaveForm.AddSaveTitle Add, "Créer Paiement"
    oPaymentSaveForm.AddSaveTitle Update, "Modifier Paiement"
    oPaymentSaveForm.AddSaveTitle Delete, "Suppression Paiement"
    oPaymentSaveForm.AddSaveTitle Show, "Détails Paiement"
    
    ' Save Controls
    
End Sub

Private Sub ResetPaymentSearchEmployee(uFrm As MSForms.UserForm)
    oPaymentSearchEmployeeForm.ResetForm
    uFrm.PaymentMotifCBx.ListIndex = -1
    uFrm.PaymentAmountTBx.Text = ""
End Sub

Private Sub ResetPaymentDetails(uFrm As MSForms.UserForm)
    Dim lbx As MSForms.ListBox, lColumnCount As Long, sColumnWidths As String, sHeaders As String
    Dim headersArr() As String, columnWidthsArr() As String
    
    lColumnCount = 7
    sColumnWidths = "0;70;100;100;70;70;100" ' Id,Mle,Nom,Prenom,Phone,Montant,Motif
    sHeaders = ";   Mlle;      Nom;       Prénom;  Phone;Montant;Motif"
    headersArr = Split(sHeaders, ";")
    columnWidthsArr = Split(sColumnWidths, ";")
    
    uFrm.PaymentDetailsLBx.Clear
    uFrm.PaymentDetailsLBx.ColumnCount = lColumnCount
    uFrm.PaymentDetailsLBx.ColumnWidths = sColumnWidths
    
    uFrm.PaymentTotalAmountTBx.Text = "0"
    
    lPaymentDetailSelectedId = -1
    
    FormatPaymentActions uFrm
    BuildPaymentDetailsHeadersCaption uFrm, headersArr, columnWidthsArr
End Sub

Private Sub BuildPaymentDetailsHeadersCaption(uFrm As MSForms.UserForm, heardersArr() As String, columnWidthsArr() As String)
    Dim i As Long, j As Long, sSpaceUnit As String, maxSpaces As Long
    
    sSpaceUnit = " "
    uFrm.PaymentDetailsHeadersLbl.Caption = ""
    
    For i = 0 To UBound(heardersArr)
        uFrm.PaymentDetailsHeadersLbl.Caption = uFrm.PaymentDetailsHeadersLbl.Caption & heardersArr(i)
        maxSpaces = CInt((CInt(columnWidthsArr(i)) * 0.4))
        For j = 0 To maxSpaces
            uFrm.PaymentDetailsHeadersLbl.Caption = uFrm.PaymentDetailsHeadersLbl.Caption & sSpaceUnit
        Next j
    Next i
    
    'MsgBox uFrm.PaymentDetailsHeadersLbl.Caption
End Sub

Public Sub RemovePaymentDetail(uFrm As MSForms.UserForm)
    Dim sMontant As String, sPaymentId As String
    
    sMontant = CStr(uFrm.PaymentDetailsLBx.List(lPaymentDetailSelectedId, 5))
    sPaymentId = CStr(uFrm.PaymentDetailsLBx.List(lPaymentDetailSelectedId, 0))
    
    If Not (sPaymentId = "0" Or sPaymentId = "") Then
        If Not gobjDB.ExecuteActionQuery("DELETE FROM paymentdetails WHERE Id = " & sPaymentId) Then
           MsgBox "Erreur suppression de ce Détail de la Base de Données !", vbCritical, GetAppName
           Exit Sub
        End If
    End If
    
    UpdateTotalAmount uFrm, CLng(sMontant), False
    
    uFrm.PaymentDetailsLBx.RemoveItem (lPaymentDetailSelectedId)
    
    uFrm.PaymentDetailsLBx.ListIndex = -1
    lPaymentDetailSelectedId = -1
    
    FormatPaymentActions uFrm
End Sub

Private Sub FormatPaymentActions(uFrm As MSForms.UserForm)
    Dim can_act As Boolean, arrPerms As Variant
    
    can_act = (uFrm.PaymentDetailsLBx.listCount > 0)
    
    If can_act Then
        arrPerms = Array("paiement-ajouter", "paiement-modifer")
    Else
        arrPerms = Empty
    End If
    SetVisibility uFrm.PaymentSaveImg, can_act, arrPerms
    SetVisibility uFrm.PaymentSaveLbl, can_act, arrPerms
    
    can_act = (lPaymentDetailSelectedId > -1)
    
    If can_act Then
        arrPerms = Array("paiement-supprimer")
    Else
        arrPerms = Empty
    End If
    SetVisibility uFrm.PaymentRemoveImg, can_act, arrPerms
End Sub

Public Sub ShowPayment(uFrm As MSForms.UserForm)
    EditPayment uFrm
    SetPaymentShowFormat uFrm, False
    oPaymentSaveForm.SwitchStatus Show
End Sub

Private Sub SetPaymentShowFormat(uFrm As MSForms.UserForm, bFormat As Boolean)
    uFrm.PaymentSaveImg.Visible = bFormat
    uFrm.PaymentSaveLbl.Visible = bFormat
    
    'uFrm.PaymentResetImg.Visible = bFormat
    'uFrm.PaymentResetLbl.Visible = bFormat
    
    uFrm.PaymentSearchEmployeeImg.Visible = bFormat
    uFrm.PaymentSearchEmployeeCancelImg.Visible = bFormat
    
    uFrm.PaymentSearchEmployeeListLBx.Locked = Not bFormat
    'uFrm.PaymentDetailsLBx.Locked = bFormat
End Sub

Public Sub EditPayment(uFrm As MSForms.UserForm)
    Dim paymentData As Variant, sqlRst As String, recordCount As Long, i As Long, editProgr As clsProgression
    
    Set editProgr = StartNewProgression("Chargement Formulaire de Paiement", 1)
    
    sqlRst = "SELECT * FROM payments_view WHERE Id = " & sPaymentSelectedId & ""
    
    Call PrepareDatabase
    paymentData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    If recordCount = 0 Then
        MsgBox "Paiement introuvable dans la Base de Données !", vbCritical, GetAppName
        
        editProgr.AddDone 1, True
        Exit Sub
    End If
    
    ResetPaymentSearchEmployee uFrm
    
    uFrm.PaymentIdTBx.Text = CStr(paymentData(0, 0))
    uFrm.PaymentTitleTBx.Text = CStr(paymentData(4, 0))
    uFrm.PaymentCreatedAtTBx.Text = Format(CStr(paymentData(1, 0)), "dd-mm-yyyy hh:mm")
    
    ResetPaymentDetails uFrm
    
    sqlRst = "SELECT * FROM paymentdetails WHERE payment_id = " & sPaymentSelectedId & ""

    Call PrepareDatabase
    paymentData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    If recordCount = 0 Then
        MsgBox "Aucun Détail pour ce Paiement dans la Base de Données !", vbCritical, GetAppName
        
        editProgr.AddDone 1, True
        Exit Sub
    End If
    
    editProgr.StartNewSubProgression "Chargement Détails du Paiement", recordCount
    For i = 0 To recordCount - 1
        AddToPaymentDetailsList uFrm, i, CStr(paymentData(0, i)), CStr(paymentData(2, i)), CStr(paymentData(3, i)), CStr(paymentData(4, i)), CStr(paymentData(6, i)), CStr(paymentData(9, i)), CStr(paymentData(10, i))
        editProgr.AddDoneLastSub 1, True
    Next i
    
    editProgr.AddDone 1, True
    
    oPaymentSaveForm.SwitchStatus Update
    SetCreatedAtFormat uFrm, True
    FormatPaymentActions uFrm
    uFrm.PaymentsMultiPage.Value = 1
    
    SetPaymentShowFormat uFrm, True
End Sub

Public Sub SavePayment(uFrm As MSForms.UserForm)
    If uFrm.PaymentTitleTBx.Value = "" Then
        MsgBox "Veuillez rensigner un Titre", vbCritical, GetAppName
        With uFrm.PaymentTitleTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If
    
    If uFrm.PaymentDetailsLBx.listCount <= 0 Then
        MsgBox "Veuillez ajouter au moins un Employé dans la liste de bénéficiaires", vbCritical, GetAppName
        With uFrm.PaymentTitleTBx
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If
    
    If oPaymentSaveForm.Status = Add Then
        If CreatePayment(uFrm) Then
            InitFormat uFrm
            uFrm.PaymentsMultiPage.Value = 0
        End If
    Else
        If UpdatePayment(uFrm) Then
            InitFormat uFrm
            uFrm.PaymentsMultiPage.Value = 0
        End If
    End If
End Sub

Public Sub ValidatePayment(uFrm As MSForms.UserForm)
    Dim paymentData As Variant, sqlRst As String, validatedStatusId As Long, audit As clsAudit

    Set audit = loggedUser.StartNewAudit("Validation Paiement (" & sPaymentSelectedId & ") Id Statut: " & CStr(validatedStatusId))
    
    If SetPaymentValidated Then
        MsgBox "Paiement Validé avec Succès !", vbInformation, GetAppName

        If SettPymntNotifyAttenteExtraction.Val Then
            MailObject.SendMailToMany "Paiement en Attente Extraction", "<p>Bonjour. </p><p>Vous avez un Nouveau Paiement (<strong>" & CStr(PaymentSelected(0)) & " - " & CStr(PaymentSelected(1)) & "</strong>) en Attente Extraction (de fichier)</p>", GetUserMailsByPermissions(Array("paiement-extraire"))
        End If
        
        audit.EndWithSuccess
        oPaymentSearchForm.Search True
        InitPaymentListFormat uFrm
        uFrm.PaymentsMultiPage.Value = 0
    Else
        audit.EndWithFailure
        MsgBox "Erreur Validation Paiement dans la Base de Données !", vbCritical, GetAppName
        Exit Sub
    End If
End Sub

Public Sub ExtractPayment(uFrm As MSForms.UserForm)
    Dim paymentData As Variant, sqlRst As String, extractedStatusId As Long, audit As clsAudit
    
    extractedStatusId = CLng(SettPymntStatusExtraitId.Val)
    Set audit = loggedUser.StartNewAudit("Extraction Paiement (" & sPaymentSelectedId & ") Id Statut: " & CStr(extractedStatusId))
    
    ClearPaymentSheet
    If AddPaymentDetailsToSheet(uFrm, sPaymentSelectedId) Then
        Dim exportFile As String, paymentFileName As String

        If ExportToCsv(paymentFileName, exportFile) Then
            sqlRst = "UPDATE payments SET status_id = " & extractedStatusId & ", extracted_at = '" & CStr(Now) & "', extracted_by = " & CStr(loggedUser.Id) & ", file_name = '" & CStr(paymentFileName) & "' WHERE Id = " & sPaymentSelectedId & ""
            Call PrepareDatabase
            If gobjDB.ExecuteActionQuery(sqlRst) Then
                MsgBox "Paiement Extrait avec Succès !", vbInformation, GetAppName
                
                If SettPymntNotifyExtraction.Val Then
                    MailObject.SendMailToMany "Paiement Extrait", "<p>Bonjour. </p><p>Le Paiement <strong>" & CStr(PaymentSelected(1)) & "</strong> (" & sPaymentSelectedId & ") a été Extrait</p>", GetUserMailsByPermissions(Array("paiement-marquer_execute"))
                End If
                SendExtractedFile exportFile
                
                audit.EndWithSuccess
                oPaymentSearchForm.Search True
                InitPaymentListFormat uFrm
                uFrm.PaymentsMultiPage.Value = 0
            Else
                audit.EndWithFailure
                MsgBox "Erreur Extraction Paiement dans la Base de Données !", vbCritical, GetAppName
                Exit Sub
            End If
        Else
            audit.EndWithFailure
        End If
        
    Else
        audit.EndWithFailure
    End If
End Sub

Public Sub SendPayment(uFrm As MSForms.UserForm)
    Dim exportFile As String, paymentFileName As String
    
    SendExtractedFile exportFile
End Sub

Public Sub ExecutePayment(uFrm As MSForms.UserForm)
    Dim paymentData As Variant, sqlRst As String, executedStatusId As Long, audit As clsAudit
    
    executedStatusId = CLng(SettPymntStatusEffectueId.Val)

    Set audit = loggedUser.StartNewAudit("Paiement (" & sPaymentSelectedId & ") marqué Comme Exécuté, Id Statut: " & CStr(executedStatusId))
    
    sqlRst = "UPDATE payments SET status_id = " & executedStatusId & ", executed_at = '" & CStr(Now) & "', executed_by = " & CStr(loggedUser.Id) & " WHERE Id = " & sPaymentSelectedId & ""
    
    Call PrepareDatabase
    If gobjDB.ExecuteActionQuery(sqlRst) Then
        MsgBox "Paiement Exécuté avec Succès !", vbInformation, GetAppName

        If SettPymntNotifyEffectue.Val Then
            MailObject.SendMailToMany "Paiement Executé", "<p>Bonjour. </p><p>Le Paiement <strong>" & CStr(PaymentSelected(1)) & "</strong> (" & sPaymentSelectedId & ") a été exécuté via MOBICASH</p>", GetUserMailsByPermissions(Array("paiement-ajouter"))
        End If
        
        audit.EndWithSuccess
        oPaymentSearchForm.Search True
        InitPaymentListFormat uFrm
        uFrm.PaymentsMultiPage.Value = 0
    Else
        audit.EndWithFailure
        MsgBox "Erreur Exécution Paiement dans la Base de Données !", vbCritical, GetAppName
        Exit Sub
    End If
End Sub

Private Sub ValidatePaymentOld(uFrm As MSForms.UserForm)
    ClearPaymentSheet
    AddPaymentListToSheet uFrm
    ExportToCsv
    InsertPaymentToDB uFrm
    ResetPaymentDetails uFrm
End Sub

Private Function CreatePayment(uFrm As MSForms.UserForm) As Boolean
    Dim queryStr As String, newPaymentId As Variant, createdStatusId As Long, createProgr As clsProgression, audit As clsAudit
    
    Set createProgr = StartNewProgression("Création Nouveau Paiement", 1)
    Set audit = loggedUser.StartNewAudit("Création Nouveau Paiement " & uFrm.PaymentTitleTBx.Text)
    
    createdStatusId = CLng(SettPymntStatusAttenteValidationId.Val)
    ' Add New Payment
    queryStr = "INSERT INTO payments (created_at,created_by,title,status_id,amount) VALUES ('" & CStr(Now) & "', " & loggedUser.Id & ", '" & SqlStringVar(uFrm.PaymentTitleTBx.Text) & "'," & createdStatusId & "," & uFrm.PaymentTotalAmountTBx.Text & ")"
    
    If gobjDB.ExecuteActionQuery(queryStr, newPaymentId) Then
        ' Add New Payment Details
        InsertPaymentDetailsToDB uFrm, newPaymentId, createProgr
        
        MsgBox "Paiement Créé avec Succès", vbInformation, GetAppName

        If SettPymntNotifyAttenteValidation.Val Then
            MailObject.SendMailToMany "Nouveau Paiement", "<p>Bonjour. </p><p>Vous avez un Nouveau Paiement (<strong>" & CStr(newPaymentId) & " - " & CStr(uFrm.PaymentTitleTBx.Text) & "</strong>) à Valider</p>", GetUserMailsByPermissions(Array("paiement-autoriser"))
        End If
        
        audit.EndWithSuccess
        CreatePayment = True
    Else
        MsgBox "Erreur Insertion Paiement dans la Base de Données", vbCritical, GetAppName
        
        audit.EndWithFailure
        CreatePayment = False
        'InitFormat uFrm
        'uFrm.PaymentsMultiPage.value = 0
    End If
    
    createProgr.AddDone 1, True
End Function

Private Function UpdatePayment(uFrm As MSForms.UserForm) As Boolean
    Dim queryStr As String, updProgr As clsProgression, audit As clsAudit
    
    Set audit = loggedUser.StartNewAudit("Modification Paiement (" & uFrm.PaymentIdTBx.Text & ") " & uFrm.PaymentTitleTBx.Text)
    Set updProgr = StartNewProgression("", 1)
    
    ' Update Payment
    queryStr = "UPDATE payments SET title = '" & SqlStringVar(uFrm.PaymentTitleTBx.Text) & "', amount = " & uFrm.PaymentTotalAmountTBx.Text & " WHERE Id = " & uFrm.PaymentIdTBx.Text & ""
    
    'MsgBox queryStr, vbInformation, GetAppName
    
    If gobjDB.ExecuteActionQuery(queryStr) Then
        ' Add New Payment Details
        InsertPaymentDetailsToDB uFrm, uFrm.PaymentIdTBx.Text, updProgr
        
        MsgBox "Paiement Mise à Jour avec Succès", vbInformation, GetAppName
        
        audit.EndWithSuccess
        UpdatePayment = True
    Else
        MsgBox "Erreur Mise à Jour Paiement dans la Base de Données", vbCritical, GetAppName
        
        audit.EndWithFailure
        UpdatePayment = False
    End If
    
    updProgr.AddDone 1, True
End Function

Public Sub DeletePayment(uFrm As MSForms.UserForm)
    Dim answer As Integer
    Dim queryStr As String
    Dim delProgr As clsProgression, audit As clsAudit
    
    answer = MsgBox("Supprimer ce Paiement ?", vbQuestion + vbYesNo + vbDefaultButton2, GetAppName)
    
    If answer = vbYes Then
        Set delProgr = StartNewProgression("Suppression Paiement", 1)
        
        Set audit = loggedUser.StartNewAudit("Suppression Paiement (" & uFrm.PaymentIdTBx.Text & ") " & uFrm.PaymentTitleTBx.Text)
        ' Delete Payment
        queryStr = "DELETE FROM payments WHERE Id = " & sPaymentSelectedId & ""
        
        If gobjDB.ExecuteActionQuery(queryStr) Then
            MsgBox "Paiement supprimé avec Succès", vbInformation, GetAppName
            
            oPaymentSearchForm.Search True
            InitPaymentListFormat uFrm
            uFrm.PaymentsMultiPage.Value = 0
            
            audit.EndWithSuccess
        Else
            audit.EndWithFailure
            MsgBox "Erreur Supression Paiement !", vbCritical, GetAppName
        End If
        
        delProgr.AddDone 1, True
    End If
    
End Sub

Private Sub ClearPaymentSheet()
    Dim ws As Worksheet
    Dim lastrow As Long
    
    Set ws = ActiveWorkbook.Sheets("Curr Payment")
    
    ws.Activate
    ActiveSheet.Cells.ClearContents
    ActiveSheet.Cells.ClearFormats
    
End Sub

Private Function AddPaymentDetailsToSheet(uFrm As MSForms.UserForm, PaymentId) As Boolean
    Dim ws As Worksheet, lastrow As Long, i As Long, nbemployees As Long
    Dim paymentData As Variant, sqlRst As String, recordCount As Long
    
    Set ws = ActiveWorkbook.Sheets("Curr Payment")
    ws.Activate
    ActiveSheet.Cells.NumberFormat = "@"
    
    lastrow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If ws.Cells(lastrow, 1) <> "" Then
        lastrow = lastrow + 1
    End If
        
    ' Headers
    ws.Cells(lastrow, 1) = "Receiver Identitfier Type"
    ws.Cells(lastrow, 2) = "Receiver Identifier"
    ws.Cells(lastrow, 3) = "Validation KYC (O)"
    ws.Cells(lastrow, 4) = "Validation KYC value(O)"
    ws.Cells(lastrow, 5) = "Amount"
    ws.Cells(lastrow, 6) = "Comment"
    
    nbemployees = 0
    sqlRst = "SELECT * FROM paymentdetails WHERE payment_id = " & CStr(PaymentId) & ""

    Call PrepareDatabase
    paymentData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    If recordCount = 0 Then
        MsgBox "Aucun Détail pour ce Paiement dans la Base de Données !", vbCritical, GetAppName
        AddPaymentDetailsToSheet = False
        Exit Function
    End If
    
    For i = 0 To recordCount - 1
        lastrow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
        If ws.Cells(lastrow, 1) <> "" Then
            lastrow = lastrow + 1
        End If
        
        ws.Cells(lastrow, 1) = CStr(paymentData(5, i))
        ws.Cells(lastrow, 2) = CStr(paymentData(6, i))
        ws.Cells(lastrow, 3) = ""
        ws.Cells(lastrow, 4) = ""
        ws.Cells(lastrow, 5) = CStr(paymentData(9, i))
        ws.Cells(lastrow, 6) = CStr(paymentData(10, i))
        'ws.Cells(lastrow, 7) = CStr(Now)
        
        nbemployees = nbemployees + 1
    Next i
    
    AddPaymentDetailsToSheet = True
    
End Function

Private Function ExportToCsv(ByRef paymentFileName As String, ByRef exportFile As String) As Boolean
  Dim ws As Worksheet
  Dim ColNum As Long
  Dim Line As String
  Dim LineValues() As Variant
  Dim OutputFileNum As Integer
  Dim PathName As String
  Dim FileName As String
  Dim RowNum As Long
  Dim RowCount As Long
  Dim SheetValues() As Variant
  
  Set ws = ActiveWorkbook.Sheets("Curr Payment")

  FileName = "mobicashpayment_" & NowTimeStamp & ".csv"
  PathName = CStr(SettPymntExtractionFileFolder.Val)
  'While PathName = ""
  ' PathName = GetFolder ' Application.ActiveWorkbook.path
  'Wend
  
  OutputFileNum = FreeFile
  PathName = PathName & Application.PathSeparator & FileName
    
  Open PathName For Output Lock Write As #OutputFileNum

  'Print #OutputFileNum, "Field1" & "," & "Field2"

  RowCount = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
  SheetValues = Sheets("Curr Payment").Range("A1:F" & CStr(RowCount)).Value
  
  ReDim LineValues(1 To 6)

  For RowNum = 1 To RowCount
    For ColNum = 1 To 6
      LineValues(ColNum) = CStr(SheetValues(RowNum, ColNum))
    Next
    Line = Join(LineValues, ",")
    Print #OutputFileNum, Line
  Next
  
  If RowCount > 0 Then
    MsgBox "Fichier créé et exporté vers " & PathName & "", vbInformation, GetAppName
  End If

  Close OutputFileNum
  
  paymentFileName = FileName
  exportFile = PathName
  ExportToCsv = True
End Function

Private Sub InsertPaymentDetailsToDB(uFrm As MSForms.UserForm, PaymentId As Variant, Optional insertProgr As clsProgression)
    Dim x As Long
    Dim nb As Long
    
    StartNewSubProgressionFromParent "Insertion Details de Paiement", uFrm.PaymentDetailsLBx.listCount, insertProgr
    
    nb = 0
    For x = 0 To uFrm.PaymentDetailsLBx.listCount - 1
        If uFrm.PaymentDetailsLBx.List(x, 0) = "0" Or uFrm.PaymentDetailsLBx.List(x, 0) = "" Then
            InsertPaymentDetailToDB uFrm, PaymentId, uFrm.PaymentDetailsLBx.List(x, 1), uFrm.PaymentDetailsLBx.List(x, 2), uFrm.PaymentDetailsLBx.List(x, 3), "MSISDN", uFrm.PaymentDetailsLBx.List(x, 4), CStr(uFrm.PaymentDetailsLBx.List(x, 5)), CStr(uFrm.PaymentDetailsLBx.List(x, 6))
        End If
        
        AddDoneLastSubFromParent 1, True, insertProgr
        nb = nb + 1
    Next x
End Sub

Private Function InsertPaymentDetailToDB(uFrm As MSForms.UserForm, payment_id, employee_matricule, employee_lastname, employee_firstname, Receiver_Identitfier_Type, Receiver_Identifier, Amount, Comment) As Integer
    Dim queryStr As String, newPaymentDetailId As Variant, audit As clsAudit
    
    Set audit = loggedUser.StartNewAudit("Ajout Employé (" & CStr(employee_matricule) & ") " & CStr(employee_lastname) & " " & CStr(employee_firstname) & " au Paiement (" & CStr(payment_id) & ") " & uFrm.PaymentTitleTBx.Text)
    
    Call PrepareDatabase
    ' Add New Payment detail
    queryStr = "INSERT INTO paymentdetails (payment_id,employee_matricule,employee_lastname,employee_firstname,Receiver_Identitfier_Type,Receiver_Identifier,Amount,Comment)" & _
     " VALUES(" & CStr(payment_id) & ",'" & SqlStringVar(employee_matricule) & "','" & SqlStringVar(employee_lastname) & "','" & SqlStringVar(employee_firstname) & "','" & SqlStringVar(Receiver_Identitfier_Type) & "','" & SqlStringVar(Receiver_Identifier) & "'," & CStr(Amount) & ",'" & SqlStringVar(Comment) & "')"
    
    If gobjDB.ExecuteActionQuery(queryStr, newPaymentDetailId) Then
        audit.EndWithSuccess
        InsertPaymentDetailToDB = CInt(newPaymentDetailId)
    Else
        audit.EndWithFailure
        InsertPaymentDetailToDB = -1
    End If
End Function

Private Sub ResetPaymentForm(uFrm As MSForms.UserForm)
    ClearPaymentEmployeeDetails uFrm
    ResetPaymentDetails uFrm
    uFrm.PaymentTitleTBx.Text = ""
    uFrm.PaymentCreatedAtTBx.Text = ""
    uFrm.PaymentTotalAmountTBx.Text = ""
    oPaymentSearchEmployeeForm.ResetForm
End Sub


Public Sub InitPaymentMotifsCbx(uFrm As MSForms.UserForm)
    Dim oSql As clsSqlQuery, oSelectFields As Collection
    
    'uFrm.PaymentMotifCBx.ColumnCount = 2
    'uFrm.PaymentMotifCBx.ColumnWidths = ";0"
    
    Set oSelectFields = New Collection
    oSelectFields.Add NewDbField("title", "Titre", 0, True, True, True, True, True), "title"
    Set oSql = NewSqlQuery("paymentreasons")
    oSql.SelectToListByCriterion NewUCTL(uFrm.PaymentMotifCBx), oSelectFields, Nothing, False
    
End Sub

Public Sub InitPaymentStatus(uFrm As MSForms.UserForm)
    Dim oSql As clsSqlQuery, oSelectFields As Collection
    
    'uFrm.SearchPaymentStatusCBx.ColumnCount = 2
    'uFrm.SearchPaymentStatusCBx.ColumnWidths = ";0"
    
    Set oSelectFields = New Collection
    oSelectFields.Add NewDbField("status_title", "Titre", 0, True, True, True, True, True), "status_title"
    Set oSql = NewSqlQuery("paymentstatus")
    oSql.SelectToListByCriterion NewUCTL(uFrm.SearchPaymentStatusCBx), oSelectFields, Nothing, False
    
End Sub

Private Function FormatMotif(ByRef sMotif) As Boolean
    Dim specCharsArr() As String, specCharSingleArr() As String, i As Long
    
    specCharsArr = Split(CStr(SettMotifSpecialChars.Val), ";")
    For i = 0 To UBound(specCharsArr)
        specCharSingleArr = Split(specCharsArr(i), ",")
        sMotif = Replace(sMotif, specCharSingleArr(0), specCharSingleArr(1))
    Next i
    
    FormatMotif = True
End Function

Public Sub ManagePaymentMotif(uFrm As MSForms.UserForm)
    ManageSubList "paymentreasons", "Motif", "Motifs", "motif_paiement-ajouter", "motif_paiement-modifier", "motif_paiement-supprimer"
    InitPaymentMotifsCbx uFrm
End Sub

Private Function SendExtractedFile(exportFile As String) As Boolean
    Dim recipientsArr As Variant, sMailSubject As String, sMailBody As String
    
    If SettPymntExtractionDontSendFile.Val Then
        SendExtractedFile = True
        Exit Function
    End If

    Dim receiversArr As Variant

    sMailSubject = "Nouveau Fichier de Paiement à Exécuter"
    sMailBody = "<p>Bonjour. </p>" & "<p>Vous avez (ci-joint) Le Fichier de Paiement <strong>" & CStr(PaymentSelected(1)) & "</strong> (" & CStr(PaymentSelected(0)) & ") à Exécuter.</p>"
    
    If SettPymntExtractionSendFileToList.Val Then
        MailObject.SendMailToMany sMailSubject, sMailBody, GetExtractedFileRecipientsFromList(receiversArr), exportFile
        
        SetPaymentFileSent
        SendExtractedFile = True
        Exit Function
    End If
    
    If SettPymntExtractionSendFileToUsers.Val Then
        MailObject.SendMailToMany sMailSubject, sMailBody, GetUserMailsByPermissions(Array("paiement-marquer_execute")), exportFile
        
        SetPaymentFileSent
        SendExtractedFile = True
        Exit Function
    End If

    If SettPymntExtractionSendFileToAll.Val Then
        receiversArr = GetUserMailsByPermissions(Array("paiement-marquer_execute"))
        MailObject.SendMailToMany sMailSubject, sMailBody, GetExtractedFileRecipientsFromList(receiversArr), exportFile
        
        SetPaymentFileSent
        SendExtractedFile = True
        Exit Function
    End If

    SendExtractedFile = False
End Function

Private Function GetExtractedFileRecipientsFromList(receiversList As Variant) As Variant
    Dim receiversArr() As String, currReceiverArr() As String, i As Integer

    receiversArr = Split(CStr(SettPymntExtractionReceiversList.Val), SettPymntExtractionReceiversList.ArraySep)
    For i = 0 To UBound(receiversArr)
        currReceiverArr = Split(receiversArr(i), ";")
        If currReceiverArr(2) = "1" Then
            receiversList = AddToArray(receiversList, Array(currReceiverArr(0), currReceiverArr(1)))
        End If
    Next i

    GetExtractedFileRecipientsFromList = receiversList
End Function

Private Function SetPaymentValidated() As Boolean
    SetPaymentValidated = UpdatePaymentStatus(CLng(SettPymntStatusAttenteExtractionId.Val), "validated_at", "validated_by")
End Function

Private Function SetPaymentExtracted() As Boolean
    SetPaymentExtracted = UpdatePaymentStatus(CInt(SettPymntStatusExtraitId.Val), "extracted_at", "extracted_by")
End Function

Private Function SetPaymentFileSent() As Boolean
    SetPaymentFileSent = UpdatePaymentStatus(CInt(SettPymntStatusFichierEnvoyeId.Val), "sent_at", "sent_by")
End Function

Private Function SetPaymentExecuted() As Boolean
    SetPaymentExecuted = UpdatePaymentStatus(CInt(SettPymntStatusEffectueId.Val), "executed_at", "executed_by")
End Function

Private Function UpdatePaymentStatus(newStatusId As Long, changeDateField As String, changeByField As String) As Boolean
    Dim paymentData As Variant, sqlRst As String
    
    sqlRst = "UPDATE payments SET status_id = " & newStatusId & ", " & changeDateField & " = '" & CStr(Now) & "', " & changeByField & " = " & CStr(loggedUser.Id) & " WHERE Id = " & sPaymentSelectedId & ""
    Call PrepareDatabase
    If gobjDB.ExecuteActionQuery(sqlRst) Then
        UpdatePaymentStatus = True
        Exit Function
    Else
        UpdatePaymentStatus = False
        Exit Function
    End If
End Function

