Option Explicit

Private m_lId As Long
Private m_reqstTable As String

Private m_fieldListColl As clsUCTLList
Private m_lLastFieldIndex As Long

Private m_vStatus As saveFormStatusType

Private m_sqlQuery As clsSqlQuery

Private m_saveTitleListColl As Collection
Private m_lLastResult As Long

Private WithEvents m_saveTitle As clsUCTL
Private WithEvents m_saveBtn As clsUCTL
Private WithEvents m_newBtn As clsUCTL
Private WithEvents m_cancelBtn As clsUCTL


'*****************************************************************************************************************************
'    Public events
'*****************************************************************************************************************************

Public Event FormSaved()


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As clsSaveForm
    Set Self = Me
End Property

' m_lId
Public Property Let Id(lId As Long)
    m_lId = lId
End Property

Public Property Get Id() As Long
    Id = m_lId
End Property

' FieldList
Public Property Set FieldList(oFieldListColl As clsUCTLList)
    Set m_fieldListColl = oFieldListColl
End Property

Public Property Get FieldList() As clsUCTLList
    Set FieldList = m_fieldListColl
End Property

' LastFieldIndex
Public Property Let LastFieldIndex(lLastFieldIndex As Long)
    m_lLastFieldIndex = lLastFieldIndex
End Property

Public Property Get LastFieldIndex() As Long
    LastFieldIndex = m_lLastFieldIndex
End Property

' SaveTitle
Public Property Set SaveTitle(oSaveTitle As clsUCTL)
    Set m_saveTitle = oSaveTitle
End Property

Public Property Get SaveTitle() As clsUCTL
    Set SaveTitle = m_saveTitle
End Property

' SaveTitleList
Public Property Set SaveTitleList(cSaveTitleList As Collection)
    Set m_saveTitleListColl = cSaveTitleList
End Property

Public Property Get SaveTitleList() As Collection
    Set SaveTitleList = m_saveTitleListColl
End Property

' SaveBtn
Public Property Set SaveBtn(oSaveBtn As clsUCTL)
    Set m_saveBtn = oSaveBtn
End Property

Public Property Get SaveBtn() As clsUCTL
    Set SaveBtn = m_saveBtn
End Property

' CancelBtn
Public Property Set CancelBtn(oCancelBtn As clsUCTL)
    Set m_cancelBtn = oCancelBtn
End Property

Public Property Get CancelBtn() As clsUCTL
    Set CancelBtn = m_cancelBtn
End Property

' NewBtn
Public Property Set NewBtn(oNewBtn As clsUCTL)
    Set m_newBtn = oNewBtn
End Property

Public Property Get NewBtn() As clsUCTL
    Set NewBtn = m_newBtn
End Property

' Status
Public Property Let Status(vStatus As saveFormStatusType)
    m_vStatus = vStatus
End Property

Public Property Get Status() As saveFormStatusType
    Status = m_vStatus
End Property

' RequestTable
Public Property Let RequestTable(sRequestTable As String)
    m_reqstTable = sRequestTable
End Property

Public Property Get RequestTable() As String
    RequestTable = m_reqstTable
End Property

' SqlQuery
Public Property Set SqlQuery(oSqlQuery As clsSqlQuery)
    Set m_sqlQuery = oSqlQuery
End Property

Public Property Get SqlQuery() As clsSqlQuery
    Set SqlQuery = m_sqlQuery
End Property

' LastResult
Public Property Let LastResult(lLastResult As Long)
    m_lLastResult = lLastResult
End Property

Public Property Get LastResult() As Long
    LastResult = m_lLastResult
End Property


'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(sRequestTable As String, oSaveTitle As MSForms.Control, oSaveBtn As MSForms.Control, oCancelBtn As MSForms.Control, Optional oNewBtn As MSForms.Control)
    Me.RequestTable = sRequestTable
    Set Me.SaveTitle = NewUCTL(oSaveTitle)
    Set Me.SaveBtn = NewUCTL(oSaveBtn)
    Set Me.CancelBtn = NewUCTL(oCancelBtn)
    
    If Not IsMissing(oNewBtn) Then
        Set Me.NewBtn = NewUCTL(oNewBtn)
    End If
    
    Me.Status = None
    Set Me.SqlQuery = NewSqlQuery(sRequestTable)
End Sub

Public Function AddFieldCtl(oCtl As MSForms.Control, sFieldName As String, sFieldTitle As String, Optional bIsRequired As Boolean, _
Optional bIsSavable As Boolean, Optional bIsCancellable As Boolean, Optional oCtlLabel As MSForms.Control = Nothing, _
Optional bIsCriteria As Boolean = False, Optional bIsResultField As Boolean = False) As clsUCTL
    Dim oUctl As clsUCTL
    
    Set oUctl = NewUCTL(oCtl)
    
    Me.LastFieldIndex = Me.LastFieldIndex + 1
    
    oUctl.SetField sFieldName, sFieldTitle, Me.LastFieldIndex, bIsSavable, bIsCriteria, bIsResultField, bIsRequired, bIsCancellable
    Me.FieldList.AddCtlByUCTL oUctl
    
    If Not oCtlLabel Is Nothing Then
        oUctl.SetCtlLabel oCtlLabel
    End If
    
    Set AddFieldCtl = oUctl
    
End Function

Public Function GetFieldCtl(sCtlName As String) As clsUCTL
    Set GetFieldCtl = Me.FieldList.GetCTL(sCtlName)
End Function

Public Sub SetFieldCtlErrorStyle(sCtlName As String)
    Dim oUctl As clsUCTL
    
    Set oUctl = Me.FieldList.GetCTL(sCtlName)
    oUctl.BackColor = oUctl.ErrorColor
End Sub

Public Sub SetFieldCtlNormalStyle(sCtlName As String)
    Dim oUctl As clsUCTL
    
    Set oUctl = Me.FieldList.GetCTL(sCtlName)
    oUctl.BackColor = oUctl.NormalColor
End Sub

Public Sub AddSaveTitle(vSaveStatus As saveFormStatusType, sTitle As String)
    Dim i As Long, listCount As Long
    
    If CollKeyExists(Me.SaveTitleList, CStr(vSaveStatus)) Then
        Me.SaveTitleList.Remove CStr(vSaveStatus)
    End If
    
    Me.SaveTitleList.Add sTitle, CStr(vSaveStatus)
    
End Sub

Public Sub SetUctlRequiered(sCtlName As String, bRequired As Boolean)
    Dim oUctl As clsUCTL
    
    Set oUctl = Me.FieldList.GetCTL(sCtlName)
    oUctl.Field.IsRequired = bRequired
End Sub

Public Sub ExecEdit(vId As Variant)
    Dim editProgr As clsProgression
    
    Set editProgr = StartNewProgression("Chargement Formulaire Edition ", 1)
    
    Me.Id = CLng(vId)
    ClearFields Update, editProgr
    FillFields vId, Update, editProgr
    
    editProgr.AddDone 1, True
End Sub

Public Sub ExecDelete(vId As Variant)
    Dim delProgr As clsProgression
    
    Set delProgr = StartNewProgression("Suppression ", 1)
    
    Me.Id = CLng(vId)
    ClearFields Delete, delProgr
    
    DeleteFromId delProgr
    
    delProgr.AddDone 1, True
End Sub

Public Function Save() As Boolean
    If Me.Status = Add Then
        Save = InsertToTable
    ElseIf Me.Status = Update Then
        Save = UpdateFromId
    End If
    
    OnFormSaved
End Function

Public Sub ResetForm(Optional vSaveStatus As saveFormStatusType)
    ClearFields vSaveStatus
End Sub

Private Function ClearFields(Optional vSaveStatus As saveFormStatusType, Optional clearProgr As clsProgression) As Long
    Dim fieldsCount As Long, itemColl As Variant, oUctl As clsUCTL
    
    StartNewSubProgressionFromParent "Nettoyage des Champs", 1, clearProgr
    
    fieldsCount = 0
    
    If Me.FieldList Is Nothing Then
    Else
        For Each itemColl In Me.FieldList.ListCtls
            AddToDoLastSubFromParent 1, clearProgr
            Set oUctl = itemColl
            
            If oUctl.Field.IsCancellable Then
                oUctl.Clear
                oUctl.BackColor = oUctl.NormalColor
            End If
            AddDoneLastSubFromParent 1, True, clearProgr
        Next itemColl
    End If
    
    If Not IsMissing(vSaveStatus) Then
        SwitchStatus vSaveStatus
    End If
    
    AddDoneLastSubFromParent 1, True, clearProgr
    
    ClearFields = fieldsCount
End Function

Private Function FillFields(Id As Variant, Optional vSaveStatus As saveFormStatusType, Optional fillProgr As clsProgression) As Long
    Dim resultCount As Long, itemColl As Variant, oUctl As clsUCTL
    
    resultCount = Me.SqlQuery.SelectToFieldsById(Me.FieldList.GetSavableUCtls, Me.Id, fillProgr)
    
    If resultCount <= 0 Then
        MsgBox "Erreur Remplissage Formulaire !" & vbCrLf & "Aucune donnée trouvée !", vbCritical, GetAppName
    End If
End Function

Private Function InsertToTable() As Boolean
    Dim insertProgr As clsProgression
    
    Set insertProgr = StartNewProgression("Création ", 1)
    
    Me.LastResult = Me.SqlQuery.SaveToTable(Add, Me.FieldList.GetSavableUCtls, -1, insertProgr)
    
    If Me.LastResult > 0 Then
        MsgBox "Création effectuée avec Succès", vbInformation, GetAppName
        InsertToTable = True
    Else
        InsertToTable = False
    End If
    
    insertProgr.AddDone 1, True
End Function

Private Function UpdateFromId() As Boolean
    Dim updProgr As clsProgression
    
    Set updProgr = StartNewProgression("Mise à Jour ", 1)
    
    Me.LastResult = Me.SqlQuery.SaveToTable(Update, Me.FieldList.GetSavableUCtls, CLng(Me.Id), updProgr)
    
    If Me.LastResult > 0 Then
        MsgBox "Mise a jour effectuée avec Succès" & vbCrLf & "(" & CStr(Me.LastResult) & ") Ligne(s) affectées !", vbInformation, GetAppName
        UpdateFromId = True
    Else
        UpdateFromId = False
    End If
    
    updProgr.AddDone 1, True
End Function

Private Function DeleteFromId(Optional delProgr As clsProgression) As Long
    
    Me.LastResult = Me.SqlQuery.ExecDeleteFromId(CLng(Me.Id), delProgr)
    
    If Me.LastResult = 0 Then
        MsgBox "Aucune Ligne Affectée !", vbCritical, GetAppName
    Else
        'Return a good return code
        MsgBox "Succès (" & CStr(Me.LastResult) & ") Suppression !", vbInformation, GetAppName
    End If
End Function

Public Sub SwitchStatus(vSaveStatus As saveFormStatusType)
    Me.Status = vSaveStatus
    
    If CollKeyExists(Me.SaveTitleList, CStr(vSaveStatus)) Then
        Me.SaveTitle.Text = Me.SaveTitleList.Item(CStr(vSaveStatus))
    Else
        MsgBox "Titre de Statut '" & CStr(vSaveStatus) & " NON EXISTANT !", vbCritical, GetAppName
    End If
End Sub

Private Sub OnFormSaved()
    RaiseEvent FormSaved
End Sub


'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    Set Me.FieldList = New clsUCTLList
    Set m_saveTitleListColl = New Collection
    
    Me.LastFieldIndex = -1
End Sub

Private Sub m_cancelBtn_Click(oUctl As clsUCTL)
    'ClearFields Add
End Sub

Private Sub m_newBtn_Click(oUctl As clsUCTL)
    ClearFields Add
End Sub

Private Sub m_saveBtn_Click(oUctl As clsUCTL)
    Me.Save
End Sub

Private Sub m_fieldListColl_MouseMove(oUctl As clsUCTL, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    oUctl.BackColor = oUctl.ActiveColor
End Sub