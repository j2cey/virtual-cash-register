Option Explicit

Private m_oUctl As clsUCTL

Private m_bIsRequired As Boolean
Private m_bIsSavable As Boolean
Private m_iscancellable As Boolean

Private m_bIsCriteria As Boolean
Private m_bIsResultField As Boolean
Private m_bIsPasswordField As Boolean

Private m_sName As String
Private m_sTitle As String
Private m_lIndex As Long
Private m_vValue As Variant
Private m_vCriteriaValue As Variant


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As clsDbField
    Set Self = Me
End Property

' Uctl
Public Property Set Uctl(oUctl As clsUCTL)
    Set m_oUctl = oUctl
    If oUctl.IsPassword Then
        Me.IsPasswordField = True
    End If
End Property

Public Property Get Uctl() As clsUCTL
    Set Uctl = m_oUctl
End Property

' Name
Public Property Let Name(sName As String)
    m_sName = sName
End Property

Public Property Get Name() As String
    Name = m_sName
End Property

' Title
Public Property Let Title(sTitle As String)
    m_sTitle = sTitle
End Property

Public Property Get Title() As String
    Title = m_sTitle
End Property

' IsRequired
Public Property Let IsRequired(bIsRequired As Boolean)
    m_bIsRequired = bIsRequired
    
    If Not Me.Uctl Is Nothing Then
        Me.Uctl.FormatCtlLabel
    End If
End Property

Public Property Get IsRequired() As Boolean
    IsRequired = m_bIsRequired
End Property

' IsSavable
Public Property Let IsSavable(bIsSavable As Boolean)
    m_bIsSavable = bIsSavable
End Property

Public Property Get IsSavable() As Boolean
    IsSavable = m_bIsSavable
End Property

' IsCancellable
Public Property Let IsCancellable(bIsCancellable As Boolean)
    m_iscancellable = bIsCancellable
End Property

Public Property Get IsCancellable() As Boolean
    IsCancellable = m_iscancellable
End Property

' Index
Public Property Let Index(lIndex As Long)
    m_lIndex = lIndex
End Property

Public Property Get Index() As Long
    Index = m_lIndex
End Property

' IsCriteria
Public Property Let IsCriteria(bIsCriteria As Boolean)
    m_bIsCriteria = bIsCriteria
End Property

Public Property Get IsCriteria() As Boolean
    IsCriteria = m_bIsCriteria
End Property

' IsResultField
Public Property Let IsResultField(bIsResultField As Boolean)
    m_bIsResultField = bIsResultField
End Property

Public Property Get IsResultField() As Boolean
    IsResultField = m_bIsResultField
End Property

' CriteriaValue
Public Property Let CriteriaValue(vCriteriaValue As Variant)
    m_vCriteriaValue = vCriteriaValue
End Property

Public Property Get CriteriaValue() As Variant
    CriteriaValue = m_vCriteriaValue
End Property

' m_vValue
Public Property Let value(vValue As Variant)
    m_vValue = vValue
End Property

Public Property Get value() As Variant
    If Me.IsPasswordField Then
        value = encrypter.encrypt(CStr(m_vValue))
    Else
        value = m_vValue
    End If
End Property

' IsPasswordField
Public Property Let IsPasswordField(bIsPasswordField As Boolean)
    m_bIsPasswordField = bIsPasswordField
End Property

Public Property Get IsPasswordField() As Boolean
    IsPasswordField = m_bIsPasswordField
End Property


Public Sub Init(sName As String, sTitle As String, lIndex As Long, bIsSavable As Boolean, bIsCriteria As Boolean, bIsResultField As Boolean, _
bIsRequired As Boolean, bIsCancellable As Boolean, Optional vValue As Variant, Optional vCriteriaValue As Variant)
    
    Me.Name = sName
    Me.Title = sTitle
    Me.Index = lIndex
    Me.IsSavable = bIsSavable
    Me.IsCriteria = bIsCriteria
    Me.IsResultField = bIsResultField
    Me.IsRequired = bIsRequired
    Me.IsCancellable = bIsCancellable
    Me.value = vValue
    Me.CriteriaValue = vCriteriaValue
    
End Sub
