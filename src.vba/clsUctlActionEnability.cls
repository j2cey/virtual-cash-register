Option Explicit

Implements IUCTLACTION

Private m_oUctl As clsUCTL

Private m_custStatus As formStatusType
Private m_vActionType As ctlActionType
Private m_vActionParameter As Variant
Private m_vActionPermissions As Variant


'*****************************************************************************************************************************
'    Interfaces Implementation
'*****************************************************************************************************************************

'    Public Properties
'******************************************************************

' Uctl
Public Property Set IUCTLACTION_Uctl(RHS As clsUCTL)
    Set m_oUctl = RHS
End Property

Public Property Get IUCTLACTION_Uctl() As clsUCTL
    Set IUCTLACTION_Uctl = m_oUctl
End Property

' Status
Public Property Let IUCTLACTION_Status(RHS As formStatusType)
    m_custStatus = RHS
End Property

Public Property Get IUCTLACTION_Status() As formStatusType
    IUCTLACTION_Status = m_custStatus
End Property

' ActionType
Public Property Let IUCTLACTION_ActionType(RHS As ctlActionType)
    m_vActionType = RHS
End Property

Public Property Get IUCTLACTION_ActionType() As ctlActionType
    IUCTLACTION_ActionType = m_vActionType
End Property

' ActionParameter
Public Property Let IUCTLACTION_ActionParameter(RHS As Variant)
    m_vActionParameter = RHS
End Property

Public Property Get IUCTLACTION_ActionParameter() As Variant
    IUCTLACTION_ActionParameter = m_vActionParameter
End Property

' ActionPermissions
Public Property Let IUCTLACTION_ActionPermissions(RHS As Variant)
    m_vActionPermissions = RHS
End Property

Public Property Get IUCTLACTION_ActionPermissions() As Variant
    IUCTLACTION_ActionPermissions = m_vActionPermissions
End Property


'    Public functions and subs
'******************************************************************
Public Sub IUCTLACTION_SetControl(oCtl As clsUCTL)
    Set m_oUctl = oCtl
End Sub

Private Sub IUCTLACTION_Execute()
    
    If Not IsEmpty(Me.IUCTLACTION_ActionPermissions) Then
        If Not loggedUser.Can(Me.IUCTLACTION_ActionPermissions) Then
            Exit Sub
        End If
    End If
    
    Select Case TypeName(IUCTLACTION_Uctl.ctl)
    Case "CheckBox"
        Dim chbx As MSForms.CheckBox, param As Boolean
        
        param = CBool(IUCTLACTION_ActionParameter)
        Set chbx = IUCTLACTION_Uctl.GetCheckBox
        chbx.Enabled = param
    Case "ComboBox"
        Dim cbx As MSForms.ComboBox, param As Boolean
        
        param = CBool(IUCTLACTION_ActionParameter)
        Set cbx = IUCTLACTION_Uctl.GetComboBox
        cbx.Enabled = param
    Case "ListBox"
        Dim lbx As MSForms.ListBox, param As Boolean
        
        param = CBool(IUCTLACTION_ActionParameter)
        Set lbx = IUCTLACTION_Uctl.GetListBox
        lbx.Enabled = param
    Case "OptionButton"
        Dim obtn As MSForms.OptionButton, param As Boolean
        
        param = CBool(IUCTLACTION_ActionParameter)
        Set obtn = IUCTLACTION_Uctl.GetOptionButton
        obtn.Enabled = param
    Case "SpinButton"
        Dim spn As MSForms.SpinButton, param As Boolean
        
        param = CBool(IUCTLACTION_ActionParameter)
        Set spn = IUCTLACTION_Uctl.GetSpinButton
        spn.Enabled = param
    Case "TextBox"
        Dim tbx As MSForms.TextBox, param As Boolean
        
        param = CBool(IUCTLACTION_ActionParameter)
        Set tbx = IUCTLACTION_Uctl.GetTextBox
        tbx.Enabled = param
    Case "Image"
        Dim img As MSForms.Image, param As Boolean
        
        param = CBool(IUCTLACTION_ActionParameter)
        Set img = IUCTLACTION_Uctl.GetImage
        img.Enabled = param
    Case "Label"
        Dim lbl As MSForms.Image, param As Boolean
        
        param = CBool(IUCTLACTION_ActionParameter)
        Set lbl = IUCTLACTION_Uctl.GetLabel
        lbl.Enabled = param
    Case "Frame"
        Dim frm As MSForms.Frame, param As Boolean
        
        param = CBool(IUCTLACTION_ActionParameter)
        Set frm = IUCTLACTION_Uctl.GetFrame
        frm.Enabled = param
    End Select
End Sub