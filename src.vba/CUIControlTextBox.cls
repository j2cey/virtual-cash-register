'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CUITextBox
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/08
' Purpose   : UI TextBox Control Wrapper Class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_control_parent As CUIControl
Private WithEvents m_txt As MSForms.TextBox


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CUITextBox
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : ControlParent
' Purpose       : The Inner Control bearing the UForm.XXX control
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/08    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set ControlParent(oControlParent As CUIControl)
    Set m_control_parent = oControlParent
End Property

Public Property Get ControlParent() As CUIControl
    Set User = m_control_parent
End Property


'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Subroutine    : Init
' Purpose       : Initialize the Object
' Arguments     : IUIControl        The Inner Control
' Returns       : Void
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/08    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(Optional oControlParent As CUIControlTextBox = Nothing)
    If Not oControlParent Is Nothing Then
        Set Me.ControlParent = oControlParent
    End If
End Sub


'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************


'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    
End Sub

