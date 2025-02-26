'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CUIControl
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/08
' Purpose   : UI Control Class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_control_inner As IUIControl
Private m_user As CModelUser

Private m_business_logic As CBusinessLogic


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CUIControl
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : ControlInner
'   Purpose       : The Inner Control bearing the UForm.XXX control
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/08      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set ControlInner(oControlInner As IUIControl)
    Set m_control_inner = oControlInner
End Property

Public Property Get ControlInner() As IUIControl
    Set User = m_control_inner
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : Name
' Purpose       : The UI Control s name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/11    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Name(sName As String)
    Me.ControlInner.Name = sName
End Property

Public Property Get Name() As String
    Name = Me.ControlInner.Name
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : Value
' Purpose       : The UI Control s value
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/11    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Value(varValue As Variant)
    Me.ControlInner.Value = varValue
End Property

Public Property Get Value() As Variant
    Value = Me.ControlInner.Value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : BusinessLogic
' Purpose       : The UI Business Logic
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/10    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set BusinessLogic(oBusinessLogic As CBusinessLogic)
    Set m_business_logic = oBusinessLogic
End Property

Public Property Get BusinessLogic() As CBusinessLogic
    Set BusinessLogic = m_business_logic
End Property



'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : IUIControl        The Inner Control
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/08      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(oControlInner As IUIControl)
    Set Me.ControlInner = oControlInner
End Sub
`
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Clear
' Purpose       : Clear the control
' Returns       : Boolean           True if clear OK, False if something goes wrong
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/11    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Clear() As Boolean
    Me.ControlInner.Clear
End Function

Public Sub ExecuteAction(oUIControlAction As IUIControlAction)
    If Not IsEmpty(oUIControlAction.ActionPermissions) Then
        If Not loggedUser.Can(Me.IUCTLACTION_ActionPermissions) Then
            Exit Sub
        End If
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FillFromField
' Purpose       : Fill the Control From Field
' Returns       : CResult
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/10    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromField(ByVal oField As CField) As CResult
    Set FillFromField = Me.ControlInner.FillFromField(oField)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FillFromRecord
' Purpose       : Fill the Control From Record
' Returns       : CResult
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/10    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromRecord(ByVal oRecord As CRecord) As CResult
    Set FillFromRecord = Me.ControlInner.FillFromRecord(oRecord)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FillFromRecord
' Purpose       : Fill the Control From Record-List
' Returns       : CResult
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/10    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromRecordList(ByVal oRecordList As CRecordList) As CResult
    Set FillFromRecordList = Me.ControlInner.FillFromRecordList(oRecord)
End Function

'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************


'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    
End Sub

