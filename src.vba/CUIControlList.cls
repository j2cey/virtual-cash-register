Option Explicit


Private m_listCtlsColl As Collection

'*****************************************************************************************************************************
'   Public Events
'*****************************************************************************************************************************

Public Event UCtlClick(oUctl As CUIControl)
Public Event UCtlMouseMove(oUctl As CUIControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CUIControlList
    Set Self = Me
End Property

' CriterionList
Public Property Get ListCtls() As Collection
    Set ListCtls = m_listCtlsColl
End Property

Public Property Get TestGet() As String
    TestGet = "Test from CUIControlList"
End Property


'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub AddCtlByControl(oCtl As MSForms.Control)
    Me.AddCtlByUCTL NewUCTL(oCtl)
End Sub

Public Sub AddCtlByUCTL(oCtl As CUIControl)
    
    ' Add the contro object to the collection
    m_listCtlsColl.Add oCtl, oCtl.Name
    ' Set this UCTLList as ParentColl of the just added control
    Set oCtl.ParentColl = Me
    
End Sub

Public Function GetCTL(sCtlName As String) As CUIControl
    If m_listCtlsColl Is Nothing Then
        Set GetCTL = Nothing
    Else
        Set GetCTL = m_listCtlsColl.Item(sCtlName)
    End If
End Function

Public Sub ImageBtnApplyNormalSizeAll()
    'Make all images button controls normal size
    Dim itemColl As Variant, imgname As Variant, Uctl As CUIControl, ctl As MSForms.Control
    
    If m_listCtlsColl Is Nothing Then
    Else
        For Each itemColl In m_listCtlsColl
            Set Uctl = itemColl
            
            If Uctl.IsTypeOf("Image") Then
                Uctl.ApplyNormalSize
            End If
        Next itemColl
    End If
End Sub

Public Function GetDbFieldsSavable() As Collection
    Dim subList As Collection, itemColl As Variant, oUctl As CUIControl
    
    Set subList = New Collection
    For Each itemColl In Me.ListCtls
        Set oUctl = itemColl
        
        If oUctl.field.IsSavable Then
            subList.Add oUctl.field, oUctl.field.Name
        End If
    Next itemColl
    
    Set GetDbFieldsSavable = subList
End Function

Public Function GetSavableUCtls() As CUIControlList
    Dim subList As CUIControlList, itemColl As Variant, oUctl As CUIControl
    
    Set subList = New CUIControlList
    For Each itemColl In Me.ListCtls
        Set oUctl = itemColl
        
        If oUctl.field.IsSavable Then
            subList.AddCtlByUCTL oUctl
        End If
    Next itemColl
    
    Set GetSavableUCtls = subList
End Function

Public Function GetDbFieldsCriterion() As Collection
    Dim subList As Collection, itemColl As Variant, oUctl As CUIControl
    
    Set subList = New Collection
    For Each itemColl In Me.ListCtls
        Set oUctl = itemColl
        
        If oUctl.field.IsCriteria Then
            subList.Add oUctl.field, oUctl.field.Name
        End If
    Next itemColl
    
    Set GetDbFieldsCriterion = subList
End Function

Public Function GetCriterionUCtls() As CUIControlList
    Dim subList As CUIControlList, itemColl As Variant, oUctl As CUIControl
    
    Set subList = New CUIControlList
    For Each itemColl In Me.ListCtls
        Set oUctl = itemColl
        
        If oUctl.field.IsCriteria Then
            subList.AddCtlByUCTL oUctl
        End If
    Next itemColl
    
    Set GetCriterionUCtls = subList
End Function

' Return Collect of clsDbField
Public Function GetDbFieldsResult() As Collection
    Dim subList As Collection, itemColl As Variant, oUctl As CUIControl
    
    Set subList = New Collection
    For Each itemColl In Me.ListCtls
        Set oUctl = itemColl
        
        If oUctl.field.IsResultField Then
            subList.Add oUctl.field, oUctl.field.Name
        End If
    Next itemColl
    
    Set GetDbFieldsResult = subList
End Function

Public Function GetResultUCtls() As CUIControlList
    Dim subList As CUIControlList, itemColl As Variant, oUctl As CUIControl
    
    Set subList = New CUIControlList
    For Each itemColl In Me.ListCtls
        Set oUctl = itemColl
        
        If oUctl.field.IsResultField Then
            subList.AddCtlByUCTL oUctl
        End If
    Next itemColl
    
    Set GetResultUCtls = subList
End Function

Public Sub RaiseMouseMoveEvent(oUctl As CUIControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent UCtlMouseMove(oUctl, Button, Shift, x, y)
End Sub


'*****************************************************************************************************************************
'    Private functions and subs
'*****************************************************************************************************************************

Private Function GetNewDbFieldFromUCTl(oUctl As CUIControl) As clsDbField
    Set GetNewDbFieldFromUCTl = NewDbField(oUctl.FieldName, oUctl.FieldTitle, oUctl.FieldIndex, oUctl.IsSavable, oUctl.IsCriteria, oUctl.IsResultField, oUctl.IsRequired, oUctl.IsCancellable, oUctl.Text, oUctl.Text)
End Function



'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    Set m_listCtlsColl = New Collection
End Sub