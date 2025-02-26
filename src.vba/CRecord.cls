'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CRecord
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/12
' Purpose   : The Record BLL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_user As CModelUser

Private m_record_id As Long
Private Const ID_KEY As String = "Id"
Private Const ID_LABEL As String = "ID"

Private m_data_access As CDataAccess
Private m_field_list As CFieldList
Private m_last_result As CResult

Private m_uicontrol As CUIControl

' TODO: remove m_value ?
'Private m_value As Variant


Private Const MODULE_NAME As String = "Record"
Private Const DEFAULT_ID As Long = 0


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CRecord
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : User
'   Purpose       : The performing user
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set User(oUser As CModelUser)
    Set m_user = oUser
End Property

Public Property Get User() As CModelUser
    Set User = m_user
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : DataAccess
'   Purpose       : The Data Access object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set DataAccess(oDataAccess As CDataAccess)
    Set m_data_access = oDataAccess
End Property

Public Property Get DataAccess() As CDataAccess
    Set DataAccess = m_data_access
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : RecordId
' Purpose       : The Record s Id
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/31    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let RecordId(lngRecordId As Long)
    Dim oField As CField
    
    Set oField = Me.FieldList.GetFieldByName(ID_KEY)
    
    If Not oField Is Nothing Then
        oField.SetValue lngRecordId
    Else
        m_record_id = lngRecordId
    End If
End Property

Public Property Get RecordId() As Long
    Dim oField As CField
    
    Set oField = Me.FieldList.GetFieldByName(ID_KEY)
    
    If Not oField Is Nothing Then
        RecordId = CLng(oField.FieldValue.GetValue)
    Else
        RecordId = m_record_id
    End If
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Value
'   Purpose       : The Value of the Record
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Property Let value(vrnValue As Variant)
'    Let m_value = vrnValue
'End Property

'Public Property Get value() As Variant
'    value = m_value
'End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : FieldList
' Purpose       : The Record Fields List
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/08    Jude Parfait        Created
' 2025/01/30    Jude Parfait        Change type to CFieldList
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set FieldList(oFieldList As CFieldList)
    Set m_field_list = oFieldList
End Property

Public Property Get FieldList() As CFieldList
    Set FieldList = m_field_list
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : FieldListCount
' Purpose       : Get the amount of Fields in the Record
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/14    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get FieldListCount() As Long
    FieldListCount = Me.FieldList.FieldList.count
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : LastResult
'   Purpose     : Get / Set The Last Result
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LastResult() As CResult
    Set LastResult = m_last_result
End Property

Public Property Set LastResult(oLastResult As CResult)
    Set m_last_result = oLastResult
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : UIControl
'   Purpose       : The related UI Control
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set UIControl(oUIControl As CUIControl)
    Set m_uicontrol = oUIControl
End Property

Public Property Get UIControl() As CUIControl
    Set UIControl = m_uicontrol
End Property



'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : oRecordDA           The Data Access object
'                   oUser               The Performing User
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(Optional oDataAccess As CDataAccess = Nothing, Optional oUser As CModelUser = Nothing, Optional ByVal lngRecordId As Long = -1)
    
    If Not oDataAccess Is Nothing Then
        Set Me.DataAccess = oDataAccess
        'Set Me.DataAccess.Record = Me
    End If
    
    If Not oUser Is Nothing Then
        Set Me.User = oUser
    End If
    
    If Not lngRecordId = -1 Then
        Me.RecordId = lngRecordId
    End If
End Sub


Public Function SetIdField(Optional ByVal lngId As Long = -1) As CField
    Dim oField As CField, lngTest As Long
    
    Set Me.LastResult = SetNewResult("SetIdField", MODULE_NAME)
    
    Set oField = NewField(NewFieldValueInteger, ID_KEY, ID_LABEL, lngId, ID_KEY, ID_KEY)
    
    AddSubResult MODULE_NAME, "SetIdField", Me.LastResult, Me.FieldList.AddField(oField)
    
    lngTest = Me.FieldListCount
    
    Set SetIdField = oField
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : Save
'   Purpose     : Save the Record
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Save() As CResult
    Dim result As CResult
    
    Set result = Me.CheckRecordDA()
    
    If result.IsOK Then
        Set result = Me.RecordDA.Save()
    End If
    
    Set Save = result
End Function

Public Sub FillControl()
    If Not Me.UIControl Is Nothing Then
        Me.UIControl.FillFromRecord Me
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : Read
'   Purpose     : Read the record from Data Access (using search criteria, if any)
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Read() As CResult
    Dim result As CResult, vrnValue As Variant
    
    Set result = Me.CheckRecordDA()
    
    If result.IsOK Then
        Set result = Me.RecordDA.Read(vrnValue)
        Me.Value = vrnValue
        Me.FillFields
    End If
    
    Set Read = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : Delete
'   Purpose     : Delete the record from Data Access
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Delete() As CResult
    Dim result As CResult
    
    Set result = Me.CheckRecordDA()
    
    If result.IsOK Then
        Set result = Me.RecordDA.Delete()
    End If
    
    Set Delete = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : Reset
'   Purpose     : Set value to previous value
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Reset() As CResult
    Dim result As CResult, itemColl As Variant, oField As clsUCTL
    
    If Me.FieldList Is Nothing Then
    Else
        Set result = Me.FieldList.Reset
    End If
    
    Set Reset = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : Clear
'   Purpose     : Clear the Field Value
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Clear() As CResult
    Dim result As CResult, itemColl As Variant, oField As clsUCTL
    
    Set result = NewResut()
    result.IsOK = True
    
    If Me.FieldList Is Nothing Then
    Else
        Me.FieldList.Clear
    End If
    
    Set Clear = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IsNew
'   Purpose     : Determine whether the object is new
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsNew() As Boolean
    Dim result As Boolean
    
    result = True
    
    If Me.RecordId = DEFAULT_ID Then
       result = True
    Else
        result = False
    End If
    
    IsNew = result
End Function



'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : CheckRecordDA
'   Purpose     : Check if the RecordDA is OK
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckRecordDA() As CResult
    Dim result As CResult
    
    Set result = NewResut()
    
    If Me.RecordDA Is Nothing Then
       result.IsOK = False
       result.Message = "Aucun Objet d'Accès aux Données"
    Else
        result.IsOK = True
    End If
    
    Set CheckRecordDA = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : ValidateFields
'   Purpose     : Validate all (required) Fields
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateFields() As CResult
    Dim result As CResult, itemColl As Variant, oField As CField
    
    If Me.FieldList Is Nothing Then
    Else
        Set result = Me.FieldList.ValidateFields
    End If
    
    Set ValidateFields = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : FillFields
'   Purpose     : Fill Fields
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FillFields() As CResult
    Dim result As CResult, itemColl As Variant, oField As CField
    
    Set result = NewResut()
    result.IsOK = True
    
    If Me.FieldList Is Nothing Then
    Else
        For Each itemColl In Me.FieldList
            Dim currResult As CResult
            
            Set oField = itemColl
            Set currResult = oField.Clear()
            
            If currResult.IsOK Then
                oField.Value
            End If
            
            ' Set the current result values to Function result
            result.IsOK = currResult.IsOK
            result.Message = result.Message & " - " & currResult.Message
            
        Next itemColl
    End If
    
    Set FillFields = result
End Function


'***************************************************************************************************************************************************************
'   Interfaces Implementation
'***************************************************************************************************************************************************************





'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    m_record_id = DEFAULT_ID
    Set m_field_list = NewFieldList()
End Sub

Private Sub Class_Terminate()
    Set m_field_list = Nothing
End Sub
