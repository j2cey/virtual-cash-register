'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CField
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/12
' Purpose   : The Record Field BLL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_guid As String
Private m_name As String
Private m_label As String
Private m_value As Variant
Private m_previous_value As Variant

Private m_is_required As Boolean
Private m_is_search_criteria As Boolean
Private m_search_criteria_must_match_exact As Boolean
Private m_is_selectable As Boolean
Private m_is_savable As Boolean

Private m_fields_list As Collection
Private m_last_result As CResult


Implements IRecordableBL


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CField
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Guid
'   Purpose       : Get the unique identifier of the object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/23      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Guid() As String
    Guid = m_guid
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Name
'   Purpose       : The field Name
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Name(strName As String)
    m_name = strName
End Property

Public Property Get Name() As String
    Name = m_name
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Label
'   Purpose       : The field Label (display name)
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Label(strLabel As String)
    m_label = strLabel
End Property

Public Property Get Label() As String
    Label = m_label
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Value
'   Purpose       : The field Value
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let value(vrnValue As Variant)
    Me.Reset
    m_previous_value = m_value
    
    m_value = vrnValue
End Property

Public Property Get value() As Variant
    value = m_value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : PreviousValue
'   Purpose       : The field Previous Value
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Property Let PreviousValue(vrnPreviousValue As Variant)
'    m_previous_value = vrnPreviousValue
'End Property

Public Property Get PreviousValue() As Variant
    PreviousValue = m_previous_value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IsRequired
'   Purpose       : Determine if this field is required
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsRequired(blnIsRequired As Boolean)
    m_is_required = blnIsRequired
End Property

Public Property Get IsRequired() As Boolean
    IsRequired = m_is_required
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IsSearchCriteria
'   Purpose       : Determine if this field is search criteria
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsSearchCriteria(blnIsSearchCriteria As Boolean)
    m_is_search_criteria = blnIsSearchCriteria
End Property

Public Property Get IsSearchCriteria() As Boolean
    IsSearchCriteria = m_is_search_criteria
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : SearchCriteriaMustMatchExact
'   Purpose       : Determine if the search criteria in this field must match exactly
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/24      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let SearchCriteriaMustMatchExact(blnSearchCriteriaMustMatchExact As Boolean)
    m_search_criteria_must_match_exact = blnSearchCriteriaMustMatchExact
End Property

Public Property Get SearchCriteriaMustMatchExact() As Boolean
    SearchCriteriaMustMatchExact = m_search_criteria_must_match_exact
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IsSelectable
'   Purpose       : Determine if this field have to be selected
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsSelectable(blnIsSelectable As Boolean)
    m_is_selectable = blnIsSelectable
End Property

Public Property Get IsSelectable() As Boolean
    IsSelectable = m_is_selectable
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IsSavable
'   Purpose       : Determine if this field have to be saved
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsSavable(blnIsSavable As Boolean)
    m_is_savable = blnIsSavable
End Property

Public Property Get IsSavable() As Boolean
    IsSelectable = m_is_savable
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : FieldsList
' Purpose       : The Record Fields List
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/18    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set FieldsList(oFieldsList As Collection)
    Set m_fields_list = oFieldsList
End Property

Public Property Get FieldsList() As Collection
    Set FieldsList = m_fields_list
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : LastResult
'   Purpose     : Get / Set The Last Result
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LastResult() As CResult
    Set LastResult = m_last_result
End Property

Public Property Set LastResult(oLastResult As CResult)
    Set m_last_result = oLastResult
End Property


'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : strName             The field Name
'                   strLabel            The Performing User
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(strName As String, strLabel As String, Optional vrnValue As Variant = Null)
    
    Me.Name = strName
    Me.Label = strLabel
    
    If Not vrnValue = Null Then
        Me.value = vrnValue
    End If
End Sub

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
    Dim result As CResult
    
    Set result = NewResut()
    result.IsOK = True
    
    If Me.PreviousValue = Null Then
       result.Message = "Pas de valeure précédente"
    Else
        Me.value = Me.PreviousValue
        result.Message = "Valeur du Champs " & Me.Label & " réinitialisée avec succès !"
    End If
    
    Reset = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine  : Clear
'   Purpose     : Clear the Field Value
'   Arguments   :
'
'   Returns     : Void
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
    Me.value = Null
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IsEmpty
'   Purpose     : Determine if this filed is empty
'   Arguments   :
'
'   Returns     : Boolean
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEmpty() As Boolean
    
    If Me.value = Null Then
       IsEmpty = True
    Else
        IsEmpty = False
    End If
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : ValidateField
'   Purpose     : Determine if this filed s value is valid to be used for saving
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateField() As CResult
    Dim result As CResult
    
    Set result = NewResut("Validate Field for Saving")
    result.IsOK = True
    
    If Me.IsRequired Then
       result.IsOK = (Not Me.IsEmpty) And (Me.value <> "")
    End If
    
    ValidateField = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : ValidateFieldCriteria
'   Purpose     : Determine if this filed s value is valid to be used as criteria
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateFieldCriteria() As CResult
    Dim result As CResult
    
    Set result = NewResut("Validate Field Criteria")
    result.IsOK = True
    
    If Not Me.IsEmpty Then
       result.IsOK = Me.value <> ""
    End If
    
    ValidateFieldCriteria = result
End Function



'***************************************************************************************************************************************************************
'   Interfaces Implementation
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IRecordableBL_LastResult
'   Purpose     : Get / Set The Last Result
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get IRecordableBL_LastResult() As CResult
    Set IRecordableBL_LastResult = Me.LastResult
End Property

Public Property Set IRecordableBL_LastResult(RHS As CResult)
    Set Me.LastResult = RHS
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IRecordableBL_FieldsList
'   Purpose     : Get / Set Fields List
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get IRecordableBL_FieldsList() As Collection
    Set IRecordableBL_FieldsList = Me.FieldsList
End Property

Public Property Set IRecordableBL_FieldsList(RHS As Collection)
    Set Me.FieldsList = RHS
End Property


'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    m_guid = CreateGUID
    m_previous_value = Null
    Set m_fields_list = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_fields_list = Nothing
End Sub
