'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CField
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/12
' Purpose   : The Record Field BLL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_guid As String
Private m_value As IFieldValue
Private m_previous_value As IFieldValue
Private m_name As String
Private m_offset_start As Long
Private m_offset_end As Long
Private m_label As String

Private m_is_required As Boolean
Private m_is_search_criteria As Boolean
Private m_search_criteria_must_match_exact As Boolean
Private m_is_selectable As Boolean
Private m_is_savable As Boolean
Private m_is_input_parameter As Boolean
Private m_is_output_parameter As Boolean


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
'   Property      : FieldValue
'   Purpose       : The field value
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set FieldValue(oFieldValue As IFieldValue)
    Set m_value = oFieldValue
End Property

Public Property Get FieldValue() As IFieldValue
    Set FieldValue = m_value
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
'   Property      : OffsetStart
'   Purpose       : The field offset start point, if any
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let OffsetStart(lngOffsetStart As Long)
    m_offset_start = lngOffsetStart
End Property

Public Property Get OffsetStart() As Long
    OffsetStart = m_offset_start
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : OffsetEnd
'   Purpose       : The field offset end point, if any
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let OffsetEnd(lngOffsetEnd As Long)
    m_offset_end = lngOffsetEnd
End Property

Public Property Get OffsetEnd() As Long
    OffsetEnd = m_offset_end
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

Public Property Get PreviousValue() As IFieldValue
    Set PreviousValue = m_previous_value
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
'   Property      : IsInputParameter
'   Purpose       : Determine if this field an INPUT parameter
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsInputParameter(blnIsInputParameter As Boolean)
    m_is_input_parameter = blnIsInputParameter
End Property

Public Property Get IsInputParameter() As Boolean
    IsInputParameter = m_is_input_parameter
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IsOutputParameter
'   Purpose       : Determine if this field an OUTPUT parameter
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsOutputParameter(blnIsOutputParameter As Boolean)
    m_is_output_parameter = blnIsOutputParameter
End Property

Public Property Get IsOutputParameter() As Boolean
    IsOutputParameter = m_is_output_parameter
End Property


'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : oFieldValue         The field value
'                   strName             The field Name
'                   strLabel            The Performing User
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal oFieldValue As IFieldValue, strName As String, strLabel As String, Optional vrnValue As Variant = Null)
    
    Set Me.FieldValue = oFieldValue
    Set oFieldValue.UpperField = Me
    
    Me.Name = strName
    Me.Label = strLabel
    
    If Not vrnValue = Null Then
        Me.SetValue vrnValue
    End If
End Sub

Public Sub SetValue(vrnValue As Variant)
    Me.Reset
    Set m_previous_value = Me.FieldValue
    
    Me.FieldValue.SetValue (vrnValue)
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
    
    Set result = NewResult("Reset Field", True)
    
    If Me.PreviousValue Is Nothing Then
       result.Message = "Pas de valeure précédente"
    Else
        Me.FieldValue.SetValue (Me.PreviousValue.GetValue)
        result.Message = "Valeur du Champs " & Me.Label & " réinitialisée avec succès !"
    End If
    
    Set Reset = result
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
    Me.FieldValue.ClearValue
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
    
    IsEmpty = Me.FieldValue.IsEmpty
    
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
       result.IsOK = (Not Me.IsEmpty)
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
       result.IsOK = (Not Me.IsEmpty)
    End If
    
    ValidateFieldCriteria = result
End Function

Public Function SetSelectable(Optional ByVal blnValue As Boolean = True) As CField
    Me.IsSelectable = blnValue
    
    Set SetSelectable = Me
End Function

Public Function SetOffsetStart(ByVal lngOffsetStart As Long) As CField
    Me.OffsetStart = lngOffsetStart
    
    Set SetOffsetStart = Me
End Function

Public Function SetOffsetEnd(ByVal lngOffsetEnd As Long) As CField
    Me.OffsetEnd = lngOffsetEnd
    
    Set SetOffsetEnd = Me
End Function

Public Function ResetCode() As CField
    m_guid = CreateGUID
    
    Set ResetCode = Me
End Function




'***************************************************************************************************************************************************************
'   Interfaces Implementation
'***************************************************************************************************************************************************************




'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    m_guid = CreateGUID
    'Set m_previous_value = New IFieldValue
End Sub

Private Sub Class_Terminate()
End Sub
