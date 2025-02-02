'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CFieldList
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/30
' Purpose   : The Data Access Recordable Object
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_field_list As Collection
Private m_fields_positions() As String
Private m_previous_field_index_picked As Long

Private Const RESET_FIELD_INDEX As Long = -1
Private Const MIN_FIELD_INDEX As Long = 0
Private Const FIELD_INDEX_INCREMENT As Long = 1


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CFieldList
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : PreviousFieldIndexPicked
' Purpose       : The Field s Id picked (using of the dedicated functions)
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/30    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let PreviousFieldIndexPicked(lngPreviousFieldIndexPicked As Long)
    m_previous_field_index_picked = lngPreviousFieldIndexPicked
End Property

Public Property Get PreviousFieldIndexPicked() As Long
    PreviousFieldIndexPicked = m_previous_field_index_picked
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : FieldList
' Purpose       : The Record s Fileds list
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/30    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set FieldList(oFieldList As Collection)
    Set m_field_list = oFieldList
End Property

Public Property Get FieldList() As Collection
    Set FieldList = m_field_list
End Property


'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     :
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/30      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init()
    
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
    Dim result As CResult, itemColl As Variant, oField As CField
    
    Set result = NewResut()
    result.IsOK = True
    
    If Me.FieldList Is Nothing Then
    Else
        For Each itemColl In Me.FieldList
            Dim currResult As CResult
            
            Set oField = itemColl
            Set currResult = oField.Reset()
            
            If Not currResult.IsOK Then
                Set result = currResult
                Exit For
            End If
            
        Next itemColl
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
    Dim result As CResult, itemColl As Variant, oField As CField
    
    Set result = NewResut()
    result.IsOK = True
    
    If Me.FieldList Is Nothing Then
    Else
        For Each itemColl In Me.FieldList
            
            Set oField = itemColl
            oField.Clear
            
        Next itemColl
    End If
    
    Set Clear = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : AddField
'   Purpose     : Add Field to the fields list
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/22  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddField(ByVal oField As CField, Optional ByVal lngAtIndex = -1) As CResult
    Dim result As CResult, lngFieldIndex As Long
    
    Set result = NewResult("Addfield", True)
    
    If lngAtIndex = -1 Then
        lngFieldIndex = Me.FieldList.Count + 1
    Else
        lngFieldIndex = lngAtIndex
    End If
    
    Me.FieldList.Add oField, oField.Guid
    AddFieldPosition oField.Guid
    
    Set AddField = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : ResetFieldsPicked
'   Purpose     : Reset the amount of Fields Picked
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/23  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ResetFieldsPicked() As CResult
    Dim result As CResult
    
    Set result = NewResult("ResetFieldsPicked", True)
    Me.PreviousFieldIndexPicked = RESET_FIELD_INDEX
    
    Set ResetFieldsPicked = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetFieldByName
'   Purpose     : Get the Field having the given name
'   Arguments   : strFieldName  The name of the field
'
'   Returns     : CField        The Field | Nothing (if not found or anything wrong)
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/31  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetFieldByName(ByVal strFieldName As String) As CField
    Dim result As CResult, itemColl As Variant, oCurrField As CField
    
    If Me.FieldList Is Nothing Then
    Else
        For Each itemColl In Me.FieldList
            
            Set oCurrField = itemColl
            
            If Not oCurrField Is Nothing Then
                If oCurrField.Name = strFieldName Then
                    Set GetFieldByName = oCurrField
                    Exit For
                End If
            End If
            
        Next itemColl
    End If
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetNextField
'   Purpose     : Get Next Field, or Field at Index (if given)
'   Arguments   : lngAtIndex    Index of field
'
'   Returns     : CField        The Field | Nothing (if not found or anything wrong)
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/23  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNextField(Optional ByVal lngAtIndex As Long = -1) As CField
    Dim fieldKey As String
    
    fieldKey = GetNextFieldKey(lngAtIndex)
    
    If fieldKey = "" Then
        Set GetNextField = Nothing
    Else
        Set GetNextField = Me.FieldList(fieldKey)
    End If
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetNextFieldCriteria
'   Purpose     : Get Next Search Criteria Field, or Field at Index (if given)
'   Arguments   : lngAtIndex    Index of field
'
'   Returns     : CField        The Field | Nothing (if not found or anything wrong)
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/23  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNextFieldCriteria(Optional ByVal lngAtIndex As Long = -1) As CField
    Dim field As CField
    
    If Not lngAtIndex = -1 Then
        Set field = GetNextField(lngAtIndex)
        If field Is Nothing Then
            Set GetNextFieldCriteria = field
        Else
            If field.IsSearchCriteria Then
                Set GetNextFieldCriteria = field
            Else
                Set GetNextFieldCriteria = Nothing
            End If
        End If
    Else
        
        Dim done As Boolean
        
        done = False
        
        Do
            Set field = GetNextField(lngAtIndex)
            
            If field Is Nothing Then
                Set GetNextFieldCriteria = Nothing
                done = True
                Exit Do
            End If
            
            If field.IsSearchCriteria Then
                Set GetNextFieldCriteria = field
                done = True
                Exit Do
            End If
            
        Loop While (Not done)
        
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetNextFieldSelectable
'   Purpose     : Get Next Selectable Field, or Field at Index (if given)
'   Arguments   : lngAtIndex    Index of field
'
'   Returns     : CField        The Field | Nothing (if not found or anything wrong)
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/23  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNextFieldSelectable(Optional ByVal lngAtIndex As Long = -1) As CField
    Dim field As CField
    
    If lngAtIndex = -1 Then
        Set field = GetNextField(lngAtIndex)
        If field Is Nothing Then
            Set GetNextFieldSelectable = field
        Else
            If field.IsSelectable Then
                Set GetNextFieldSelectable = field
            Else
                Set GetNextFieldSelectable = Nothing
            End If
        End If
    Else
        
        Dim done As Boolean
        
        done = False
        
        Do
            Set field = GetNextField(lngAtIndex)
            
            If field Is Nothing Then
                Set GetNextFieldSelectable = Nothing
                done = True
                Exit Do
            End If
            
            If field.IsSelectable Then
                Set GetNextFieldSelectable = field
                done = True
                Exit Do
            End If
            
        Loop While (Not done)
        
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetNextFieldSavable
'   Purpose     : Get Next Savable Field, or Field at Index (if given)
'   Arguments   : lngAtIndex    Index of field
'
'   Returns     : CField        The Field | Nothing (if not found or anything wrong)
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/23  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNextFieldSavable(Optional ByVal lngAtIndex As Long = -1) As CField
    Dim field As CField
    
    If lngAtIndex = -1 Then
        Set field = GetNextField(lngAtIndex)
        If field Is Nothing Then
            Set GetNextFieldSavable = field
        Else
            If field.IsSavable Then
                Set GetNextFieldSavable = field
            Else
                Set GetNextFieldSavable = Nothing
            End If
        End If
    Else
        
        Dim done As Boolean
        
        done = False
        
        Do
            Set field = GetNextField(lngAtIndex)
            
            If field Is Nothing Then
                Set GetNextFieldSavable = Nothing
                done = True
                Exit Do
            End If
            
            If field.IsSavable Then
                Set GetNextFieldSavable = field
                done = True
                Exit Do
            End If
            
        Loop While (Not done)
        
    End If
End Function

Public Function LastFieldReached() As Boolean
    LastFieldReached = ((Me.PreviousFieldIndexPicked + FIELD_INDEX_INCREMENT) >= Me.FieldList.Count)
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
Public Function ValidateFields() As CResult
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
                Set currResult = oField.ValidateField()
            End If
            
            ' Set the current result values to Function result
            result.IsOK = currResult.IsOK
            result.Message = result.Message & " - " & currResult.Message
            
        Next itemColl
    End If
    
    Set ValidateFields = result
End Function


'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : AddFieldPosition
'   Purpose     : Add new Field Position to the position list
'   Arguments   : strElem       The New Element (GUID)
'                 lngIndex      The position to insert
'
'   Returns     : Long          The new list size
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/23  Jude Parfait    Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddFieldPosition(ByVal strElem As String, Optional lngIndex As Long = -1) As Long
    m_fields_positions = AddToStringArray(m_fields_positions, strElem, lngIndex)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetNextFieldKey
'   Purpose     : Get the Next Field Key to be picked
'   Arguments   : lngAtIndex    The Index of the Key (if given)
'
'   Returns     : String        The key | "" (if not found or anything wrong)
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/23  Jude Parfait    Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetNextFieldKey(Optional ByVal lngAtIndex = -1) As String
    If lngAtIndex = -1 Then
        Dim currIndex As Long
        
        ' We increment the previous field index picked
        currIndex = IncrementPreviousFieldIndexPicked()
        
        ' If the max Index is reached
        If currIndex >= Me.FieldList.Count Then
            GetNextFieldKey = ""
            Exit Function
        End If
        
        GetNextFieldKey = m_fields_positions(currIndex)
        Me.PreviousFieldIndexPicked = MIN_FIELD_INDEX
    Else
        GetNextFieldKey = m_fields_positions(currIndex)
    End If
End Function

Private Function IncrementPreviousFieldIndexPicked() As Long
    If Me.PreviousFieldIndexPicked = RESET_FIELD_INDEX Then
        IncrementPreviousFieldIndexPicked = MIN_FIELD_INDEX
        'Me.PreviousFieldIndexPicked = MIN_FIELD_INDEX
    Else
        IncrementPreviousFieldIndexPicked = Me.PreviousFieldIndexPicked + FIELD_INDEX_INCREMENT
        'Me.PreviousFieldIndexPicked = Me.PreviousFieldIndexPicked + FIELD_INDEX_INCREMENT
    End If
End Function


'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    Set m_field_list = New Collection
    m_previous_field_index_picked = RESET_FIELD_INDEX
End Sub

Private Sub Class_Terminate()
    Set m_field_list = Nothing
End Sub
