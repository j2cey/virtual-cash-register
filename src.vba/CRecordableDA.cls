'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CRecordableDA
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/12
' Purpose   : The Data Access Recordable Object
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_guid As String
Private m_user As CUser
Private m_data_source As IDataSource

Private m_record_id As Long
Private m_record_table As String

Private m_fields_list As Collection
Private m_fields_positions() As String
Private m_previous_field_index_picked As Long

Private Const DEFAULT_ID As Long = 0
Private Const MIN_FIELD_INDEX As Long = 0



'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CRecordableDA
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
'   Property      : User
'   Purpose       : The performing user
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set User(oUser As CUser)
    Set m_user = oUser
End Property

Public Property Get User() As CUser
    Set User = m_user
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : DataSource
'   Purpose       : The Data Source
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set DataSource(oDataSource As IDataSource)
    Set m_data_source = oDataSource
End Property

Public Property Get DataSource() As IDataSource
    Set DataSource = m_data_source
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : RecordTable
' Purpose       : The Record s Table Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/21    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let RecordTable(strRecordTable As String)
    m_record_table = strRecordTable
End Property

Public Property Get RecordTable() As String
    RecordTable = m_record_table
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : RecordId
' Purpose       : The Record s Id
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/21    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let RecordId(lngRecordId As Long)
    m_record_id = lngRecordId
End Property

Public Property Get RecordId() As Long
    RecordId = m_record_id
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : PreviousFieldIndexPicked
' Purpose       : The Field s Id picked (using of the dedicated functions)
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/21    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let PreviousFieldIndexPicked(lngPreviousFieldIndexPicked As Long)
    m_previous_field_index_picked = lngPreviousFieldIndexPicked
End Property

Public Property Get PreviousFieldIndexPicked() As Long
    PreviousFieldIndexPicked = m_previous_field_index_picked
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : FieldsList
' Purpose       : The Record s Fileds list
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/21    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set FieldsList(oFieldsList As Collection)
    Set m_fields_list = oFieldsList
End Property

Public Property Get FieldsList() As Collection
    Set FieldsList = m_fields_list
End Property


'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : oRecordDA           The Data Access object
'                   oUser               The Performing User
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal oUser As CUser, ByVal oDataSource As IDataSource, ByVal strRecordTable As String, Optional ByVal lngRecordId As Long = -1)
    
    Set Me.User = oUser
    Set Me.DataSource = oDataSource
    Me.RecordTable = strRecordTable
    
    If Not lngRecordId = -1 Then
        Me.RecordId = lngRecordId
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : Save
'   Purpose     : Save the Record
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Save() As CResult
    Dim result As CResult
    
    If IsNew Then
        Me.DataSource.AddValues Me
    Else
        Me.DataSource.UpdateValues Me
    End If
    
    Set Save = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : List
'   Purpose     : List or Fill values
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/22  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function List() As CResult
    Dim result As CResult, vrnValues As Variant
    
    Set result = Me.DataSource.GetValues(Me, vrnValues)
    
    Set List = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : Delete
'   Purpose     : Delete the Object
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/22  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Delete() As CResult
    Dim result As CResult
    
    Set Delete = result
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
        lngFieldIndex = Me.FieldsList.Count + 1
    Else
        lngFieldIndex = lngAtIndex
    End If
    
    Me.FieldsList.Add oField, oField.Guid
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
    Me.PreviousFieldIndexPicked = MIN_FIELD_INDEX
    
    Set ResetFieldsPicked = result
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
        Set GetNextField = Me.FieldsList(fieldKey)
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
    
    If lngAtIndex = -1 Then
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
    LastFieldReached = (Me.PreviousFieldIndexPicked >= Me.FieldsList.Count)
End Function


'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************

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
Private Function IsNew() As Boolean
    Dim result As Boolean
    
    Set result = True
    
    If Me.RecordId = DEFAULT_ID Then
       result = True
    Else
        result = False
    End If
    
    Set IsNew = result
End Function

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
        Dim nextIndex As Long
        
        ' If the max Index is reached
        If Me.FieldsList.Count = Me.PreviousFieldIndexPicked Then
            GetNextFieldKey = ""
            Exit Function
        End If
        
        nextIndex = Me.PreviousFieldIndexPicked + 1
        Me.PreviousFieldIndexPicked = nextIndex
        
        GetNextFieldKey = m_fields_positions(nextIndex)
    Else
        GetNextFieldKey = m_fields_positions(lngAtIndex)
    End If
End Function



'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    m_guid = CreateGUID
    m_record_id = 0
    Set m_fields_list = New Collection
    m_previous_field_index_picked = MIN_FIELD_INDEX
End Sub

Private Sub Class_Terminate()
    Set m_fields_list = Nothing
End Sub