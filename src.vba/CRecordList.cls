'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CRecordList
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/03
' Purpose   : A list of records
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_user As CModelUser
Private m_data_access As CDataAccess

Private m_list As Collection
Private m_previous_record_index_picked As Long

Private m_uicontrol As CUIControl

Private Const MODULE_NAME As String = "Record-List Class"
Private Const FIRST_INDEX As Long = 1
Private Const RECORD_INDEX_INCREMENT As Long = 1



'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CRecordList
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : User
'   Purpose       : The performing user
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03      Jude Parfait        Created
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
'   2025/02/03      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set DataAccess(oDataAccess As CDataAccess)
    Set m_data_access = oDataAccess
End Property

Public Property Get DataAccess() As CDataAccess
    Set DataAccess = m_data_access
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : List
' Purpose       : The List of Records
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/03    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set List(oList As Collection)
    Set m_list = oList
End Property

Public Property Get List() As Collection
    Set List = m_list
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : FirstRecord
' Purpose       : Get the first record
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/03    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get FirstRecord() As CRecord
    If Me.List.count = 0 Then
        Set FirstRecord = Nothing
    Else
        Set FirstRecord = Me.List.Item(CStr(FIRST_INDEX))
    End If
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : LastRecord
' Purpose       : Get the last record
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/03    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LastRecord() As CRecord
    If Me.List.count = 0 Then
        Set LastRecord = Nothing
    Else
        Set LastRecord = Me.List.Item(CStr(Me.List.count))
    End If
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : PreviousRecordIndexPicked
' Purpose       : The Record s Id picked (using of the dedicated functions)
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/03    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let PreviousRecordIndexPicked(lngPreviousRecordIndexPicked As Long)
    m_previous_record_index_picked = lngPreviousRecordIndexPicked
End Property

Public Property Get PreviousRecordIndexPicked() As Long
    PreviousRecordIndexPicked = m_previous_record_index_picked
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
'   2025/02/03      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(Optional oDataAccess As CDataAccess = Nothing, Optional oUser As CModelUser = Nothing)
    
    If Not oDataAccess Is Nothing Then
        Set Me.DataAccess = oDataAccess
    End If
    
    If Not oUser Is Nothing Then
        Set Me.User = oUser
    End If
    
End Sub

Public Sub FillControl()
    If Not Me.UIControl Is Nothing Then
        Me.UIControl.FillFromRecordList Me
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : AddRecord
'   Purpose       : Add a record to the record list
'   Arguments     : oRecord             The record to be added
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddRecord(ByVal oRecord As CRecord)
    Dim lngCode As Long
    
    lngCode = Me.List.count + 1
    
    Me.List.Add oRecord, CStr(lngCode)
    
    'Set AddSubResult = oResult
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : ResetRecordsPicked
'   Purpose     : Reset the index of Records Picked
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ResetRecordsPicked() As CResult
    Dim result As CResult
    
    Set result = NewResult(MODULE_NAME, "ResetRecordsPicked", True)
    Me.ResetRecordsPicked = FIRST_INDEX - 1
    
    Set ResetRecordsPicked = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetNextRecord
'   Purpose     : Get Next Record, or Record at Index (if given)
'   Arguments   :
'
'   Returns     : CRecord        The Record | Nothing (if not found or anything wrong)
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNextRecord() As CRecord
    
    If Me.LastRecordReached Then
        Set GetNextRecord = Nothing
    Else
        
        Dim lngAtIndex As Long, oRecord As CRecord
        
        lngAtIndex = Me.PreviousRecordIndexPicked + RECORD_INDEX_INCREMENT
        Set oRecord = Me.GetRecordAt(lngAtIndex)
        
        If Not oRecord Is Nothing Then
            Me.PreviousRecordIndexPicked = lngAtIndex
        End If
        
        Set GetNextRecord = oRecord
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetRecordAt
'   Purpose     : Get Record at given index
'   Arguments   : lngAtIndex    Index of Record
'
'   Returns     : CRecord        The Record | Nothing (if not found or anything wrong)
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetRecordAt(ByVal lngAtIndex As Long) As CRecord
    If lngAtIndex < FIRST_INDEX Or lngAtIndex > (Me.List.count + 1) Then
        Set GetRecordAt = Nothing
    Else
        Set GetRecordAt = Me.List.Item(lngAtIndex)
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : LastRecordReached
'   Purpose     : Determine whether the last record has been picked
'   Arguments   :
'
'   Returns     : Boolean
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastRecordReached() As Boolean
    Dim lngRecordCount As Long
    
    lngRecordCount = Me.List.count
    LastRecordReached = ((Me.PreviousFieldIndexPicked + RECORD_INDEX_INCREMENT) >= lngRecordCount)
End Function



'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************


Private Sub Class_Initialize()
    Set m_list = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_list = Nothing
End Sub