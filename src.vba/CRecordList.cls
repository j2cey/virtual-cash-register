'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CRecordList
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/03
' Purpose   : A list of records
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


Private m_list As Collection
Private m_previous_record_index_picked As Long

Private Const FIRST_INDEX As Long = 1



'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CRecordList
    Set Self = Me
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
    If Me.List.Count = 0 Then
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
    If Me.List.Count = 0 Then
        Set LastRecord = Nothing
    Else
        Set LastRecord = Me.List.Item(CStr(Me.List.Count))
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
'   2025/02/03      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init()
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
    
    lngCode = Me.List.Count + 1
    
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
    
    Set result = NewResult("ResetRecordsPicked", True)
    Me.ResetRecordsPicked = FIRST_INDEX - 1
    
    Set ResetRecordsPicked = result
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