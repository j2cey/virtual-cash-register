'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CSetting
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/14
' Purpose   : Setting Super Class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_business_logic As CBusinessLogic
Private m_last_result As CResult

Private m_start_row As Long

Private Const MODULE_NAME As String = "Setting"

'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CSetting
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : BusinessLogic
'   Purpose       : The Business Logic object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/14      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set BusinessLogic(oBusinessLogic As CBusinessLogic)
    Set m_business_logic = oBusinessLogic
End Property

Public Property Get BusinessLogic() As CBusinessLogic
    Set BusinessLogic = m_business_logic
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : LastResult
'   Purpose       : The Last Result
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/14      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set LastResult(oLastResult As CResult)
    Set m_last_result = oLastResult
End Property

Public Property Get LastResult() As CResult
    Set LastResult = m_last_result
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : StartRow
'   Purpose       : The the start row, if any
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/14      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let StartRow(lngStartRow As Long)
    m_start_row = lngStartRow
End Property

Public Property Get StartRow() As Long
    StartRow = m_start_row
End Property

'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : blnIsOK             Tells if the operation has been well performed
'                   lngCode             The Result Code
'                   strMessage          The Result Message
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/14      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal oUser As CModelUser, ByVal oDataSource As IDataSource, ByVal strTableForSaving As String, Optional ByVal strTableForSelecting As String = "", Optional ByVal lngStartRow As Long = -1)
    
    Call InitBusinessLogic(oUser, oDataSource, strTableForSaving, strTableForSelecting)
    Me.StartRow = lngStartRow
    
End Sub

Public Sub LoadValues()
    Dim result As CResult, oRecordValue As CRecord
    
    Set result = Me.BusinessLogic.ReadValues(False)
    
    If Not result.IsOK Then
        result.Notify
    End If
End Sub

Public Sub SaveValues()
    Dim result As CResult, oRecordValue As CRecord
    
    Set result = Me.BusinessLogic.Save(False)
    
    If Not result.IsOK Then
        result.Notify
    End If
End Sub

Public Function AddField(ByVal oFieldType As IFieldValue, ByVal strFieldName As String, ByVal strFieldLabel As String, Optional ByVal lngOffsetEnd As Long = -1, Optional ByVal lngOffsetStart As Long = -1, Optional ByVal eStoreFieldName As enuStoreFieldName) As CField
    Dim oNewField As CField
    
    Set oNewField = Me.BusinessLogic.AddField(oFieldType, strFieldName, strFieldLabel)
    If Not oNewField Is Nothing Then
        SetFieldOffsets oNewField, lngOffsetEnd, lngOffsetStart
        oNewField.StoreFieldName = eStoreFieldName
    End If
    
    Set AddField = oNewField
End Function

Public Sub SetFieldValue(ByVal strFieldKey As String, ByVal varFieldValue As Variant)
    Set Me.LastResult = SetNewResult("Set Field Value", MODULE_NAME)
    
    If Not Me.BusinessLogic Is Nothing Then
        Set Me.LastResult = Me.BusinessLogic.SetFieldValue(strFieldKey, varFieldValue)
    Else
        Me.LastResult.SetFailure "The Business Logic Object not set yet !", True
    End If
End Sub

Public Function GetField(ByVal strFieldKey As String) As CField
    Set Me.LastResult = SetNewResult("Get Field", MODULE_NAME)
    
    If Not Me.BusinessLogic Is Nothing Then
        Set GetField = Me.BusinessLogic.GetField(strFieldKey)
    Else
        Me.LastResult.SetFailure "The Business Logic Object not set yet !", True
        Set GetField = Nothing
    End If
End Function

Public Function GetFieldValue(ByVal strFieldKey As String) As Variant
    Set Me.LastResult = SetNewResult("Get Field Value", MODULE_NAME)
    
    If Not Me.BusinessLogic Is Nothing Then
        GetFieldValue = Me.BusinessLogic.GetFieldValue(strFieldKey).GetValue()
    Else
        Me.LastResult.SetFailure "The Business Logic Object not set yet !", True
        GetFieldValue = Empty
    End If
End Function

Public Function IsFieldEmpty(ByVal strFieldKey As String) As Boolean
    Dim oField As CField
    
    Set oField = Me.GetField(strFieldKey)
    
    If oField Is Nothing Then
        IsFieldEmpty = True
    Else
        IsFieldEmpty = oField.FieldValue.IsEmpty
    End If
    
End Function



'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************

Private Sub InitBusinessLogic(ByVal oUser As CModelUser, ByVal oDataSource As IDataSource, ByVal strTableForSaving As String, Optional ByVal strTableForSelecting As String = "")
    Set Me.BusinessLogic = NewBusinessLogic(NewDataAccess(oUser, oDataSource, strTableForSaving, strTableForSaving), oUser)
End Sub

Private Sub SetFieldOffsets(ByVal oField As CField, Optional ByVal lngOffsetEnd As Long = -1, Optional ByVal lngOffsetStart As Long = -1)
    oField.OffsetEnd = lngOffsetEnd
    oField.OffsetStart = lngOffsetStart
    
    If Not (lngOffsetEnd = -1) Then
        If lngOffsetStart = -1 Then
            Dim lngRecFieldsCount As Long
            
            lngRecFieldsCount = Me.BusinessLogic.GetRecord().FieldListCount
            
            oField.OffsetStart = Me.StartRow + lngRecFieldsCount
        End If
    End If
End Sub


'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    
End Sub






