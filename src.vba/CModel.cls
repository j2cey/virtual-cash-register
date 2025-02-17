'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CModel
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/15
' Purpose   : Model Super Class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_business_logic As CBusinessLogic
Private m_last_result As CResult

Private Const MODULE_NAME As String = "Model"

'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CModel
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : BusinessLogic
'   Purpose       : The Business Logic object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/15      Jude Parfait        Created
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
'   2025/02/15      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set LastResult(oLastResult As CResult)
    Set m_last_result = oLastResult
End Property

Public Property Get LastResult() As CResult
    Set LastResult = m_last_result
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
Public Sub Init(ByVal oDataSource As IDataSource, ByVal strTableForSaving As String, Optional ByVal strTableForSelecting As String = "", Optional ByVal oUser As CModelUser = Nothing)
    
    Call InitBusinessLogic(oUser, oDataSource, strTableForSaving, strTableForSelecting)
    
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

Public Sub SetId(ByVal lngId As Long)
    Set Me.LastResult = SetNewResult("Set Record ID", MODULE_NAME)
    
    If Not Me.BusinessLogic Is Nothing Then
        Me.LastResult.SetSuccess
        Me.BusinessLogic.GetRecord().RecordId = lngId
    Else
        Me.LastResult.SetFailure "The Business Logic Object not set yet !", True
    End If
End Sub

Public Function GetId() As Long
    Set Me.LastResult = SetNewResult("Get Record ID", MODULE_NAME)
    
    If Not Me.BusinessLogic Is Nothing Then
        Me.LastResult.SetSuccess
        GetId = Me.BusinessLogic.GetRecord().RecordId
    Else
        Me.LastResult.SetFailure "The Business Logic Object not set yet !", True
        GetId = Null
    End If
End Function

Public Function AddField(ByVal oFieldType As IFieldValue, ByVal strNameForSaving As String, ByVal strFieldLabel As String, Optional ByVal strNameForSelecting As String = "") As CField
    Dim oNewField As CField
    
    Set oNewField = Me.BusinessLogic.AddField(oFieldType, strNameForSaving, strFieldLabel, strNameForSelecting)
    
    Set AddField = oNewField
End Function

Public Sub SetFieldValue(ByVal oFieldType As IFieldValue, ByVal strNameForSaving As String, ByVal strFieldLabel As String, ByVal strNameForSelecting As String, ByVal varFieldValue As Variant)
    Set Me.LastResult = SetNewResult("Set Field Value", MODULE_NAME)
    
    If Not Me.BusinessLogic Is Nothing Then
        Set Me.LastResult = Me.BusinessLogic.SetFieldValue(strNameForSaving, varFieldValue, oFieldType, strNameForSaving, strFieldLabel, strNameForSelecting)
    Else
        Me.LastResult.SetFailure "The Business Logic Object not set yet !", True
    End If
End Sub

Public Function GetField(ByVal oFieldType As IFieldValue, ByVal strNameForSaving As String, ByVal strFieldLabel As String, ByVal strNameForSelecting As String) As CField
    Set Me.LastResult = SetNewResult("Get Field", MODULE_NAME)
    
    If Not Me.BusinessLogic Is Nothing Then
        Set GetField = Me.BusinessLogic.GetOrAddField(strNameForSaving, oFieldType, strNameForSaving, strFieldLabel, strNameForSelecting)
    Else
        Me.LastResult.SetFailure "The Business Logic Object not set yet !", True
        Set GetField = Nothing
    End If
End Function

Public Function GetFieldValue(ByVal oFieldType As IFieldValue, ByVal strNameForSaving As String, ByVal strFieldLabel As String, ByVal strNameForSelecting As String) As Variant
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
    
    ' Set RecordId as Field
    Me.BusinessLogic.GetRecord.SetIdField
End Sub


'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    
End Sub







