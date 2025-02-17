'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CBusinessLogic
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/08
' Purpose   : The Business Logic Object
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


Private m_user As CModelUser
Private m_data_access As CDataAccess

Private m_output_record_list As CRecordList
Private m_output_record As CRecord
Private m_output_field As CField

Private m_last_result As CResult


Private Const MODULE_NAME As String = "Business-Logic"


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CBusinessLogic
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : User
'   Purpose       : The performing user
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/08      Jude Parfait        Created
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
'   2025/02/08      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set DataAccess(oDataAccess As CDataAccess)
    Set m_data_access = oDataAccess
End Property

Public Property Get DataAccess() As CDataAccess
    Set DataAccess = m_data_access
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : OutputRecordList
'   Purpose       : The Output Record List
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/08      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set OutputRecordList(oOutputRecordList As CRecordList)
    Set m_output_record_list = oOutputRecordList
End Property

Public Property Get OutputRecordList() As CRecordList
    Set OutputRecordList = m_output_record_list
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : OutputRecord
'   Purpose       : The Output Record
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/08      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set OutputRecord(oOutputRecord As CRecord)
    Set m_output_record = oOutputRecord
End Property

Public Property Get OutputRecord() As CRecord
    Set OutputRecord = m_output_record
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : OutputField
'   Purpose       : The Output Field
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/08      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set OutputField(oOutputField As CField)
    Set m_output_field = oOutputField
End Property

Public Property Get OutputField() As CField
    Set OutputField = m_output_field
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : LastResult
'   Purpose       : The Last Result
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/08      Jude Parfait        Created
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
'   Arguments     : oUser               The Performing User
'                   oRecordDA           The Data Access object
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/08      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal oDataAccess As CDataAccess, Optional ByVal oUser As CModelUser = Nothing)
    
    Set Me.DataAccess = oDataAccess
    
    If Not oUser Is Nothing Then
        Set Me.User = oUser
    End If
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : ReadValues
'   Purpose     : Read Values (One Record) from Data Access
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/02/08  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReadValues(Optional ByVal blnQuietly As Boolean = True) As CResult
    
    Set Me.LastResult = SetNewResult("Read Values", MODULE_NAME)
    
    If CheckDataAccess Then
        Set Me.OutputRecord = Me.DataAccess.GetRecord(Me.LastResult)
        
        If Me.LastResult.IsOK Then
            Set Me.DataAccess.Record.FieldList = Me.OutputRecord.FieldList
        End If
    End If
    
    Set ReadValues = Me.LastResult
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : ReadValuesList
'   Purpose     : Read Values List (Several Records) from Data Access
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/02/08  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReadValuesList(Optional ByVal blnQuietly As Boolean = True) As CRecordList
    
    Set Me.LastResult = SetNewResult("Read Values List", MODULE_NAME)
    
    If CheckDataAccess Then
        Set Me.OutputRecordList = Me.DataAccess.GetRecordList(Me.LastResult)
    End If
    
    Set ReadValuesList = Me.LastResult
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : ReadSingleValue
'   Purpose     : Read a Single Value (One Field) from Data Access
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/02/08  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReadSingleValue(Optional ByVal blnQuietly As Boolean = True) As CField
    
    Set Me.LastResult = SetNewResult("Read Single Value", MODULE_NAME)
    
    If CheckDataAccess Then
        Set Me.OutputField = Me.DataAccess.GetValue(Me.LastResult)
    End If
    
    Set ReadSingleValue = Me.LastResult
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
'   2025/02/08  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Save(Optional ByVal blnQuietly As Boolean = True) As CResult
    
    Set Me.LastResult = SetNewResult("Save", MODULE_NAME)
    
    If CheckDataAccess Then
        If Me.DataAccess.Record.IsNew Then
            Set Me.LastResult = Me.DataAccess.Create()
        Else
            Set Me.LastResult = Me.DataAccess.Update()
        End If
    End If
    
    Set Save = Me.LastResult
End Function

Public Function AddField(ByVal oFieldType As IFieldValue, ByVal strNameForSaving As String, ByVal strFieldLabel As String, Optional ByVal strNameForSelecting As String = "") As CField
    Dim result As CResult, oNewField As CField
    
    If CheckRecord Then
        Set oNewField = NewField(oFieldType, strNameForSaving, strFieldLabel, Null, strNameForSaving, strNameForSelecting)
        
        Set result = Me.DataAccess.Record.FieldList.AddField(oNewField)
        
        Set AddField = oNewField
    Else
        Set AddField = Nothing
    End If
End Function

Public Function GetRecord() As CRecord
    Set GetRecord = Me.DataAccess.Record
End Function

Public Function GetField(ByVal strFieldName As String) As CField
    If CheckRecord Then
        Set GetField = Me.DataAccess.Record.FieldList.GetFieldByName(strFieldName)
    Else
        Set GetField = Nothing
    End If
End Function

Public Function GetOrAddField(ByVal strFieldName As String, ByVal oFieldType As IFieldValue, ByVal strNameForSaving As String, ByVal strFieldLabel As String, ByVal strNameForSelecting As String)
    Dim oField As CField
    
    Set oField = Me.GetField(strFieldName)
    
    If Not oField Is Nothing Then
        Set GetOrAddField = oField
    Else
        Set GetOrAddField = Me.AddField(oFieldType, strNameForSaving, strFieldLabel, strNameForSelecting)
    End If
End Function

Public Function GetFieldValue(ByVal strFieldName As String) As IFieldValue
    Dim oField As CField
    
    Set oField = Me.GetField(strFieldName)
    
    If Not oField Is Nothing Then
        Set GetFieldValue = oField.FieldValue
    Else
        Set GetFieldValue = Nothing
    End If
End Function

Public Function SetFieldValue(ByVal strFieldName As String, ByVal varFieldValue As Variant, Optional ByVal oFieldType As IFieldValue = Nothing, Optional ByVal strNameForSaving As String = "", Optional ByVal strFieldLabel As String = "", Optional ByVal strNameForSelecting As String = "") As CResult
    
    Set Me.LastResult = SetNewResult("Set Field Value", MODULE_NAME)
    
    If CheckRecord Then
        Dim oField As CField
        
        If Not oFieldType Is Nothing Then
            Set oField = GetOrAddField(strFieldName, oFieldType, strNameForSaving, strFieldLabel, strNameForSelecting)
        Else
            Set oField = Me.GetField(strFieldName)
        End If
        
        If Not oField Is Nothing Then
            oField.SetValue (varFieldValue)
        Else
            Me.LastResult.SetFailure "Field '" & strFieldName & "' not found !", True
        End If
    End If
    
    Set SetFieldValue = Me.LastResult
End Function


'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************

Private Function CheckDataAccess() As Boolean
    If Not Me.DataAccess Is Nothing Then
        CheckDataAccess = True
    Else
        Me.LastResult.SetFailure "The Data Access object is not set yet !", True
        CheckDataAccess = False
    End If
End Function

Private Function CheckRecord() As Boolean
    If CheckDataAccess Then
        If Not Me.DataAccess.Record Is Nothing Then
            CheckRecord = True
        Else
            Me.LastResult.SetFailure "The Record object is not set yet !", True
            CheckRecord = False
        End If
    End If
End Function




'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    
End Sub
