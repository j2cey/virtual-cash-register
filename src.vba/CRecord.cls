'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CRecord
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/12
' Purpose   : The Record BLL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_record_da As CRecordDA
Private m_user As CUser
Private m_fields_list As Collection
Private m_last_result As CResult
Private m_value As Variant

Implements IRecordableBL


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
Public Property Set User(oUser As CUser)
    Set m_user = oUser
End Property

Public Property Get User() As CUser
    Set User = m_user
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : RecordDA
'   Purpose       : The Data Access object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set RecordDA(oRecordDA As CRecordDA)
    Set m_record_da = oRecordDA
End Property

Public Property Get RecordDA() As CRecordDA
    Set RecordDA = m_record_da
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Value
'   Purpose       : The Value of the Record
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Value(vrnValue As Variant)
    Let m_value = vrnValue
End Property

Public Property Get Value() As Variant
    Value = m_value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : FieldsList
' Purpose       : The Record Fields List
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/08    Jude Parfait        Created
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
'   2025/01/12  Jude Parfait        Created
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
'   Arguments     : oRecordDA           The Data Access object
'                   oUser               The Performing User
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(Optional oRecordDA As CRecordDA = Nothing, Optional oUser As CUser = Nothing)
    
    If Not oRecordDA Is Nothing Then
        Me.RecordDA = oRecordDA
    End If
    
    If Not oUser Is Nothing Then
        Me.User = oUser
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : AddField
'   Purpose       : Add a Filed to the fields list
'   Arguments     : oField              The Field to be added
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddField(oField As CField)
    Me.FieldsList.Add oField
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
    
    Set result = NewResut()
    result.IsOK = True
    
    If Me.FieldsList Is Nothing Then
    Else
        For Each itemColl In Me.FieldsList
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
    Dim result As CResult, itemColl As Variant, oField As clsUCTL
    
    Set result = NewResut()
    result.IsOK = True
    
    If Me.FieldsList Is Nothing Then
    Else
        For Each itemColl In Me.FieldsList
            Dim currResult As CResult
            
            Set oField = itemColl
            Set currResult = oField.Clear()
            
            If Not currResult.IsOK Then
                Set result = currResult
                Exit For
            End If
            
        Next itemColl
    End If
    
    Set Clear = result
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
    
    Set result = NewResut()
    result.IsOK = True
    
    If Me.FieldsList Is Nothing Then
    Else
        For Each itemColl In Me.FieldsList
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
    
    If Me.FieldsList Is Nothing Then
    Else
        For Each itemColl In Me.FieldsList
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property    : IRecordableBL_LastResult
'   Purpose     : Get / Set The Last Result
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get IRecordableBL_LastResult() As CResult
    Set IRecordableBL_LastResult = Me.LastResult
End Property

Public Property Set IRecordableBL_LastResult(RHS As CResult)
    Set Me.LastResult = RHS
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property    : IRecordableBL_FieldsList
'   Purpose     : Get / Set Fields List
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
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
    Set m_fields_list = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_fields_list = Nothing
End Sub
