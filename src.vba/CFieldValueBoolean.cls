'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CFieldValueBoolean
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/05
' Purpose   : The String Field Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_value As Boolean
Private m_upper_field As CField
Private m_module_name As String
Private m_is_empty As Boolean

Implements IFieldValue


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CFieldValueBoolean
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Value
'   Purpose       : The Inner Value
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/05      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Value(strValue As Boolean)
    m_value = strValue
End Property

Public Property Get Value() As Boolean
    Value = m_value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IsEmpty
'   Purpose       : Determine whether the (inner) value is empty
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsEmpty(blnIsEmpty As Boolean)
    m_is_empty = blnIsEmpty
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = m_is_empty
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : UpperField
'   Purpose       : The Upper Field
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set UpperField(oUpperField As CField)
    Set m_upper_field = oUpperField
End Property

Public Property Get UpperField() As CField
    Set UpperField = m_upper_field
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : ModuleName
'   Purpose       : The Name of this Module
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ModuleName(strModuleName As String)
    Let m_module_name = strModuleName
End Property

Public Property Get ModuleName() As String
    ModuleName = m_module_name
End Property




'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : oUpperField the Upper Field
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(Optional ByVal oUpperField As CField = Nothing)
    If Not oUpperField Is Nothing Then
        Set Me.UpperField = oUpperField
    End If
End Sub


'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine  : SetValue
'   Purpose     : Set the inner Value
'
'   Arguments   : vrnValue          The Inner Value
'
'   Returns     : void
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetValue(ByVal vrnValue As Variant)
    Me.Value = CBool(vrnValue)
    Me.IsEmpty = False
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetValue
'   Purpose     : Get the inner Value
'
'   Arguments   :
'
'   Returns     : void
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetValue() As Variant
    GetValue = CVar(Me.Value)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine  : ClearValue
'   Purpose     : Clear the inner Value
'
'   Arguments   :
'
'   Returns     : void
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearValue()
    Me.Value = False
    Me.IsEmpty = True
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : SetInputParameter
'   Purpose     : Set INPUT parameter to a given command
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetInputParameter(ByRef oRequestCommand As ADODB.Command) As CResult
    Dim result As CResult
    
    Set result = NewResult(Me.ModuleName & " - SetInputParameter", True)
    
    If Me.UpperField.SearchCriteriaMustMatchExact Then
        oRequestCommand.Parameters.Append oRequestCommand.CreateParameter("@" & Me.UpperField.Name, adBoolean, adParamInput, 255, CBool(Me.Value))
    Else
        oRequestCommand.Parameters.Append oRequestCommand.CreateParameter("@" & Me.UpperField.Name, adVarChar, adParamInput, 255, CBool(Me.Value))
    End If
    
    Set SetInputParameter = result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : SetOutputParameter
'   Purpose     : Set OUTPUT parameter to a given command
'   Arguments   :
'
'   Returns     : CResult
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetOutputParameter(ByRef oRequestCommand As ADODB.Command) As CResult
    Dim result As CResult
    
    Set result = NewResult(Me.ModuleName & " - SetOutputParameter", True)
    
    oRequestCommand.Parameters.Append oRequestCommand.CreateParameter("@" & Me.UpperField.Name, adBoolean, adParamOutput, 255, CBool(Me.Value))
    
    Set SetInputParameters = result
End Function



'***************************************************************************************************************************************************************
'   Interfaces Implementation
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IFieldValue_UpperField
'   Purpose       : Get / Set the Upper Field
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/31      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set IFieldValue_UpperField(oUpperField As CField)
    Set Me.UpperField = oUpperField
End Property

Public Property Get IFieldValue_UpperField() As CField
    Set IFieldValue_UpperField = Me.UpperField
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine  : IFieldValue_SetValue
'   Purpose     : Set the inner Value
'
'   Arguments   : vrnValue          The Inner Value
'
'   Returns     : void
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub IFieldValue_SetValue(ByVal vrnValue As Variant)
    SetValue vrnValue
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IFieldValue_GetValue
'   Purpose     : Get the inner Value
'
'   Arguments   :
'
'   Returns     : void
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IFieldValue_GetValue() As Variant
    IFieldValue_GetValue = GetValue
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine  : IFieldValue_ClearValue
'   Purpose     : Clear the inner Value
'
'   Arguments   :
'
'   Returns     : void
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub IFieldValue_ClearValue()
    ClearValue
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IFieldValue_IsEmpty
'   Purpose     : Determine whether the inner Value is empty
'
'   Arguments   :
'
'   Returns     : Boolean
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IFieldValue_IsEmpty() As Boolean
    IFieldValue_IsEmpty = Me.IsEmpty
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property    : IFieldValue_SetInputParameter
'   Purpose     : Set Input Parameters
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IFieldValue_SetInputParameter(ByRef oRequestCommand As ADODB.Command) As CResult
    Set IFieldValue_SetInputParameter = SetInputParameter(oRequestCommand)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property    : IFieldValue_SetOutputParameters
'   Purpose     : Set Output Parameters
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IFieldValue_SetOutputParameter(ByRef oRequestCommand As ADODB.Command) As CResult
    Set IFieldValue_SetOutputParameter = SetOutputParameter(oRequestCommand)
End Function



'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    m_module_name = "CFieldValueBoolean"
End Sub

Private Sub Class_Terminate()
    
End Sub


