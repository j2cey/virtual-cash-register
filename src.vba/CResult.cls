'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CResult
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/12
' Purpose   : An Operation result
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_is_ok As Boolean
Private m_code As Long
Private m_message As String
Private m_label As String

Private m_sub_results As Collection



'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CResult
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IsOK
'   Purpose       : Tells if the operation has been well performed
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IsOK(blnIsOK As Boolean)
    m_is_ok = blnIsOK
End Property

Public Property Get IsOK() As Boolean
    IsOK = m_is_ok
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : Code
'   Purpose       : The Result Code
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Code(lngCode As Long)
    m_code = lngCode
End Property

Public Property Get Code() As Long
    Code = m_code
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : Message
' Purpose       : The Result Message
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/08    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Message(strMessage As String)
    m_message = strMessage
End Property

Public Property Get Message() As String
    Message = m_message
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : Label
' Purpose       : The Result Label
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/18    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Label(strLabel As String)
    m_label = strLabel
End Property

Public Property Get Label() As String
    Label = m_label
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : SubResults
' Purpose       : The Sub-Results List
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/03    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set SubResults(oSubResults As Collection)
    Set m_sub_results = oSubResults
End Property

Public Property Get SubResults() As Collection
    Set SubResults = m_sub_results
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : LastResult
' Purpose       : Get the last result if any or self
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/03    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LastResult() As CResult
    If Me.SubResults.Count = 0 Then
        Set LastResult = Self
    Else
        Set LastResult = Me.SubResults.Item(CStr(Me.SubResults.Count))
    End If
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
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(Optional strLabel As String = "", Optional blnIsOK As Boolean = False, Optional lngCode As Long = 0, Optional strMessage As String = "")
    
    Me.Label = strLabel
    Me.IsOK = blnIsOK
    Me.Code = lngCode
    Me.Message = strMessage
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : DisplayMessage
'   Purpose       : Display the Result Message
'   Arguments     :
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayMessage()
    If Me.IsOK Then
        MsgBox "Erreur de Traitement !" & vbCrLf & Me.Message, vbInformation, SettingApp.Name
    Else
        MsgBox "Succes de Traitement !" & vbCrLf & Me.Message, vbCritical, SettingApp.Name
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : AddSubResult
'   Purpose       : Add a sub-result to this result
'   Arguments     : oResult             The result to be added
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSubResult(ByVal oResult As CResult)
    Dim lngCode As Long
    
    lngCode = Me.SubResults.Count + 1
    oResult.Code = lngCode
    
    Me.SubResults.Add oResult, CStr(lngCode)
    
    'Set AddSubResult = oResult
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : SetSuccess
'   Purpose       : Set this result as success
'   Arguments     : strMessage          The result message, if any
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetSuccess(Optional ByVal strMessage As String = "")
    Me.IsOK = True
    If strMessage = "" Then
        Me.Message = "SUCCESS !"
    Else
        Me.Message = strMessage
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : SetFailure
'   Purpose       : Set this result as failure
'   Arguments     : strMessage          The result message, if any
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetFailure(Optional ByVal strMessage As String = "")
    Me.IsOK = False
    If strMessage = "" Then
        Me.Message = "FAILURE !"
    Else
        Me.Message = strMessage
    End If
End Sub


'***************************************************************************************************************************************************************
'    Private Events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    m_code = 0
    Set m_sub_results = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_sub_results = Nothing
End Sub

