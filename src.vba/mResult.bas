'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : mResult
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/03
' Purpose   : Manage Result related operations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine  : AddSubResult
'   Purpose     : Add a sub result to a given one
'   Arguments   : strLabel          The Main result Label
'                 oResult           The Main result
'                 oSubResult        The sub-result
'
'   Returns     : void
'
'   Date          Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/03    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSubResult(ByVal strLabel As String, ByRef oResult As CResult, ByVal oSubResult As CResult)
    If oResult Is Nothing Then
        If oSubResult Is Nothing Then
            Set oResult = NewResult(strLabel)
        ElseIf strLabel = oSubResult.Label Then
            Set oResult = oSubResult
        Else
            Set oResult = NewResult(strLabel)
        End If
    End If
    
    If Not strLabel = oResult.Label Then
        oResult.AddSubResult oSubResult
    End If
End Sub