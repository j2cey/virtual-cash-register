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
        Set oResult = NewResult(strLabel)
    End If
    
    oResult.AddSubResult oSubResult
End Sub