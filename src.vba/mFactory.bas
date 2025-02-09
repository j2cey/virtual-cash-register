'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : mFactory
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/08
' Purpose   : Manage all factories for Main Classes instantiation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewUser
'   Purpose     : Create and Initialize a New User
'   Arguments   : blnIsOK           Tells if the operation has been well performed
'                 lngCode           The Result Code
'                 strMessage        The Result Message
'
'   Returns     : CResult
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewUser(Optional ByVal lId As Long = -1, Optional ByVal oRecord As CRecord = Nothing, Optional ByVal sLogin As String = "", Optional ByVal sName As String = "", Optional ByVal sPwd As String = "") As CUser
    With New CUser
        .Init lId, oRecord, sLogin, sName, sPwd
        Set NewUser = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewResult
'   Purpose     : Initialize the Object
'   Arguments   : blnIsOK           Tells if the operation has been well performed
'                 lngCode           The Result Code
'                 strMessage        The Result Message
'
'   Returns     : CResult
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewResult(Optional ByVal strModuleName As String = "", Optional ByVal strLabel As String = "", Optional ByVal blnIsOK As Boolean = True, Optional ByVal lngCode As Long = 0, Optional ByVal strMessage As String = "") As CResult
    With New CResult
        .Init strModuleName, strLabel, blnIsOK, lngCode, strMessage
        Set NewResult = .Self 'returns the newly created instance
    End With
End Function
