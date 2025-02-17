'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : mFactory
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/08
' Purpose   : Manage all factories for Main Classes instantiation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewModel
'   Purpose     : Create and Initialize a New Model
'   Arguments   : blnIsOK           Tells if the operation has been well performed
'                 lngCode           The Result Code
'                 strMessage        The Result Message
'
'   Returns     : CResult
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/15      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewModel(ByVal oDataSource As IDataSource, ByVal strTableForSaving As String, Optional ByVal strTableForSelecting As String = "", Optional ByVal oUser As CModelUser = Nothing) As CModel
    With New CModel
        .Init oDataSource, strTableForSaving, strTableForSelecting, oUser
        Set NewModel = .Self 'returns the newly created instance
    End With
End Function

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
Public Function NewUser(Optional ByVal oModel As CModel = Nothing, Optional ByVal lngId As Long = -1, Optional ByVal strLogin As String = "", Optional ByVal strName As String = "", Optional ByVal strPwd As String = "") As CModelUser
    With New CModelUser
        .Init lngId, oModel, strLogin, strName, strPwd
        Set NewUser = .Self 'returns the newly created instance
    End With
End Function

Public Function NewUserFromBD(Optional ByVal lngId As Long = -1, Optional ByVal strLogin As String = "", Optional ByVal strName As String = "", Optional ByVal strPwd As String = "") As CModelUser
    Dim oModel As CModel
    
    Set oModel = NewModel(GetMainDataSource, USERS_TABLE_FOR_SAVING_DEFAULT, USERS_TABLE_FOR_SELECTING_DEFAULT)
    
    Set NewUserFromBD = NewUser(oModel, lngId, strLogin, strName, strPwd)
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
