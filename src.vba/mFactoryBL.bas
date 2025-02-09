'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : mFactoryBL
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/31
' Purpose   : Manage all factories for Business Logic related Classes instantiation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewRecord
'   Purpose     : Create and Initialize a New Record
'   Arguments   : oUser             The User
'                 oDataAccess       The Data Access object
'
'   Returns     : CBusinessLogic
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/31      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewBusinessLogic(ByVal oUser As CUser, ByVal oDataAccess As CDataAccess) As CBusinessLogic
    With New CBusinessLogic
        .Init oUser, oDataAccess
        Set NewBusinessLogic = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewRecord
'   Purpose     : Create and Initialize a New Record
'   Arguments   : oDataAccess       The Data Access object
'                 oUser             The User
'                 lngRecordId       The Record ID, in case already saved
'
'   Returns     : CFieldList
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/31      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewRecord(Optional oDataAccess As CDataAccess = Nothing, Optional oUser As CUser = Nothing, Optional ByVal lngRecordId As Long = -1) As CRecord
    With New CRecord
        .Init oDataAccess, oUser, lngRecordId
        Set NewRecord = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewRecordList
'   Purpose     : Create and Initialize a New Record List
'   Arguments   : oDataAccess           The Data Access object
'                 oUser                 The User
'
'   Returns     : CFieldList
'
'   Date        Developer               Action
'   ---------------------------------------------------------------------------------------
'   2025/01/31      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewRecordList(Optional oDataAccess As CDataAccess = Nothing, Optional oUser As CUser = Nothing) As CRecordList
    With New CRecordList
        .Init oDataAccess, oUser
        Set NewRecordList = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewField
'   Purpose     : Create and Initialize a New Field
'   Arguments   :
'
'   Returns     : CField
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/31      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewField(ByVal oFieldValue As IFieldValue, strName As String, strLabel As String, Optional vrnValue As Variant = Null) As CField
    With New CField
        .Init oFieldValue, strName, strLabel, vrnValue
        Set NewField = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewFieldList
'   Purpose     : Create and Initialize a New FieldList
'   Arguments   :
'
'   Returns     : CFieldList
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/30      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewFieldList() As CFieldList
    With New CFieldList
        .Init
        Set NewFieldList = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewFieldValueInteger
'   Purpose     : Create and Initialize a New Integer Field Value
'   Arguments   :
'
'   Returns     : CFieldValueInteger
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/31      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewFieldValueInteger(Optional ByVal oUpperField As CField = Nothing) As CFieldValueInteger
    With New CFieldValueInteger
        .Init oUpperField
        Set NewFieldValueInteger = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewFieldValueString
'   Purpose     : Create and Initialize a New String Field Value
'   Arguments   :
'
'   Returns     : CFieldValueString
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/31      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewFieldValueString(Optional ByVal oUpperField As CField = Nothing) As CFieldValueString
    With New CFieldValueString
        .Init oUpperField
        Set NewFieldValueString = .Self 'returns the newly created instance
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : NewFieldValueBoolean
'   Purpose     : Create and Initialize a New Boolean Field Value
'   Arguments   :
'
'   Returns     : CFieldValueBoolean
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/31      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewFieldValueBoolean(Optional ByVal oUpperField As CField = Nothing) As CFieldValueBoolean
    With New CFieldValueBoolean
        .Init oUpperField
        Set NewFieldValueBoolean = .Self 'returns the newly created instance
    End With
End Function