'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : IDatabaseDS
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/19
' Purpose   : Database Data Source Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : UpperDatabase
'   Purpose       : Get / Set the Upper Database
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set UpperDatabase(oUpperDatabase As CDatabaseDS)
End Property

Public Property Get UpperDatabase() As CDatabaseDS
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : CurrentRecordable
'   Purpose       : The current Recordable Object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set CurrentRecordable(oCurrentRecordable As CRecordableDA)
End Property

Public Property Get CurrentRecordable() As CRecordableDA
End Property



'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : BuildConnectionString
'   Purpose     : Construct a full connection string, set local properties, and
'                 return string to calling routine.
'
'   Arguments   : CDatabaseDS   currentDb   The relevant database to build connection string
'
'   Returns     : String
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function BuildConnectionString(ByRef strConnectionString As String) As CResult
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : BuildSelectStatement
'   Purpose     : Construct a full SELECT statement, set local properties, and
'                 return string to calling routine.
'
'   Arguments   : dataAccess   currentDb   The relevant database to build connection string
'
'   Returns     : CResult       The Treatment Result
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
'   2025/01/24  Jude Parfait        Add ByRef Arguments: strSelectStatement, lngCriteriaCount; Change Return Type: CResult
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildSelectStatement(dataAccess As CRecordableDA, ByRef strSelectStatement As String, ByRef lngCriteriaCount As Long) As CResult
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : BuildInsertStatement
'   Purpose     : Construct a full INSERT statement, set local properties, and
'                 return string to calling routine.
'
'   Arguments   : CDatabaseDS           currentDb   The relevant database to build connection string
'                 strInsertStatement    The Returned Insert Statement
'
'   Returns     : String
'
'   Date        Developer               Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait            Created
'   2025/01/24  Jude Parfait            Add ByRef Arguments: strInsertStatement; Change Return Type: CResult
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildInsertStatement(dataAccess As CRecordableDA, ByRef strInsertStatement As String) As CResult
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : BuildUpdateStatement
'   Purpose     : Construct a full UPDATE statement, set local properties, and
'                 return string to calling routine.
'
'   Arguments   : CDatabaseDS           currentDb   The relevant database to build connection string
'                 strUpdateStatement    The returned Update Statement, if any
'
'   Returns     : String
'
'   Date        Developer               Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait            Created
'   2025/01/24  Jude Parfait            Add ByRef Arguments: strUpdateStatement; Change Return Type: CResult
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildUpdateStatement(dataAccess As CDatabaseDS, ByRef strUpdateStatement As String) As CResult
End Function