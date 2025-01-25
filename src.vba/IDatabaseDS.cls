'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : IDatabaseDS
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/19
' Purpose   : Database Data Source Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit



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

Public Function BuildConnectionString(currentDb As CDatabaseDS) As String
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : BuildSelectStatement
'   Purpose     : Construct a full SELECT statement, set local properties, and
'                 return string to calling routine.
'
'   Arguments   : CDatabaseDS   currentDb   The relevant database to build connection string
'
'   Returns     : CResult       The Treatment Result
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
'   2025/01/24  Jude Parfait        Add ByRef Arguments: strSelectStatement, lngCriteriaCount; Change Return Type: CResult
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildSelectStatement(currentDb As CRecordableDA, ByRef strSelectStatement As String, ByRef lngCriteriaCount As Long) As CResult
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
Public Function BuildInsertStatement(currentDb As CDatabaseDS, ByRef strInsertStatement As String) As CResult
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
Public Function BuildUpdateStatement(currentDb As CDatabaseDS, ByRef strUpdateStatement As String) As CResult
End Function