'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CAccess2007
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/19
' Purpose   : Access 2007 Class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_user As CUser
Private m_upper_database As CDatabaseDS
Private m_current_recordable As CRecordableDA

Private m_request_string As String
Private m_criteria_count As Long


Implements IDatabaseDS


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CAccess2007
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : User
'   Purpose       : The performing user
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set User(oUser As CUser)
    Set m_user = oUser
End Property

Public Property Get User() As CUser
    Set User = m_user
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : UpperDatabase
'   Purpose       : The Upper Database
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/19      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set UpperDatabase(oUpperDatabase As CDatabaseDS)
    Set m_upper_database = oUpperDatabase
End Property

Public Property Get UpperDatabase() As CDatabaseDS
    Set UpperDatabase = m_upper_database
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
    Set m_current_recordable = oCurrentRecordable
End Property

Public Property Get CurrentRecordable() As CRecordableDA
    Set CurrentRecordable = m_current_recordable
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : RequestString
'   Purpose       : The Request String
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let RequestString(strRequestString As String)
    m_request_string = strRequestString
End Property

Public Property Get RequestString() As String
    RequestString = m_request_string
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : CriteriaCount
'   Purpose       : The number of criteria
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let CriteriaCount(lngCriteriaCount As Long)
    m_criteria_count = lngCriteriaCount
End Property

Public Property Get CriteriaCount() As Long
    CriteriaCount = m_criteria_count
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
Public Sub Init(Optional oUser As CUser = Nothing, Optional oRecordableDA As CRecordableDA = Nothing)
    
    If Not oRecordableDA Is Nothing Then
        Me.CurrentRecordable = oRecordableDA
    End If
    
    If Not oUser Is Nothing Then
        Me.User = oUser
    End If
End Sub



'***************************************************************************************************************************************************************
'    Private Functions and Subs
'***************************************************************************************************************************************************************


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : BuildConnectionString
'   Purpose     : Construct a full connection string, set local properties, and return string to calling routine.
'   Arguments   :
'
'   Returns     : CResult       The Treatment Result
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetConnectionString(ByRef strConnectionString As String) As CResult
    Dim targetDbFullPath As String, rslt As CResult
    
    strConnectionString = ""
    Set rslt = NewResult("GetConnectionString", True)
    
    If Not Me.UpperDatabase Is Nothing Then
        'Build connection string
        
        ' Exit Sub if the CurrentServer is not set in the Upper Database
        If Me.UpperDatabase.ServerOrPath = "" Then
            rslt.IsOK = False
            rslt.Message = "The Server / Path is not set in the Upper Database"
        End If
        
        ' Exit Sub if the CurrentDatabase is not set in the Upper Database
        If rslt.IsOK And Me.UpperDatabase.DatabaseOrFileName = "" Then
            rslt.IsOK = False
            rslt.Message = "The Database / File Name is not set in the Upper Database"
        End If
        
        If rslt.IsOK Then
            'targetDbFullPath = AppPath & Application.PathSeparator & gblDbLocalFolder & Application.PathSeparator & gblDbName
            targetDbFullPath = Me.UpperDatabase.ServerOrPath & Application.PathSeparator & Me.UpperDatabase.DatabaseOrFileName
            strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & _
            targetDbFullPath & ".accdb;"
            'strConnectionString = Me.UpperDatabase.Connection.ConnectionString
        End If
    Else
        rslt.IsOK = False
        rslt.Message = "The Upper Database Is Missing"
    End If
    
    Set GetConnectionString = rslt
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetSelectStatement
'   Purpose     : Construct a full SELECT statement, set local properties, and
'                 return string to calling routine.
'
'   Arguments   : CRecordableDA     dataAccess   The relevant data access to build connection string
'
'   Returns     : String            The Build SELECT statement | empty string
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSelectStatement(dataAccess As CRecordableDA, ByRef strSelectStatement As String, ByRef lngCriteriaCount As Long) As CResult
    Dim finalRqstStr As String, criterionDone As Boolean, rslt As CResult
    
    Me.RequestString = ""
    Me.CriteriaCount = 0
    Set Me.CurrentRecordable = dataAccess
    Set rslt = NewResult("GetConnectionString", True)
    
    ' If the Table Name is not given, we return empty string
    If dataAccess.RecordTable = "" Then
        rslt.IsOK = False
        rslt.Message = "The Table Name is not given"
    Else
        Me.RequestString = "SELECT * FROM " & dataAccess.RecordTable
    End If
    
    If rslt.IsOK Then
        AddCriteriaToRequest
    End If
    
    'If Me.CriterionProvidedCount = 0 And lLimitLines > 0 Then
    '    Me.RequestString = SqlSelectSetLIMIT(Me.RequestStr, lLimitLines)
    'End If
    
    strSelectStatement = Me.RequestString
    lngCriteriaCount = Me.CriteriaCount
    Set GetSelectStatement = rslt
End Function

Private Function AddCriteriaToRequest() As CResult
    Dim rslt As CResult, loopDone As Boolean, field As CField
    
    Set rslt = NewResult("AddCriteriaToRequest", True)
    loopDone = False
    
    If Me.CurrentRecordable Is Nothing Then
        rslt.IsOK = False
        rslt.Message = "The Current Recordable is not set"
    End If
    
    If rslt.IsOK Then
        Me.CurrentRecordable.Record.FieldList.ResetFieldsPicked
        Do
            Set field = Me.CurrentRecordable.Record.FieldList.GetNextFieldCriteria()
            
            If Not field Is Nothing Then
                ' Set criterion values
                If Not field.IsEmpty Then
                    StrRequestAppendCriterion field
                End If
            End If
            
            If Me.CurrentRecordable.Record.FieldList.LastFieldReached() Then
                ' The end of list is reached
                Exit Do
                loopDone = True
            End If
            
        Loop While (Not loopDone)
    End If
    
    Set AddCriteriaToRequest = rslt
End Function

Private Sub StrRequestAppendCriterion(oField As CField)
    Dim strOperator As String
    
    strOperator = IIf(oField.SearchCriteriaMustMatchExact, "=", "LIKE")
    
    If Me.CriteriaCount = 0 Then
        Me.RequestString = Me.RequestString & " WHERE "
    Else
        Me.RequestString = Me.RequestString & " AND "
    End If
    
    Me.RequestString = Me.RequestString & " " & oField.Name & strOperator & " @" & oField.Name
    
    Me.CriteriaCount = Me.CriteriaCount + 1
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetInsertStatement
'   Purpose     : Construct a full INSERT statement, set local properties, and
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
Private Function GetInsertStatement(dataAccess As CRecordableDA, ByRef strInsertStatement As String) As CResult
    Dim finalRqstStr As String, rslt As CResult
    
    rslt = NewResult("GetInsertStatement", True)
    
    Set GetInsertStatement = rslt
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetUpdateStatement
'   Purpose     : Construct a full UPDATE statement, set local properties, and
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
Private Function GetUpdateStatement(dataAccess As CDatabaseDS, ByRef strUpdateStatement As String) As CResult
    Dim finalRqstStr As String, rslt As CResult
    
    rslt = NewResult("GetUpdateStatement", True)
    
    Set GetUpdateStatement = rslt
End Function



'***************************************************************************************************************************************************************
'   Interfaces Implementation
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IAccessDB_UpperDatabase
'   Purpose       : Get / Set the Upper Database
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set IDatabaseDS_UpperDatabase(oUpperDatabase As CDatabaseDS)
    Set Me.UpperDatabase = oUpperDatabase
End Property

Public Property Get IDatabaseDS_UpperDatabase() As CDatabaseDS
    Set IDatabaseDS_UpperDatabase = Me.UpperDatabase
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IAccessDB_CurrentRecordable
'   Purpose       : The current Recordable Object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set IDatabaseDS_CurrentRecordable(oCurrentRecordable As CRecordableDA)
    Set Me.CurrentRecordable = oCurrentRecordable
End Property

Public Property Get IDatabaseDS_CurrentRecordable() As CRecordableDA
    Set IDatabaseDS_CurrentRecordable = Me.CurrentRecordable
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : RequestString
'   Purpose       : The Request String
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IDatabaseDS_RequestString(strRequestString As String)
    Me.RequestString = strRequestString
End Property

Public Property Get IDatabaseDS_RequestString() As String
    IDatabaseDS_RequestString = Me.RequestString
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IDatabaseDS_BuildConnectionString
'   Purpose     : Construct a full connection string, set local properties, and
'                 return string to calling routine.
'
'   Arguments   :
'
'   Returns     : CResult       The Treatment Result
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/19  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IDatabaseDS_BuildConnectionString(ByRef strConnectionString As String) As CResult
    Set IDatabaseDS_BuildConnectionString = GetConnectionString(strConnectionString)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IDatabaseDS_BuildSelectStatement
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
Public Function IDatabaseDS_BuildSelectStatement(dataAccess As CRecordableDA, ByRef strSelectStatement As String, ByRef lngCriteriaCount As Long) As CResult
    Set IDatabaseDS_BuildSelectStatement = GetSelectStatement(dataAccess, strSelectStatement, lngCriteriaCount)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IDatabaseDS_BuildInsertStatement
'   Purpose     : Construct a full INSERT statement, set local properties, and
'                 return string to calling routine.
'
'   Arguments   : dataAccess           currentDb   The relevant database to build connection string
'                 strInsertStatement    The Returned Insert Statement
'
'   Returns     : String
'
'   Date        Developer               Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait            Created
'   2025/01/24  Jude Parfait            Add ByRef Arguments: strInsertStatement; Change Return Type: CResult
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IDatabaseDS_BuildInsertStatement(dataAccess As CRecordableDA, ByRef strInsertStatement As String) As CResult
    Set IDatabaseDS_BuildInsertStatement = GetInsertStatement(dataAccess, strInsertStatement)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IDatabaseDS_BuildUpdateStatement
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
Public Function IDatabaseDS_BuildUpdateStatement(dataAccess As CDatabaseDS, ByRef strUpdateStatement As String) As CResult
    Set IDatabaseDS_BuildUpdateStatement = GetUpdateStatement(dataAccess, strUpdateStatement)
End Function

