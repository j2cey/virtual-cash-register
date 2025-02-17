'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CDatabaseAccess2007
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/19
' Purpose   : Access 2007 Class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_user As CModelUser
Private m_upper_database As CDataSourceDatabase
Private m_current_recordable As CDataAccess

Private m_request_string As String
Private m_criteria_count As Long

Private Const MODULE_NAME As String = "Database Access 2007"


Implements IDataSourceDatabase


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CDatabaseAccess2007
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
Public Property Set User(oUser As CModelUser)
    Set m_user = oUser
End Property

Public Property Get User() As CModelUser
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
Public Property Set UpperDatabase(oUpperDatabase As CDataSourceDatabase)
    Set m_upper_database = oUpperDatabase
End Property

Public Property Get UpperDatabase() As CDataSourceDatabase
    Set UpperDatabase = m_upper_database
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : CurrentDataAccess
'   Purpose       : The current DataAccess Object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set CurrentDataAccess(oCurrentDataAccess As CDataAccess)
    Set m_current_recordable = oCurrentDataAccess
End Property

Public Property Get CurrentDataAccess() As CDataAccess
    Set CurrentDataAccess = m_current_recordable
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
Public Sub Init(Optional oUser As CModelUser = Nothing, Optional oDataAccess As CDataAccess = Nothing)
    
    If Not oDataAccess Is Nothing Then
        Me.CurrentDataAccess = oDataAccess
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
    Set rslt = NewResult(MODULE_NAME, "GetConnectionString", True)
    
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
'   Arguments   : CDataAccess     dataAccess   The relevant data access to build connection string
'
'   Returns     : String            The Build SELECT statement | empty string
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSelectStatement(oDataAccess As CDataAccess, ByRef strSelectStatement As String, ByRef lngCriteriaCount As Long) As CResult
    Dim finalRqstStr As String, criterionDone As Boolean, rslt As CResult
    
    Me.RequestString = ""
    Me.CriteriaCount = 0
    Set Me.CurrentDataAccess = oDataAccess
    Set rslt = NewResult(MODULE_NAME, "GetConnectionString", True)
    
    ' If the Table Name is not given, we return empty string
    If oDataAccess.TableForSelecting = "" Then
        rslt.IsOK = False
        rslt.Message = "The Table Name For Selecting is not given"
    Else
        Me.RequestString = "SELECT * FROM " & oDataAccess.TableForSelecting
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
    
    Set rslt = NewResult(MODULE_NAME, "AddCriteriaToRequest", True)
    loopDone = False
    
    If Me.CurrentDataAccess Is Nothing Then
        rslt.IsOK = False
        rslt.Message = "The Current DataAccess is not set"
    End If
    
    If rslt.IsOK Then
        Me.CurrentDataAccess.Record.FieldList.ResetFieldsPicked
        Do
            Set field = Me.CurrentDataAccess.Record.FieldList.GetNextFieldCriteria()
            
            If Not field Is Nothing Then
                ' Set criterion values
                If Not field.IsEmpty Then
                    StrRequestAppendCriterion field
                End If
            End If
            
            If Me.CurrentDataAccess.Record.FieldList.LastFieldReached() Then
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
'   Arguments   : CDataSourceDatabase   currentDb   The relevant database to build connection string
'
'   Returns     : String
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetInsertStatement(oDataAccess As CDataAccess, ByRef strInsertStatement As String) As CResult
    Dim finalRqstStr As String, rslt As CResult
    
    rslt = NewResult(MODULE_NAME, "GetInsertStatement", True)
    
    Set GetInsertStatement = rslt
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetUpdateStatement
'   Purpose     : Construct a full UPDATE statement, set local properties, and
'                 return string to calling routine.
'
'   Arguments   : CDataSourceDatabase   currentDb   The relevant database to build connection string
'
'   Returns     : String
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetUpdateStatement(oDataAccess As CDataSourceDatabase, ByRef strUpdateStatement As String) As CResult
    Dim finalRqstStr As String, rslt As CResult
    
    rslt = NewResult(MODULE_NAME, "GetUpdateStatement", True)
    
    Set GetUpdateStatement = rslt
End Function



'***************************************************************************************************************************************************************
'   Interfaces Implementation  -   IDataSourceDatabase
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IAccessDB_UpperDatabase
'   Purpose       : Get / Set the Upper Database
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set IDataSourceDatabase_UpperDatabase(oUpperDatabase As CDataSourceDatabase)
    Set Me.UpperDatabase = oUpperDatabase
End Property

Public Property Get IDataSourceDatabase_UpperDatabase() As CDataSourceDatabase
    Set IDataSourceDatabase_UpperDatabase = Me.UpperDatabase
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IAccessDB_CurrentDataAccess
'   Purpose       : The current DataAccess Object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set IDataSourceDatabase_CurrentDataAccess(oCurrentDataAccess As CDataAccess)
    Set Me.CurrentDataAccess = oCurrentDataAccess
End Property

Public Property Get IDataSourceDatabase_CurrentDataAccess() As CDataAccess
    Set IDataSourceDatabase_CurrentDataAccess = Me.CurrentDataAccess
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : RequestString
'   Purpose       : The Request String
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IDataSourceDatabase_RequestString(strRequestString As String)
    Me.RequestString = strRequestString
End Property

Public Property Get IDataSourceDatabase_RequestString() As String
    IDataSourceDatabase_RequestString = Me.RequestString
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IDataSourceDatabase_BuildConnectionString
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
Public Function IDataSourceDatabase_BuildConnectionString(ByRef strConnectionString As String) As CResult
    Set IDataSourceDatabase_BuildConnectionString = GetConnectionString(strConnectionString)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IDataSourceDatabase_BuildSelectStatement
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
Public Function IDataSourceDatabase_BuildSelectStatement(oDataAccess As CDataAccess, ByRef strSelectStatement As String, ByRef lngCriteriaCount As Long) As CResult
    Set IDataSourceDatabase_BuildSelectStatement = GetSelectStatement(oDataAccess, strSelectStatement, lngCriteriaCount)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IDataSourceDatabase_BuildInsertStatement
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
Public Function IDataSourceDatabase_BuildInsertStatement(oDataAccess As CDataAccess, ByRef strInsertStatement As String) As CResult
    Set IDataSourceDatabase_BuildInsertStatement = GetInsertStatement(oDataAccess, strInsertStatement)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : IDataSourceDatabase_BuildUpdateStatement
'   Purpose     : Construct a full UPDATE statement, set local properties, and
'                 return string to calling routine.
'
'   Arguments   : CDataSourceDatabase           currentDb   The relevant database to build connection string
'                 strUpdateStatement    The returned Update Statement, if any
'
'   Returns     : String
'
'   Date        Developer               Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait            Created
'   2025/01/24  Jude Parfait            Add ByRef Arguments: strUpdateStatement; Change Return Type: CResult
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IDataSourceDatabase_BuildUpdateStatement(oDataAccess As CDataSourceDatabase, ByRef strUpdateStatement As String) As CResult
    Set IDataSourceDatabase_BuildUpdateStatement = GetUpdateStatement(oDataAccess, strUpdateStatement)
End Function
