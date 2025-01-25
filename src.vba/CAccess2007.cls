'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CAccess2007
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/19
' Purpose   : Access 2007 Class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


Private m_upper_database As CAccessDB
Private m_current_recordable As CRecordableDA

Private m_request_string As String
Private m_criteria_count As Long

Implements IAccessDB


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CAccess2007
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : UpperDatabase
'   Purpose       : The Upper Database
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/19      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set UpperDatabase(oUpperDatabase As CAccessDB)
    Set m_upper_database = oUpperDatabase
End Property

Public Property Get UpperDatabase() As CAccessDB
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
        If Me.UpperDatabase.CurrentServer = "" Then
            rslt.IsOK = False
            rslt.Message = "The CurrentServer is not set in the Upper Database"
        End If
        
        ' Exit Sub if the CurrentDatabase is not set in the Upper Database
        If rslt.IsOK And Me.UpperDatabase.CurrentDatabase = "" Then
            rslt.IsOK = False
            rslt.Message = "The CurrentDatabase is not set in the Upper Database"
        End If
        
        If rslt.IsOK Then
            'targetDbFullPath = AppPath & Application.PathSeparator & gblDbLocalFolder & Application.PathSeparator & gblDbName
            targetDbFullPath = Me.UpperDatabase.CurrentServer & Application.PathSeparator & Me.UpperDatabase.CurrentDatabase
            Me.UpperDatabase.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & _
            targetDbFullPath & ".accdb;"
            strConnectionString = Me.UpperDatabase.ConnectionString
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
'   Arguments   : CDatabaseDS   currentDb   The relevant database to build connection string
'
'   Returns     : String        The Build SELECT statement | empty string
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSelectStatement(currentDb As CRecordableDA, ByRef strSelectStatement As String, ByRef lngCriteriaCount As Long) As CResult
    Dim finalRqstStr As String, criterionDone As Boolean, rslt As CResult
    
    Me.RequestString = ""
    Me.CriteriaCount = 0
    Set Me.CurrentRecordable = currentDb
    Set rslt = NewResult("GetConnectionString", True)
    
    ' If the Table Name is not given, we return empty string
    If currentDb.RecordTable = "" Then
        rslt.IsOK = False
        rslt.Message = "The Table Name is not given"
    Else
        Me.RequestString = "SELECT * FROM " & currentDb.RecordTable
    End If
    
    If rslt.IsOK Then
        AddCriteriaToRequest
    End If
    
    If Me.CriterionProvidedCount = 0 And lLimitLines > 0 Then
        Me.RequestString = SqlSelectSetLIMIT(Me.RequestStr, lLimitLines)
    End If
    
    strSelectStatement = Me.RequestString
    lngCriteriaCount = Me.CriteriaCount
    Set GetSelectStatement = rslt
End Function

Private Sub AddCriteriaToRequest()
    Dim itemColl As Variant, oField As clsDbField, str As String
    
    If oFieldList Is Nothing Then
    Else
        For Each itemColl In oFieldList
            str = TypeName(itemColl)
            If Not (TypeName(itemColl) = "Nothing") Then
                Set oField = itemColl
                str = oField.Name
                
                If CStr(oField.value) <> "" Then
                    StrRequestAppendCriteria CStr(oField.Name)
                    Me.CriterionProvidedCount = Me.CriterionProvidedCount + 1
                End If
            End If
        Next itemColl
    End If
End Sub

Private Sub StrRequestAppendCriterion(oField As CField)
    
    If Me.CriterionProvidedCount = 0 Then
        Me.RequestStr = Me.RequestStr & " WHERE "
    Else
        Me.RequestStr = Me.RequestStr & " AND "
    End If
    
    Me.RequestStr = Me.RequestStr & " " & s_critColName & " LIKE @" & s_critColName
    Me.CriterionProvidedCount = Me.CriterionProvidedCount + 1
    
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
Private Function GetInsertStatement(currentDb As CDatabaseDS) As String
    GetInsertStatement = ""
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
Private Function GetUpdateStatement(currentDb As CDatabaseDS) As String
    GetUpdateStatement = ""
End Function



'***************************************************************************************************************************************************************
'   Interfaces Implementation
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : BuildConnectionString
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
Public Function IAccessDB_BuildConnectionString(ByRef strConnectionString As String) As CResult
    Set IAccessDB_BuildConnectionString = GetConnectionString(strConnectionString)
End Function

