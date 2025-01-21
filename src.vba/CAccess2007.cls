'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CAccess2007
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/19
' Purpose   : Access 2007 Class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


Private m_upper_database As CAccessDB

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


'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : BuildConnectionString
'   Purpose     : Construct a full connection string, set local properties, and return string to calling routine.
'   Arguments   :
'
'   Returns     : String
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/12  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildConnectionString() As String
    Dim targetDbFullPath As String
    
    If Not Me.UpperDatabase Is Nothing Then
        'Build connection string
        
        ' Exit Sub if the CurrentServer is not set in the Upper Database
        If Me.UpperDatabase.CurrentServer = "" Then
            BuildConnectionString = ""
            Exit Sub
        End If
        
        ' Exit Sub if the CurrentDatabase is not set in the Upper Database
        If Me.UpperDatabase.CurrentDatabase = "" Then
            BuildConnectionString = ""
            Exit Sub
        End If
        
        'targetDbFullPath = AppPath & Application.PathSeparator & gblDbLocalFolder & Application.PathSeparator & gblDbName
        targetDbFullPath = Me.UpperDatabase.CurrentServer & Application.PathSeparator & Me.UpperDatabase.CurrentDatabase
        Me.UpperDatabase.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & _
        targetDbFullPath & ".accdb;"
        
        BuildConnectionString = Me.UpperDatabase.ConnectionString
    Else
        BuildConnectionString = ""
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : BuildSelectStatement
'   Purpose     : Construct a full SELECT statement, set local properties, and
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
Public Function BuildSelectStatement(currentDb As CRecordableDA, ByRef lngCriterionCount As Long) As String
    Dim finalRqstStr As String

    lngCriterionCount = 0
    finalRqstStr = "SELECT * FROM " & cu
    If Not (lIdCriteria = -1) Then
        Me.CriterionProvidedCount = 1
        Me.RequestStr = Me.RequestStr & " WHERE Id = @Id"
    End If
    
    If Not oCriterionFields Is Nothing Then
        AddCriterionToRequest oCriterionFields
    End If
    
    If Me.CriterionProvidedCount = 0 And lLimitLines > 0 Then
        Me.RequestStr = SqlSelectSetLIMIT(Me.RequestStr, lLimitLines)
    End If
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
'   Returns     : String
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/19  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IAccessDB_BuildConnectionString() As String
    IAccessDB_BuildConnectionString = Me.BuildConnectionString()
End Function

