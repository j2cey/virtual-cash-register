'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CDatabaseCommand
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/21
' Purpose   : Database Command Object
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_request_command As ADODB.Command

Private m_upper_database As CDataSourceDatabase
Private m_request_string As String
Private m_current_recordable As CDataAccess

Private m_last_result As CResult

Private Const MODULE_NAME As String = "Database Command"


'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CDatabaseCommand
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : UpperDatabase
'   Purpose       : The Upper Database
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set UpperDatabase(oUpperDatabase As CDataSourceDatabase)
    Set m_upper_database = oUpperDatabase
End Property

Public Property Get UpperDatabase() As CDataSourceDatabase
    Set UpperDatabase = m_upper_database
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : RequestCommand
' Purpose       : The Request Command
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/01/21    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set RequestCommand(oRequestCommand As ADODB.Command)
    Set m_request_command = oRequestCommand
End Property

Public Property Get RequestCommand() As ADODB.Command
    Set RequestCommand = m_request_command
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
'   Property      : CurrentDataAccess
'   Purpose       : The current DataAccess Object
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set CurrentDataAccess(oCurrentDataAccess As CDataAccess)
    Set m_current_recordable = oCurrentDataAccess
End Property

Public Property Get CurrentDataAccess() As CDataAccess
    Set CurrentDataAccess = m_current_recordable
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : ModuleName
'   Purpose       : The Name of this Module
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ModuleName() As String
    ModuleName = MODULE_NAME
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : LastResult
'   Purpose     : Get / Set The Last Result
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/01/21  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LastResult() As CResult
    Set LastResult = m_last_result
End Property

Public Property Set LastResult(oLastResult As CResult)
    Set m_last_result = oLastResult
End Property



'***************************************************************************************************************************************************************
'   Public Functions and Subs
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : oUpperDatabase          The Data Access object
'
'   Returns       : Void
'
'   Date            Developer               Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait            Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(oUpperDatabase As CDataSourceDatabase)
    
    Set Me.UpperDatabase = oUpperDatabase
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function      : BuildRequestCommand
'   Purpose       : Build the Request Command
'   Arguments     : oUpperDatabase          The Data Access object
'
'   Returns       : CResult
'
'   Date            Developer               Action
'   ---------------------------------------------------------------------------------------
'   2025/01/28      Jude Parfait            Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildRequestCommand(ByVal oDataAccess As CDataAccess, ByVal strRequestString As String, Optional lngIdCriteria As Long = -1) As CResult
    Dim itemColl As Variant
    
    Set Me.LastResult = NewResult(MODULE_NAME, "BuildRequestCommand", True)
    Set Me.RequestCommand = New ADODB.Command
    Set Me.CurrentDataAccess = oDataAccess
    Me.RequestString = strRequestString
    
    If Me.UpperDatabase Is Nothing Then
        Me.LastResult.IsOK = False
        Me.LastResult.Message = "The Upper Database Is Missing"
    Else
        
        With Me.RequestCommand
            .ActiveConnection = Me.UpperDatabase.Connection.Connection
            .CommandType = adCmdText
            .CommandText = Me.RequestString
            
            ' add Id Parameter
            If Not (lngIdCriteria = -1) Then
                .Parameters.Append .CreateParameter("@Id", adBigInt, adParamInput, 255, lngIdCriteria)
            End If
            
            ' add Parameters if any
            Dim rslt As CResult
            Set rslt = AddCommandParameters
            
            If Not rslt.IsOK Then
                Me.LastResult.IsOK = False
                Me.LastResult.Message = rslt.Message
            End If
        End With
    End If
    
    Set BuildRequestCommand = Me.LastResult
End Function

Public Function ExecuteCommand(ByRef lngRecAffected As Long) As ADODB.Recordset
    Set ExecuteCommand = Me.RequestCommand.Execute(lngRecAffected)
End Function



'***************************************************************************************************************************************************************
'   Private Functions and Subs
'***************************************************************************************************************************************************************

Private Function AddCommandParameters() As CResult
    Dim rslt As CResult, loopDone As Boolean, oField As CField
    
    Set rslt = NewResult(MODULE_NAME, "AddCommandParameters", True)
    loopDone = False
    
    If rslt.IsOK Then
        Me.CurrentDataAccess.Record.FieldList.ResetFieldsPicked
        Do
            Set oField = Me.CurrentDataAccess.Record.FieldList.GetNextFieldCriteria()
            
            If Not oField Is Nothing Then
                
                If Not oField.IsEmpty Then
                    ' Try to Add INPUT Parameter
                    Set rslt = oField.FieldValue.SetInputParameter(Me.RequestCommand)
                    
                    ' Try to Add OUTPUT Parameter
                    Set rslt = oField.FieldValue.SetOutputParameter(Me.RequestCommand)
                End If
            End If
            
            If Me.CurrentDataAccess.Record.FieldList.LastFieldReached() Then
                ' The end of list is reached
                Exit Do
                loopDone = True
            End If
            
        Loop While (Not loopDone)
    End If
    
    Set AddCommandParameters = rslt
End Function

Private Function AddInputParameter(ByVal oField As CField, ByRef oResult As CResult) As Boolean
    Set oResult = NewResult(MODULE_NAME, "AddInputParameter for field " & oField.Name, True)
    
    Set oResult = oField.FieldValue.SetInputParameter(Me.RequestCommand)
End Function


'***************************************************************************************************************************************************************
'    Private events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    Set m_request_command = New ADODB.Command
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
        Set m_request_command = Nothing
        Set m_request_recordset = Nothing
    On Error GoTo 0
End Sub
