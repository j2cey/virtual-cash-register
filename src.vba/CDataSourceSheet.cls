'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : CDataSourceSheet
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/02/05
' Purpose   : Excel Sheet Data Source
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_user As CModelUser

Private m_workbook_source As Workbook
Private m_current_sheet As Worksheet
Private m_current_cell_rownum As Long
Private m_current_cell_colnum As Long

Private Const MODULE_NAME As String = "Data-Source Sheet"

Implements IDataSource

'***************************************************************************************************************************************************************
'   Public Properties
'***************************************************************************************************************************************************************

Public Property Get Self() As CDataSourceSheet
    Set Self = Me
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : User
'   Purpose       : The performing user
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/05      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set User(oUser As CModelUser)
    Set m_user = oUser
End Property

Public Property Get User() As CModelUser
    Set User = m_user
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : WorkbookSource
'   Purpose       : The workbook source
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/05      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set WorkbookSource(oWorkbookSource As Workbook)
    Set m_workbook_source = oWorkbookSource
End Property

Public Property Get WorkbookSource() As Workbook
    Set WorkbookSource = m_workbook_source
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : CurrentSheet
'   Purpose       : The current worksheet
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/05      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set CurrentSheet(oCurrentSheet As Worksheet)
    Set m_current_sheet = oCurrentSheet
End Property

Public Property Get CurrentSheet() As Worksheet
    Set CurrentSheet = m_current_sheet
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : CurrentCellRownum
'   Purpose       : The current cell row number
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/05      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let CurrentCellRownum(lngCurrentCellRownum As Long)
    m_current_cell_rownum = lngCurrentCellRownum
End Property

Public Property Get CurrentCellRownum() As Long
    CurrentCellRownum = m_current_cell_rownum
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : CurrentCellColnum
'   Purpose       : The current cell column number
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/05      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let CurrentCellColnum(lngCurrentCellColnum As Long)
    m_current_cell_colnum = lngCurrentCellColnum
End Property

Public Property Get CurrentCellColnum() As Long
    CurrentCellColnum = m_current_cell_colnum
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : ServerOrPath
'   Purpose       : The server or database file path
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ServerOrPath(strServerOrPath As String)
End Property

Public Property Get ServerOrPath() As String
    ServerOrPath = Me.WorkbookSource.path
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : ServerOrPath
'   Purpose       : The server or database file path
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let DatabaseOrFileName(strDatabaseOrFileName As String)
End Property

Public Property Get DatabaseOrFileName() As String
    DatabaseOrFileName = Me.WorkbookSource.Name
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Property      : IntegratedSecurity
'   Purpose       : The server or database file path
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/18      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IntegratedSecurity(blnIntegratedSecurity As Boolean)
End Property

Public Property Get IntegratedSecurity() As Boolean
    IntegratedSecurity = True
End Property



'***************************************************************************************************************************************************************
'   Public Functions and Subs
'***************************************************************************************************************************************************************

Private Sub ActivateCurrentSheet()
    If Not Me.CurrentSheet Is ActiveSheet Then
        Me.CurrentSheet.Activate
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Subroutine    : Init
'   Purpose       : Initialize the Object
'   Arguments     : oUser               The Performing User
'
'   Returns       : Void
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/02/05      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(Optional oUser As CModelUser = Nothing)
    
    If Not oUser Is Nothing Then
        Set Me.User = oUser
    End If
End Sub

Public Function FirstBlankCell(Optional lngSourceCol As Long = 1, Optional ByVal StartRow& = 1, Optional ByVal SelectCell As Boolean = False) As Long
    Dim lngRowCount As Long, lngCurrentRow As Long
    Dim strCurrentRowValue As String
    Dim Data()
    
    If Me.CurrentSheet Is Nothing Then Set Me.CurrentSheet = ActiveSheet
    
    With Me.CurrentSheet
    
        lngRowCount = .Cells(.Rows.count, SourceCol).End(xlUp).Row
        Data = .Range(.Cells(1, SourceCol), .Cells(RowCount, SourceCol)).Value2
    
        For lngCurrentRow = StartRow To lngRowCount
            If Data(lngCurrentRow, lngSourceCol) = vbNullString Then
                If SelectCell Then .Cells(CurrentRow, SourceCol).Select
                'if selection is out of screen, intead of .select , use : application.goto reference:=.cells(...), scroll:= true
                FirstBlankCell = CurrentRow
                Exit For
            End If
        Next
    
    End With ' Sh
    
    Erase Data
    Set Sh = Nothing
End Function



'***************************************************************************************************************************************************************
'   Private Functions and Subs
'***************************************************************************************************************************************************************

Private Function ValidateDataAccessSheetParams(ByVal oDataAccess As CDataAccess, Optional ByVal blnSaveCurrentSheet As Boolean = False, Optional ByVal blnForSaving As Boolean = True) As CResult
    Dim result As CResult
    
    Set result = NewResult(MODULE_NAME, "ValidateDataAccessSheetParams", True)
    
    If oDataAccess Is Nothing Then
        result.SetFailure "Data Access not set yet !"
        Set ValidateDataAccessSheetParams = result
        
        Exit Function
    End If
    
    If blnForSaving Then
        If oDataAccess.TableForSaving = "" Then
            result.SetFailure "Table For Saving not set yet !"
            Set ValidateDataAccessSheetParams = result
            
            Exit Function
        End If
    Else
        If oDataAccess.TableForSelecting = "" Then
            result.SetFailure "Table For Selecting not set yet !"
            Set ValidateDataAccessSheetParams = result
            
            Exit Function
        End If
    End If
    
    If blnSaveCurrentSheet Then
        If blnForSaving Then
            Set Me.CurrentSheet = Me.WorkbookSource.Sheets(oDataAccess.TableForSaving)
        Else
            Set Me.CurrentSheet = Me.WorkbookSource.Sheets(oDataAccess.TableForSelecting)
        End If
    End If
    
    Set ValidateDataAccessSheetParams = result
End Function

Private Function SetCurrentCellParamsFromField(ByVal oField As CField) As CResult
    Dim result As CResult
    
    Set result = NewResult(MODULE_NAME, "SetCurrentCellParamsFromField", True)
    
    ' Assign the current Cell row number
    If oField.OffsetStart < 0 Then
        result.SetFailure "A normal Offset Start must be set for Row Number!"
        Set SetCurrentCellParamsFromField = result
        
        Exit Function
    Else
        Me.CurrentCellRownum = oField.OffsetStart
    End If
    
    ' Assign the current Cell column number
    If oField.OffsetEnd < 0 Then
        result.SetFailure "A normal Offset End must be set for Column Number!"
        Set SetCurrentCellParamsFromField = result
        
        Exit Function
    Else
        Me.CurrentCellColnum = oField.OffsetEnd
    End If
    
    Set SetCurrentCellParamsFromField = result
End Function

Private Function SetCellValueToField(ByVal oField As CField) As CResult
    Dim result As CResult
    
    Set result = NewResult(MODULE_NAME, "SetCellValueToField", True)
    
    AddSubResult MODULE_NAME, "SetCellValueToField", result, SetCurrentCellParamsFromField(oField)
    
    ' Set the cell value to Field
    If Me.CurrentSheet Is Nothing Then
        result.SetFailure "No CurrentSheet set yet !"
        Set SetCellValueToField = result
        
        Exit Function
    Else
        
        ActivateCurrentSheet
        
        oField.SetValue (Me.CurrentSheet.Cells(Me.CurrentCellRownum, Me.CurrentCellColnum))
        result.SetSuccess
    End If
    
    Set SetCellValueToField = result
End Function

Private Function SetFieldValueToCell(ByVal oField As CField) As CResult
    Dim result As CResult
    
    Set result = NewResult(MODULE_NAME, "SetFieldValueToCell", True)
    
    AddSubResult MODULE_NAME, "SetFieldValueToCell", result, SetCurrentCellParamsFromField(oField)
    
    If Not result.LastResult.IsOK Then
        result.SetFailure result.LastResult.Message
        Set SetFieldValueToCell = result
        
        Exit Function
    End If
    
    ' Set the Field value to cell
    If Me.CurrentSheet Is Nothing Then
        result.SetFailure "No CurrentSheet set yet !"
        Set SetFieldValueToCell = result
        
        Exit Function
    Else
        ActivateCurrentSheet
        
        Me.CurrentSheet.Cells(Me.CurrentCellRownum, Me.CurrentCellColnum) = CStr(oField.FieldValue.GetValue)
        
        saveFieldName oField
        
        result.SetSuccess
    End If
    
    Set SetFieldValueToCell = result
End Function

Private Sub saveFieldName(ByVal oField As CField)
    Dim lngCol As Long
    
    If Not oField.StoreFieldName = doNotStore Then
        ActivateCurrentSheet
        
        lngCol = -1
        If oField.StoreFieldName = toTheLeft Then
            lngCol = Me.CurrentCellColnum - 1
        ElseIf oField.StoreFieldName = toTheRight Then
            lngCol = Me.CurrentCellColnum + 1
        End If
        
        ' save the field name
        If lngCol > 0 Then
            Me.CurrentSheet.Cells(Me.CurrentCellRownum, lngCol) = CStr(oField.NameForSaving)
        End If
    End If
End Sub

Private Function SetSelectableRecord(ByVal oRecordTemplate As CRecord, ByRef oRecordValue As CRecord) As CResult
    Dim result As CResult, loopDone As Boolean
    
    Set result = NewResult(MODULE_NAME, "SetSelectableRecord", True)
    
    ' create new Record, if not done yet
    If oRecordValue Is Nothing Then
        Set oRecordValue = NewRecord(oRecordTemplate.DataAccess, oRecordTemplate.User)
    End If
    
    ' Reset Record field list's fields picked
    oRecordTemplate.FieldList.ResetFieldsPicked
    loopDone = False
    
    Do
        Dim currField As CField
        Set currField = oRecordTemplate.FieldList.GetNextField().ResetCode()
        AddSubResult MODULE_NAME, "SetSelectableRecord", result, SetCellValueToField(currField)
        
        If result.LastResult.IsOK Then
            oRecordValue.FieldList.AddField currField
        End If
        
        If oRecordTemplate.FieldList.LastFieldReached() Then
            ' The end of list is reached
            Exit Do
            loopDone = True
        End If
        
    Loop While (Not loopDone)
    
    Set SetSelectableRecord = result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetValue
'   Purpose     : Get Value from Sheet
'   Arguments   :
'
'   Returns     : CField
'
'   Date        Developer       Action
'   ---------------------------------------------------------------------------------------
'   2025/02/05  Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetValue(ByVal oDataAccess As CDataAccess, ByVal strFieldName As String, ByRef oResult As CResult) As CField
    Dim lngRowNum As Long, lngColNum As Long, oField As CField
    
    ' Validate Data Access Sheet parameters
    AddSubResult MODULE_NAME, "GetValue", oResult, ValidateDataAccessSheetParams(oDataAccess, True, False)
    If Not oResult.LastResult.IsOK Then
        Set GetValue = Nothing
        Exit Function
    End If
    
    ' Get the Field Template
    Set oField = oDataAccess.Record.FieldList.GetFieldByName(strFieldName)
    oField.ResetCode
    
    ' Set Current CEll value
    AddSubResult MODULE_NAME, "GetValue", oResult, SetCellValueToField(oField)
    If Not oResult.LastResult.IsOK Then
        Set GetValue = Nothing
        Exit Function
    Else
        Set GetValue = oField
    End If
    
End Function

Private Function GetRecord(ByVal oDataAccess As CDataAccess, ByRef oResult As CResult) As CRecord
    Dim result As CResult, oRecordValue As CRecord, loopDone As Boolean, oField As CField
    Dim blnSaveCurrentSheet As Boolean
    Dim blnForSaving As Boolean
    
    blnSaveCurrentSheet = True
    blnForSaving = False
    
    ' Initialize the result object
    AddSubResult MODULE_NAME, "GetRecord", result, oResult
    
    ' Validate Data Access Sheet parameters
    AddSubResult MODULE_NAME, "GetRecord", result, ValidateDataAccessSheetParams(oDataAccess, blnSaveCurrentSheet, blnForSaving)
    If Not result.LastResult.IsOK Then
        Set GetRecord = Nothing
        Exit Function
    End If
    
    ' Create the OUTPUT Record
    'Set oRecordValue = NewRecord(dataAccess, dataAccess.User)
    
    If result.LastResult.IsOK Then
        AddSubResult MODULE_NAME, "GetRecord", oResult, SetSelectableRecord(oDataAccess.Record, oRecordValue)
    End If
    
    If result.LastResult.IsOK Then
        result.SetSuccess
    Else
        result.SetFailure "Last Operation Failed: " & result.LastResult.Message
    End If
    
    Set GetRecord = oRecordValue
End Function


Public Function InsertValues(ByVal oDataAccess As CDataAccess) As CResult
    Dim result As CResult, loopDone As Boolean
    
    ' Validate Data Access Sheet parameters
    AddSubResult MODULE_NAME, "InsertValues", result, ValidateDataAccessSheetParams(oDataAccess, True, True)
    If Not result.LastResult.IsOK Then
        result.SetFailure "Last Operation Failed: " & result.LastResult.Message
        Set InsertValues = result
        Exit Function
    End If
    
    ' Reset Record field list's fields picked
    oDataAccess.Record.FieldList.ResetFieldsPicked
    loopDone = False
    
    Do
        Dim currField As CField
        Set currField = oDataAccess.Record.FieldList.GetNextField()
        
        AddSubResult MODULE_NAME, "InsertValues", result, SetFieldValueToCell(currField)
        
        If oDataAccess.Record.FieldList.LastFieldReached() Then
            ' The end of list is reached
            Exit Do
            loopDone = True
        End If
        
    Loop While (Not loopDone)
    
    result.SetSuccess
    Set InsertValues = result
End Function

Public Function UpdateRecord(ByVal oDataAccess As CDataAccess) As CResult
    Set UpdateRecord = InsertValues(oDataAccess)
End Function

Public Function DeleteRecord(ByVal oDataAccess As CDataAccess) As CResult
End Function



'***************************************************************************************************************************************************************
'   Interfaces Implementation
'***************************************************************************************************************************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : IDataSource_User
' Purpose       : The Record Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Set IDataSource_User(oUser As CModelUser)
    Set Me.User = oUser
End Property

Public Property Get IDataSource_User() As CModelUser
    Set IDataSource_User = Me.User
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : ServerOrPath
' Purpose       : The Record Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IDataSource_ServerOrPath(strServerOrPath As String)
    Me.ServerOrPath = strServerOrPath
End Property

Public Property Get IDataSource_ServerOrPath() As String
    IDataSource_ServerOrPath = Me.ServerOrPath
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : DatabaseOrFileName
' Purpose       : The Record Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IDataSource_DatabaseOrFileName(strDatabaseOrFileName As String)
    Me.DatabaseOrFileName = strDatabaseOrFileName
End Property

Public Property Get IDataSource_DatabaseOrFileName() As String
    IDataSource_DatabaseOrFileName = Me.DatabaseOrFileName
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property      : IntegratedSecurity
' Purpose       : The Record Name
'
' Date          Developer           Action
' ---------------------------------------------------------------------------------------
' 2025/02/05    Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let IDataSource_IntegratedSecurity(blnIntegratedSecurity As Boolean)
    Me.IntegratedSecurity = blnIntegratedSecurity
End Property

Public Property Get IDataSource_IntegratedSecurity() As Boolean
    IDataSource_IntegratedSecurity = Me.IntegratedSecurity
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function      : IDataSource_GetValue
'   Purpose       : Get a value from database (select scalar)
'
'   Date            Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/25      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IDataSource_GetValue(ByVal oDataAccess As CDataAccess, ByVal strFieldName As String, ByRef oResult As CResult) As CField
    Set IDataSource_GetValue = GetValue(oDataAccess, strFieldName, oResult)
End Function

Public Function IDataSource_GetRecord(ByVal oDataAccess As CDataAccess, ByRef oResult As CResult) As CRecord
    Set IDataSource_GetRecord = GetRecord(oDataAccess, oResult)
End Function

Public Function IDataSource_GetRecordList(ByVal oDataAccess As CDataAccess, ByRef oResult As CResult) As CRecordList
    Set IDataSource_GetRecordList = GetRecordList(oDataAccess, oResult)
End Function



Public Function IDataSource_StoreRecord(ByVal oDataAccess As CDataAccess) As CResult
    Set IDataSource_StoreRecord = InsertValues(oDataAccess)
End Function

Public Function IDataSource_UpdateRecord(ByVal oDataAccess As CDataAccess) As CResult
    Set IDataSource_UpdateRecord = UpdateRecord(oDataAccess)
End Function

Public Function IDataSource_DeleteRecord(ByVal oDataAccess As CDataAccess) As CResult
    Set IDataSource_DeleteRecord = DeleteRecord(oDataAccess)
End Function



'***************************************************************************************************************************************************************
'    Private events
'***************************************************************************************************************************************************************

Private Sub Class_Initialize()
    Set m_workbook_source = ActiveWorkbook
End Sub

Private Sub Class_Terminate()
    Set m_workbook_source = Nothing
End Sub

