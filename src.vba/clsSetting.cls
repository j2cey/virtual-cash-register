'Build 000
'*****************************************************************************************************************************
'
' Authors:  Jude Parfait NGOM NZE, jud10parfait@gmail.com
'
' (c)2024, all rights reserved to the authors
'
'*****************************************************************************************************************************
'-------------------------------------------------------------------------
' Module    : clsSetting
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 11-06-2024
' Purpose   : Manage App Setting
'-------------------------------------------------------------------------
Option Explicit


Private m_oParent As clsSetting

Private m_sLabel As String          ' setting Label
Private m_sSheetName As String      ' setting sheet name
Private m_iRowNum As Integer        ' setting row number
Private m_iColNum As Integer        ' setting column number

Private m_sKey As String            ' setting key
Private m_vValType As valueType     ' setting value type
Private m_vVal As Variant           ' setting value
Private m_arraySep As String        ' array separator. in case of array value

Private m_saveToDb As Boolean       ' determine if the setting must be save to DB too
Private m_tableName As String       ' settings table name
Private m_id As Integer             ' setting row id (in db table)
Private m_createdAt As Date         ' setting created date
Private m_updatedAt As Date         ' setting updated date


'*****************************************************************************************************************************
'    Public events
'*****************************************************************************************************************************



'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As clsSetting
    Set Self = Me
End Property

' Parent
Public Property Set Parent(oParent As clsSetting)
    Set m_oParent = oParent
End Property

Public Property Get Parent() As clsSetting
    Set Parent = m_oParent
End Property

' Label
Public Property Let Label(sLabel As String)
    m_sLabel = sLabel
End Property

Public Property Get Label() As String
    Label = m_sLabel
End Property

' SheetName
Public Property Let SheetName(sSheetName As String)
    m_sSheetName = sSheetName
End Property

Public Property Get SheetName() As String
    SheetName = m_sSheetName
End Property

' RowNum
Public Property Let RowNum(iRowNum As Integer)
    m_iRowNum = iRowNum
End Property

Public Property Get RowNum() As Integer
    RowNum = m_iRowNum
End Property

' ColNum
Public Property Let ColNum(iColNum As Integer)
    m_iColNum = iColNum
End Property

Public Property Get ColNum() As Integer
    ColNum = m_iColNum
End Property

' Key
Public Property Let Key(sKey As String)
    m_sKey = sKey
End Property

Public Property Get Key() As String
    Key = m_sKey
End Property

' ValType
Public Property Let ValType(vValType As valueType)
    m_vValType = vValType
End Property

Public Property Get ValType() As valueType
    ValType = m_vValType
End Property

' Val
Public Property Let Val(vVal As Variant)
    m_vVal = vVal
End Property

Public Property Get Val() As Variant
    Val = m_vVal
End Property

' ArraySep
Public Property Let ArraySep(sArraySep As String)
    m_arraySep = sArraySep
End Property

Public Property Get ArraySep() As String
    ArraySep = m_arraySep
End Property

' SaveToDb
Public Property Let SaveToDb(bSaveToDb As Boolean)
    m_saveToDb = bSaveToDb
End Property

Public Property Get SaveToDb() As Boolean
    SaveToDb = m_saveToDb
End Property

' TableName
Public Property Let TableName(sTableName As String)
    m_tableName = sTableName
End Property

Public Property Get TableName() As String
    TableName = m_tableName
End Property

' Id
Public Property Let Id(iId As Integer)
    m_id = iId
End Property

Public Property Get Id() As Integer
    Id = m_id
End Property

' CreatedAt
Public Property Let CreatedAt(dCreatedAt As Date)
    m_createdAt = dCreatedAt
End Property

Public Property Get CreatedAt() As Date
    CreatedAt = m_createdAt
End Property

' UpdatedAt
Public Property Let UpdatedAt(dUpdatedAt As Date)
    m_updatedAt = dUpdatedAt
End Property

Public Property Get UpdatedAt() As Date
    UpdatedAt = m_updatedAt
End Property



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(sKey As String, sLabel As String, sSheetName As String, iRowNum As Integer, iColNum As Integer, vValType As valueType, defaultValue As Variant, bSaveToDb As Boolean, Optional oParent As clsSetting)
    Me.Label = sLabel
    Me.SheetName = sSheetName
    Me.RowNum = iRowNum
    Me.ColNum = iColNum
    Me.Key = sKey
    Me.SaveToDb = bSaveToDb
    
    Me.ValType = vValType
    
    If Not IsMissing(oParent) Then
        Set Me.Parent = oParent
    End If
    
    ReadValue
    
    If IsEmpty(Me.Val) Then
        Me.SaveValue defaultValue
    End If
    
End Sub

Public Sub SaveValue(vVal As Variant)
    
    If Me.Key = "" Then
        MsgBox "Key not set for this Setting ! ", vbCritical, GetAppName
        Exit Sub
    End If
    If Me.Label = "" Then
        MsgBox "No Label set for this Setting " & Me.Key & " ! ", vbCritical, GetAppName
        Exit Sub
    End If
    
    If Not SaveValueToSheet(vVal) Then
        Exit Sub
    End If
    
    If Me.SaveToDb Then
        SaveValueToDB (vVal)
    End If
End Sub

Public Function GetArrayValue() As String()
    If (Not IsEmpty(Me.ArraySep)) And (Not Me.ArraySep = "") Then
        
        GetArrayValue = Split(CStr(Me.Val), Me.ArraySep)
        Exit Function
    End If
    GetArrayValue = Split(CStr(Me.Val), "")
End Function



'*****************************************************************************************************************************
'    Private functions and subs
'*****************************************************************************************************************************

Private Function GetFormattedValue(vVal As Variant)
    If Me.ValType = boolval Then
        If LCase(CStr(vVal)) = "true" Or LCase(CStr(vVal)) = "vrai" Or CStr(vVal) = "1" Then
            GetFormattedValue = 1
        Else
            GetFormattedValue = 0
        End If
        Exit Function
    End If
    GetFormattedValue = vVal
End Function

Private Function SaveValueToSheet(vVal As Variant) As Boolean
    Dim ws As Worksheet
    
    If Me.SheetName = "" Then
        MsgBox "No Sheet Name set for this Setting " & Me.Key & " ! ", vbCritical, GetAppName
        
        SaveValueToSheet = False
        Exit Function
    End If
    If Me.RowNum = -1 Then
        MsgBox "No Row Number set for this Setting " & Me.Key & " ! ", vbCritical, GetAppName
        
        SaveValueToSheet = False
        Exit Function
    End If
    If Me.ColNum = -1 Then
        MsgBox "No Row Number set for this Setting " & Me.Key & " ! ", vbCritical, GetAppName
        
        SaveValueToSheet = False
        Exit Function
    End If
    
    Set ws = ActiveWorkbook.Sheets(Me.SheetName)
    ws.Activate
    
    Me.Val = GetFormattedValue(vVal)
    ws.Cells(Me.RowNum, Me.ColNum) = CStr(Me.Val)
    
    SaveValueToSheet = True
End Function

Private Function SaveValueToDB(vVal As Variant) As Boolean
    
    If Me.Id < 1 Then
        ReadValueFromDB
    End If
    
    If Me.Id < 1 Then
        SaveValueToDB = InsertSettingToDB(vVal)
    Else
        SaveValueToDB = UpdateSettingToDB(vVal)
    End If
    
End Function

Private Function InsertSettingToDB(vVal As Variant) As Boolean
    Dim queryStr As String, newSettingId As Variant
    
    ' Add New Setting
    queryStr = "INSERT INTO " & Me.TableName & " (settingkey,settingvalue,created_at,updated_at) VALUES ('" & CStr(Me.Key) & "', '" & CStr(vVal) & "', '" & CStr(Now) & "', '" & CStr(Now) & "'" & ")"
    
    If gobjDB.ExecuteActionQuery(queryStr, newSettingId) Then
        Me.Id = newSettingId
        Me.Val = GetFormattedValue(vVal)
        Me.CreatedAt = Now
        Me.UpdatedAt = Now
        
        InsertSettingToDB = True
    Else
        InsertSettingToDB = False
    End If
    
End Function

Private Function UpdateSettingToDB(vVal As Variant) As Boolean
    Dim settingData As Variant, sqlRst As String, recordCount As Long, i As Integer
    
    sqlRst = "UPDATE " & Me.TableName & " SET settingvalue = '" & CStr(vVal) & "', updated_at = '" & Now & "' WHERE Id = " & CStr(Me.Id)
    
    Call PrepareDatabase
    If gobjDB.ExecuteActionQuery(sqlRst) Then
        Me.Val = GetFormattedValue(vVal)
        Me.UpdatedAt = Now
        
        UpdateSettingToDB = True
    Else
        UpdateSettingToDB = False
    End If
End Function


Private Sub ReadValue()
    Dim settingReadOk As Boolean
    
    If Me.SaveToDb Then
        settingReadOk = ReadValueFromDB
    End If
    
    If Not settingReadOk Then
        ReadValueFromSheet
        
        If Me.SaveToDb Then
            SaveValue Me.Val
        End If
    End If
    
End Sub

Private Sub ReadValueFromSheet()
    Dim ws As Worksheet
    
    If Me.Key = "" Then
        Exit Sub
    End If
    If Me.Label = "" Then
        Exit Sub
    End If
    If Me.SheetName = "" Then
        Exit Sub
    End If
    If Me.RowNum = -1 Then
        Exit Sub
    End If
    If Me.ColNum = -1 Then
        Exit Sub
    End If
    
    Set ws = ActiveWorkbook.Sheets(Me.SheetName)
    ws.Activate
    
    Me.Val = CVar(ws.Cells(Me.RowNum, Me.ColNum))
    
End Sub

Public Function ReadValueFromDB() As Boolean
    Dim settingData As Variant, sqlRst As String, recordCount As Long, i As Integer
    
    sqlRst = "SELECT * FROM " & Me.TableName & " WHERE settingkey = '" & Me.Key & "' "
    
    Call PrepareDatabase
    settingData = gobjDB.GetRecordsetToArray(sqlRst, recordCount)
    
    If recordCount = 0 Then
        ReadValueFromDB = False
    Else
        Me.Id = CLng(settingData(0, 0))
        Me.Val = CStr(settingData(2, 0))
        Me.CreatedAt = CDate(settingData(3, 0))
        Me.UpdatedAt = CDate(settingData(4, 0))
        
        ReadValueFromDB = True
    End If
End Function


'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    m_sLabel = ""
    m_sSheetName = ""
    m_iRowNum = -1
    m_iColNum = -1
    
    m_sKey = ""
    m_tableName = "settings"
End Sub
