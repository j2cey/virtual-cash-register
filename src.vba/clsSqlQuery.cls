Option Explicit

Private m_lId As Long
Private m_reqstTable As String

Private m_reqstCommand As ADODB.Command
Private m_reqstRecordset As New ADODB.Recordset
Private m_reqstStr As String

Private m_criterionProvidedCount As Long                        ' The number of criterion filled when the select is launched


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As clsSqlQuery
    Set Self = Me
End Property

' m_lId
Public Property Let id(lId As Long)
    m_lId = lId
End Property

Public Property Get id() As Long
    id = m_lId
End Property

' RequestTable
Public Property Let RequestTable(sRequestTable As String)
    m_reqstTable = sRequestTable
End Property

Public Property Get RequestTable() As String
    RequestTable = m_reqstTable
End Property

' RequestStr
Public Property Let RequestStr(sRequestStr As String)
    m_reqstStr = sRequestStr
End Property

Public Property Get RequestStr() As String
    RequestStr = m_reqstStr
End Property

' RequestCommand
Public Property Set RequestCommand(oRequestCommand As ADODB.Command)
    Set m_reqstCommand = oRequestCommand
End Property

Public Property Get RequestCommand() As ADODB.Command
    Set RequestCommand = m_reqstCommand
End Property

' RequestRecordset
Public Property Set reqstRecordset(oRequestRecordset As ADODB.Recordset)
    Set m_reqstRecordset = oRequestRecordset
End Property

Public Property Get reqstRecordset() As ADODB.Recordset
    Set reqstRecordset = m_reqstRecordset
End Property

' CriterionProvidedCount
Public Property Let CriterionProvidedCount(lcriterionProvidedCount As Long)
    m_criterionProvidedCount = lcriterionProvidedCount
End Property

Public Property Get CriterionProvidedCount() As Long
    CriterionProvidedCount = m_criterionProvidedCount
End Property


'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(sRequestTable As String)
    Me.RequestTable = sRequestTable
End Sub

Public Function SelectToListByCriterion(oResultList As clsUCTL, oSelectFields As Collection, Optional oCriterionFields As Collection = Nothing, Optional bSelectId As Boolean = False, Optional selectProgr As clsProgression, Optional lLimitLines As Long = 0) As Long
    SelectToListByCriterion = SelectToList(oResultList, oSelectFields, oCriterionFields, -1, bSelectId, selectProgr, lLimitLines)
End Function

Public Function SelectToListById(oResultList As clsUCTL, oSelectFields As Collection, Optional oCriterionFields As Collection = Nothing, Optional lIdCriteria As Long = -1) As Long
    SelectToListById = SelectToList(oResultList, oSelectFields, Nothing, lIdCriteria)
End Function

Public Function SelectToFieldsByCriterion(oResultFields As clsUCTLList, Optional oCriterionFields As Collection = Nothing) As Long
    SelectToFieldsByCriterion = SelectToFields(oResultFields, oCriterionFields)
End Function

Public Function SelectToFieldsById(oResultFields As clsUCTLList, lIdCriteria As Long, Optional selectProgr As clsProgression) As Long
    SelectToFieldsById = SelectToFields(oResultFields, Nothing, lIdCriteria, selectProgr)
End Function

Public Function SaveToTable(vSaveType As saveFormStatusType, oSaveFields As clsUCTLList, Optional lIdCriteria As Long = -1, Optional saveProgr As clsProgression) As Long
    Dim resultCount As Long, itemColl As Variant, oUctl As clsUCTL, lRecordsAffected As Long, lastId As Variant
    
    SaveToTable = -1
    
    StartNewSubProgressionFromParent "Enregistrement dans la Table " & Me.RequestTable, 1, saveProgr
    
    Call PrepareDatabase
    
    resultCount = 0
    
    If AddFieldsToSaveRequest(vSaveType, oSaveFields, saveProgr) > 0 Then
        
        If vSaveType = Add Then
            Me.RequestStr = Me.RequestStr & ";"
        Else
            If lIdCriteria < 0 Then
                MsgBox "Erreur Enregistrement " & vbCrLf & "Aucun Id fourni pour la Mise a Jour !", vbCritical, GetAppName
                
                AddDoneLastSubFromParent 1, True, saveProgr
                Exit Function
            End If
            Me.RequestStr = Me.RequestStr & " WHERE Id=" & CStr(lIdCriteria) & ";"
        End If
        
        'MsgBox Me.RequestStr
        
        On Error GoTo SaveTableError
        
        If gobjDB.OpenDB() Then
            
            Set Me.RequestCommand = New ADODB.Command
            BuildRequestCommand False, oSaveFields.GetDbFieldsSavable, Nothing, -1
            
            Me.RequestCommand.Execute lRecordsAffected
            
            If vSaveType = Add Then
                lastId = gobjDB.Connection.Execute("select @@identity").fields(0)
                
                SaveToTable = CLng(lastId)
            Else
                SaveToTable = lRecordsAffected
            End If
            
            gobjDB.CloseDB
        End If
    End If
    
    AddDoneLastSubFromParent 1, True, saveProgr
    
SaveTableResume:
    gobjDB.CloseDB
    Exit Function
SaveTableError:
    MsgBox "Erreur SaveToTable"
    Resume SaveTableResume
End Function

Public Function ExecDeleteFromId(lIdCriteria As Long, Optional delProgr As clsProgression) As Long
    Dim resultCount As Long, lRecordsAffected As Long
    Dim adoRS As Object
    
    StartNewSubProgressionFromParent "Exécution de la Requête DELETE", 1, delProgr
    
    ExecDeleteFromId = -1
    Call PrepareDatabase
    
    resultCount = 0
    
    Me.RequestStr = "DELETE FROM " & Me.RequestTable & " WHERE Id=" & CStr(lIdCriteria) & ";"
    
    'MsgBox Me.RequestStr
    
    If gobjDB.OpenDB() Then
        
        Set RequestCommand = New ADODB.Command
        
        With Me.RequestCommand
            .ActiveConnection = gobjDB.Connection
            .CommandType = adCmdText
            .CommandText = Me.RequestStr
            
        End With
        
        Me.RequestCommand.Execute lRecordsAffected
        
        ExecDeleteFromId = lRecordsAffected
        
        gobjDB.CloseDB
    End If
    
    AddDoneLastSubFromParent 1, True, delProgr
End Function



'*****************************************************************************************************************************
'    Private functions and subs
'*****************************************************************************************************************************

Private Function SelectToList(oResultList As clsUCTL, oSelectFields As Collection, Optional oCriterionFields As Collection = Nothing, Optional lIdCriteria As Long = -1, Optional bSelectId As Boolean = False, Optional selectProgr As clsProgression, Optional lLimitLines As Long = 0) As Long
    Dim resultCount As Long, itemColl As Variant, oField As clsDbField
    
    resultCount = 0
    ExecSelect oCriterionFields, lIdCriteria, selectProgr, lLimitLines
    
    StartNewSubProgressionFromParent "Transfert des Résultat dans la Liste", 1, selectProgr
    
    If Not (Me.reqstRecordset Is Nothing) Then
        If Not Me.reqstRecordset.EOF Then
            Do
                AddToDoLastSubFromParent 1, selectProgr
                ' Add Id for list selection
                If bSelectId Then
                    AddToResultList oResultList, resultCount, CStr(Me.reqstRecordset.fields("Id").value), 0
                End If
                If Not oSelectFields Is Nothing Then
                    For Each itemColl In oSelectFields
                        Set oField = itemColl
                        
                        If IsNull(Me.reqstRecordset.fields(oField.Name).value) Then
                            AddToResultList oResultList, resultCount, "", oField.Index + 1
                        Else
                            AddToResultList oResultList, resultCount, CStr(Me.reqstRecordset.fields(oField.Name).value), oField.Index + 1
                        End If
                    Next itemColl
                End If
                
                resultCount = resultCount + 1
                AddDoneLastSubFromParent 1, True, selectProgr
                
                Me.reqstRecordset.MoveNext
            Loop Until Me.reqstRecordset.EOF
        End If
    End If
    
    AddDoneLastSubFromParent 1, True, selectProgr
    
    gobjDB.CloseDB
    
    SelectToList = resultCount
End Function

Private Function SelectToFields(oResultFields As clsUCTLList, Optional oCriterionFields As Collection = Nothing, Optional lIdCriteria As Long = -1, Optional selectProgr As clsProgression) As Long
    Dim resultCount As Long, itemColl As Variant, oUctl As clsUCTL
    
    resultCount = 0
    ExecSelect oCriterionFields, lIdCriteria, selectProgr
    
    StartNewSubProgressionFromParent "Transfert des Résultat dans les Champs", 1, selectProgr
    
    If Not (Me.reqstRecordset Is Nothing) Then
        If Not Me.reqstRecordset.EOF Then
            Do
                AddToDoLastSubFromParent 1, selectProgr
                
                If Not oResultFields Is Nothing Then
                    For Each itemColl In oResultFields.ListCtls
                        Set oUctl = itemColl
                        
                        If Not IsNull(Me.reqstRecordset.fields(oUctl.field.Name).value) Then
                            oUctl.Text = CStr(Me.reqstRecordset.fields(oUctl.field.Name).value)
                        End If
                    Next itemColl
                End If
                
                resultCount = resultCount + 1
                AddDoneLastSubFromParent 1, True, selectProgr
                
                Me.reqstRecordset.MoveNext
            Loop Until Me.reqstRecordset.EOF
        End If
    End If
    
    AddDoneLastSubFromParent 1, True, selectProgr
    
    gobjDB.CloseDB
    
    SelectToFields = resultCount
End Function

Private Function SelectOneField(oSelectField As clsDbField, Optional oCriterionFields As Collection = Nothing, Optional lIdCriteria As Long = -1) As Variant
    Dim resultCount As Long
    
    resultCount = 0
    ExecSelect oCriterionFields, lIdCriteria
    
    If Not (Me.reqstRecordset Is Nothing) Then
    
        With Me.reqstRecordset
            .MoveNext
            If IsNull(.fields(oSelectField.Name).value) Then
                SelectOneField = CVar(.fields(oSelectField.Name).value)
            Else
                SelectOneField = Null
            End If
        End With
    End If
    
    gobjDB.CloseDB
    
End Function

Private Function ExecSelect(Optional oCriterionFields As Collection = Nothing, Optional lIdCriteria As Long = -1, Optional selectProgr As clsProgression, Optional lLimitLines As Long = 0) As Long
    Dim connectionOpened As Boolean, recAffected As Long
    
    StartNewSubProgressionFromParent "Exécution de la Requête SELECT", 1, selectProgr
    
    Call PrepareDatabase
    
    'Call RsltListInit
    
    BuildSelectCmdText oCriterionFields, lIdCriteria, lLimitLines
    connectionOpened = gobjDB.OpenDB()
    
    If connectionOpened Then
        
        BuildRequestCommand True, oCriterionFields, Nothing, lIdCriteria
        
        Set Me.reqstRecordset = Me.RequestCommand.Execute(recAffected)
        
        ExecSelect = 1
    Else
        MsgBox "Erreur Ouverture de la Base de Donnees"
        ExecSelect = -1
    End If
    
    AddDoneLastSubFromParent 1, True, selectProgr
    
    'MsgBox Me.RequestStr
End Function

Private Sub BuildSelectCmdText(Optional oCriterionFields As Collection = Nothing, Optional lIdCriteria As Long, Optional lLimitLines As Long = 0)
    Dim finalRqstStr As String

    Me.CriterionProvidedCount = 0
    Me.RequestStr = "SELECT * FROM " & Me.RequestTable
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
    
    'MsgBox Me.RequestStr
End Sub

Private Function AddFieldsToSaveRequest(vSaveType As saveFormStatusType, oSaveFields As clsUCTLList, Optional saveProgr As clsProgression) As Long
    Dim itemColl As Variant, oUctl As clsUCTL, lFieldCount As Long, sRqst As String, sInsertValues As String
    
    StartNewSubProgressionFromParent "Ajout de Champs a la Requête ENREGISTREMENT", 1, saveProgr
    
    If vSaveType = Add Then
        Me.RequestStr = "INSERT INTO " & Me.RequestTable
    Else
        Me.RequestStr = "UPDATE " & Me.RequestTable & " SET"
    End If
    
    lFieldCount = 0
    sRqst = ""
    sInsertValues = " VALUES("
    If oSaveFields Is Nothing Then
        MsgBox "Erreur Enregistrement " & vbCrLf & "Aucun champs fourni !", vbCritical, GetAppName
    Else
        For Each itemColl In oSaveFields.ListCtls
            
            AddToDoLastSubFromParent 1, saveProgr
            
            Set oUctl = itemColl
            
            If oUctl.field.IsRequired And oUctl.field.value = "" Then
                AddFieldsToSaveRequest = -1
                oUctl.BackColor = oUctl.ErrorColor
                MsgBox "Le Champs " & oUctl.field.Title & " est Requis !", vbCritical, GetAppName
                
                AddDoneLastSubFromParent 2, True, saveProgr
                Exit Function
            End If
            
            AppendFieldsToSaveRequest vSaveType, sRqst, oUctl.field.Name, lFieldCount
            If vSaveType = Add Then
                AppendFieldsToSaveRequest vSaveType, sInsertValues, "@" & oUctl.field.Name, lFieldCount
            End If
            
            AddDoneLastSubFromParent 1, True, saveProgr
            
            lFieldCount = lFieldCount + 1
        Next itemColl
    End If
    
    If vSaveType = Add Then
        Me.RequestStr = Me.RequestStr & " (" & sRqst & ") " & sInsertValues & ")"
    Else
        Me.RequestStr = Me.RequestStr & " " & sRqst
    End If
    
    AddDoneLastSubFromParent 1, True, saveProgr
    
    AddFieldsToSaveRequest = lFieldCount
End Function

Private Sub AppendFieldsToSaveRequest(vSaveType As saveFormStatusType, ByRef sRequest As String, s_critColName As String, lFieldCount As Long)
    
    If lFieldCount = 0 Then
        sRequest = sRequest & " "
    Else
        sRequest = sRequest & ", "
    End If
    
    If vSaveType = Add Then
        sRequest = sRequest & s_critColName
    Else
        sRequest = sRequest & s_critColName & " = @" & s_critColName
    End If
    
End Sub

Private Sub BuildRequestCommand(bOpenPrecision As Boolean, Optional oInputFields As Collection = Nothing, Optional oOutputFields As Collection = Nothing, Optional lIdCriteria As Long = -1)
    Dim itemColl As Variant, oField As clsDbField
    
    Set Me.RequestCommand = New ADODB.Command
        
    With Me.RequestCommand
        .ActiveConnection = gobjDB.Connection
        .CommandType = adCmdText
        .CommandText = Me.RequestStr
        
        ' add Id Parameter
        If Not (lIdCriteria = -1) Then
            .Parameters.Append .CreateParameter("@Id", adBigInt, adParamInput, 255, lIdCriteria)
        End If
        
        ' add Input Parameters if any
        If Not oInputFields Is Nothing Then
            For Each itemColl In oInputFields
                Set oField = itemColl
                
                If oField.value <> "" Then
                    If bOpenPrecision Then
                        .Parameters.Append .CreateParameter("@" & oField.Name, adVarChar, adParamInput, 255, "%" & oField.value & "%")
                    Else
                        .Parameters.Append .CreateParameter("@" & oField.Name, adVarChar, adParamInput, 255, oField.value)
                    End If
                End If
            Next itemColl
        End If
        
        ' add Output Parameters if any
        If Not oOutputFields Is Nothing Then
            For Each itemColl In oOutputFields
                Set oField = itemColl
                
                If oField.value <> "" Then
                    .Parameters.Append .CreateParameter("@" & oField.Name, adVarChar, adParamOutput, 255, "%" & oField.value & "%")
                End If
            Next itemColl
        End If
        
    End With
End Sub

Private Sub AddCriterionToRequest(oFieldList As Collection)
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

Private Sub StrRequestAppendCriteria(s_critColName As String)
    
    If Me.CriterionProvidedCount = 0 Then
        Me.RequestStr = Me.RequestStr & " WHERE "
    Else
        Me.RequestStr = Me.RequestStr & " AND "
    End If
    
    Me.RequestStr = Me.RequestStr & " " & s_critColName & " LIKE @" & s_critColName
    Me.CriterionProvidedCount = Me.CriterionProvidedCount + 1
    
End Sub

Private Sub ClearCriterionField()
    Dim itemColl As Variant, Uctl As clsUCTL
    
    If Me.FieldList Is Nothing Then
    Else
        For Each itemColl In Me.FieldList.ListCtls
            Set Uctl = itemColl
            
            If Uctl.IsCriteria Then
                Uctl.Clear
            End If
        Next itemColl
    End If
End Sub

Private Sub AddToResultList(oRsltList As clsUCTL, lrow As Long, sValue As String, lCol As Long)
    Select Case TypeName(oRsltList.ctl)
    Case "CheckBox"
        
    Case "ComboBox"
        AddToResultComboBox oRsltList, lrow, sValue
    Case "ListBox"
        AddToResultListBox oRsltList, lrow, sValue, lCol
    Case "OptionButton"
        
    Case "SpinButton"
        
    Case "TextBox"
        
    Case "Image"
        
    Case "Label"
        
    Case "Frame"
        
    End Select
End Sub

Private Sub AddToResultListBox(oRsltList As clsUCTL, lrow As Long, sValue As String, lCol As Long)
    Dim lbxList As MSForms.ListBox
    
    Set lbxList = oRsltList.ctl
    
    If (lbxList.listCount - 1) < lrow Then
        lbxList.AddItem
    End If
    
    lbxList.List(lrow, lCol) = sValue
End Sub

Private Sub AddToResultComboBox(oRsltList As clsUCTL, lrow As Long, sValue As String)
    Dim cbxList As MSForms.ComboBox
    
    Set cbxList = oRsltList.ctl
    
    If (cbxList.listCount - 1) < lrow Then
        ' New ID
        cbxList.AddItem sValue
        If cbxList.ColumnCount = 2 Then
            cbxList.List(lrow, 1) = sValue
        End If
    Else
        cbxList.List(lrow, 0) = sValue
    End If
    
End Sub
