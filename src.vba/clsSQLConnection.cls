Option Explicit

'Module-level objects
Private mobjConn As ADODB.Connection
'Module-level variables
Private mblnIntegratedSecurity As Boolean
Private mstrModuleName As String
Private mstrConnectionString As String
Private mstrCurrentServer As String
Private mstrCurrentDatabase As String
Private mstrDatabaseType As String

'***************************
' CONCTRUCTOR & DESTRUCTOR *
'***************************

Private Sub Class_Initialize()
    mstrModuleName = "clsSQLConnection"
    Set mobjConn = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
        mobjConn.Close
        Set mobjConn = Nothing
    On Error GoTo 0
End Sub


'***************************
' PROPERTIES               *
'***************************


Public Property Get ConnectionString() As String
    ConnectionString = mstrConnectionString
End Property

Public Property Let ConnectionString(ByVal newValue As String)
    mstrConnectionString = newValue
End Property

Public Property Get CommandTimeOut() As Integer
    CommandTimeOut = mobjConn.CommandTimeOut
End Property

Public Property Let CommandTimeOut(ByVal NewTimeOut As Integer)
    mobjConn.CommandTimeOut = NewTimeOut
End Property

Public Property Get CurrentDatabase() As Integer
    'Read-only property
    CurrentDatabase = mstrCurrentDatabase
End Property

Public Property Get CurrentServer() As String
    'Read-only property
    CurrentServer = mstrCurrentServer
End Property

Public Property Get DatabaseType() As String
    DatabaseType = mstrDatabaseType
End Property

Public Property Let DatabaseType(ByVal newValue As String)
    mstrDatabaseType = newValue
End Property

Public Property Get IntegratedSecurity() As Boolean
    IntegratedSecurity = mblnIntegratedSecurity
End Property

Public Property Let IntegratedSecurity(ByVal newValue As Boolean)
    mblnIntegratedSecurity = newValue
End Property

Public Property Get Connection() As ADODB.Connection
    Set Connection = mobjConn
End Property




'***************************
' PUBLIC SUB & FUNCTIONS   *
'***************************

Public Sub CloseDB()
    'Close connection to database.
    On Error Resume Next
        If Not mobjConn Is Nothing Then
            If mobjConn.State And adStateOpen Then
                'Connection is defined and still open; close it.
                mobjConn.Close
            End If
        End If
    On Error GoTo 0
End Sub

Public Sub ShowErrorMessages(ByVal StandardErrorObject As VBA.ErrObject, _
    ByVal SourceModule As String, ByVal SourceMethod As String)
    'Construct a comprehensive error message based on the passed objects.
    Dim strMsg As String
    'Handle the standard error, if any.
    If StandardErrorObject.Number <> 0 Then
        strMsg = "Error: " & CStr(Err.Number) & vbCrLf & vbCrLf & "Description: " & _
        Err.Description & vbCrLf & vbCrLf
    End If
    'Include the database-related errors, if any.
    If Not (gobjDB Is Nothing) Then
        strMsg = strMsg & gobjDB.ADOErrors()
    End If
    'Finally, tack on the module/method names.
    strMsg = strMsg & SourceModule & "::" & SourceMethod & vbCrLf
    
    Err.Clear
    MsgBox strMsg, vbCritical, "Contact Technical Support For Assistance"
End Sub


Public Function ADOErrors() As String
    'Return a fully formatted string containing any current ADO errors in the collection.
    Dim lngCount As Long
    Dim objErr As ADODB.Error
    Dim strReturn As String
    
    strReturn = vbNullString
    
    For Each objErr In mobjConn.Errors
        strReturn = strReturn & CStr(objErr.Number) _
        & vbCrLf & objErr.Description & vbCrLf & vbCrLf
    Next objErr
    
    mobjConn.Errors.Clear
    ADOErrors = strReturn
End Function

Public Function OpenDB(Optional ByVal DatabaseType As String, _
Optional ByVal TargetServer As String = "", _
Optional ByVal TargetDatabase As String = "", _
Optional ByVal IntegratedSecurity As Boolean = False) As Boolean
    'Open a connection to the specified server and database, if not already open.
    'If the parameters are not specified, then just re-use the last connection
    ' string that was created.
    Dim blnNewConnect As Boolean, blnReturn As Boolean
    Dim strDataSource As String
    Dim strDB As String
    Dim strConn As String

    blnReturn = True
    blnNewConnect = True
    
    On Error GoTo ODError
        If Not mobjConn Is Nothing Then
            If mobjConn.State And adStateOpen Then
                'Connection is alread defined and opened.
                blnNewConnect = False
            End If
        End If
        
        If blnNewConnect Then
            'Must create a new connection
            If TargetServer = "" Or TargetDatabase = "" Then
                'Calling routine has not specified Server or Database values;
                ' default to last used.
                If Len(mstrConnectionString) > 0 Then
                    strConn = mstrConnectionString
                Else
                    'If full connection string is blank, use the individual Server/Environment and
                    ' Database variables to determine connection.
                    strConn = BuildConnectionString(mstrDatabaseType, mstrCurrentServer, _
                    mstrCurrentDatabase, mblnIntegratedSecurity)
                End If
            Else
                'Build connection string, based on passed arguments
                strConn = BuildConnectionString(DatabaseType, TargetServer, TargetDatabase, _
                mblnIntegratedSecurity)
            End If
            
            mobjConn.ConnectionString = strConn
            mobjConn.ConnectionTimeout = 30
            mobjConn.Open
        End If
        
ODResume:
    OpenDB = blnReturn
    Exit Function
ODError:
    blnReturn = False
    ShowErrorMessages Err, mstrModuleName, "OpenDB"
    Resume ODResume
End Function

Public Function BuildConnectionString(ByVal DatabaseType As String, _
ByVal TargetServer As String, _
ByVal TargetDatabase As String, _
ByVal IntegratedSecurity As Boolean, _
Optional ByVal UserID As String, _
Optional ByVal Password As String) As String
    Dim targetDbFullPath As String
    'Construct a full connection string, set local properties, and
    ' return string to calling routine.
    'Build connection string
    Select Case DatabaseType
        Case "SQLServer" 'Microsoft SQL Server; use proper network library.
            mstrConnectionString = "driver={SQL Server}" & _
            ";server=" & TargetServer & _
            ";uid=" & UserID & _
            ";pwd=" & Password & _
            ";database=" & TargetDatabase
            If IntegratedSecurity Then
                mstrConnectionString = "Provider=SQLNCLI10;" & _
                "Server=" & TargetServer & ";" & _
                "Database=" & TargetDatabase & ";" & _
                "Trusted_Connection=yes;"
            End If
        Case "Access2003" 'Access 2003 and prior
            mstrConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & ThisWorkbook.path & "\" & _
            TargetDatabase & ".mdb;"
        Case "Access2007_old" 'Access 2007 or 2010
            targetDbFullPath = AppPath & Application.PathSeparator & gblDbLocalFolder & Application.PathSeparator & TargetDatabase
            mstrConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & _
            targetDbFullPath & ".accdb;Persist Security Info=Flase;"
        Case "Access2007" 'Access 2007 or 2010
            'targetDbFullPath = AppPath & Application.PathSeparator & gblDbLocalFolder & Application.PathSeparator & gblDbName
            targetDbFullPath = TargetServer & Application.PathSeparator & TargetDatabase
            mstrConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & _
            targetDbFullPath & ".accdb;"
        Case Else
            'Some other server; modify to accommodate.
            mstrConnectionString = "Provider=SQLNCLI10;Server=" & TargetServer & _
            ";Database=" & TargetDatabase & ";Trusted_Connection=yes;"
    End Select
    
    'Update the module variables
    mstrCurrentServer = TargetServer
    mstrCurrentDatabase = TargetDatabase
    mblnIntegratedSecurity = IntegratedSecurity
    BuildConnectionString = mstrConnectionString
End Function

Public Function ExecuteSPWithADOCommand(StoredProcName As String, _
OutputParameter As String, OutputValue As Variant, _
ParamArray InputParameters() As Variant) As Boolean
    'Execute a stored procedure, using an ADO command object. The InputParameters
    ' paramarray must contain pairs of data (i.e. element 0 should be Parameter name,
    ' with leading '@', and element 1 should be the associated parameter value; then
    ' element 2 is the next parameter name, element 3 is the next parameter value, etc.).
    ' Also allows for a single Output parameter and associated return value
    ' (OutputParmeter and OutputValue ByRef parameters). Remember
    ' to prefix the OutputParmeter name with '@', same as the input parameters.
    Dim objCmd As ADODB.Command
    Dim blnReturn As Boolean
    Dim intParam As Integer
    Dim strParamName As String
    Dim vntParamValue As Variant
    blnReturn = True
    
    On Error GoTo ESPError
        If OpenDB() Then
            Set objCmd = New ADODB.Command
            
            With objCmd
                .ActiveConnection = mobjConn
                .CommandText = StoredProcName
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                
                'Define the Input Parameters
                For intParam = 0 To UBound(InputParameters) Step 2
                    .Parameters(InputParameters(intParam)).value = _
                    InputParameters(intParam + 1)
                Next intParam
                
                'Define the single allowed output parameter, if any
                If Len(Trim(OutputParameter)) > 0 Then
                    .Parameters(OutputParameter).Direction = adParamOutput
                End If
                'Execute the stored procedure
                .Execute
            End With
            
            'Retrieve the value of the output parameter (if any)
            If Len(Trim(OutputParameter)) > 0 Then
                OutputValue = objCmd.Parameters(OutputParameter).value
            Else
                OutputValue = vbNullString
            End If
        Else
        
            'No database connection could be established, but no error
            ' was raised (should never happen).
            OutputValue = vbNullString
            blnReturn = False
        End If
        
ESPResume:
    ExecuteSPWithADOCommand = blnReturn
    Set objCmd = Nothing
    CloseDB
Exit Function
ESPError:
    blnReturn = False
    ShowErrorMessages Err, mstrModuleName, "ExecuteSPWithADOCommand"
    Resume ESPResume
End Function

Public Sub PickNums(nbResults As Integer, nbLaureates As Integer, nbResultsLib As Integer, nbResultsMoo As Integer, var As String)
    Dim rst As New ADODB.Recordset
    Dim objCmd As ADODB.Command
    'Dim Rs As New ADODB.RecordSet
    'Dim StrSproc As String
    Dim i As Long
    Dim strResult As String

    On Error GoTo PickNumsError
    strResult = ""
    If OpenDB() Then
        Set objCmd = New ADODB.Command
        
        With objCmd
            .ActiveConnection = mobjConn
            .CommandType = adCmdStoredProc
            .CommandText = "sp_pickNums"
            .Parameters.Append .CreateParameter("@nbResults", adInteger, adParamInput, , nbResults)
            .Parameters.Append .CreateParameter("@nbLaureates", adInteger, adParamInput, , nbLaureates)
            .Parameters.Append .CreateParameter("@nbResultsLib", adInteger, adParamInput, , nbResultsLib)
            .Parameters.Append .CreateParameter("@nbResultsMoo", adInteger, adParamInput, , nbResultsMoo)
            .Parameters.Append .CreateParameter("@var", adChar, adParamInput, 500, var)
        End With
    
        'With Rs
        '    .CursorType = adOpenStatic
        '    .CursorLocation = adUseClient
        '    .LockType = adLockOptimistic
        '    .Open objCmd
        'End With
        
        Set rst = objCmd.Execute(, , adAsyncExecute)
        
        Do While objCmd.State = adStateExecuting 'or 4
            DoEvents
            'Options here include a timer to update the form or putting a cancel button
            'on the sheet
            'so the application does not appear to be frozen.
        Loop
        
        If Not (rst Is Nothing) Then
            If Not rst.EOF Then
                Do
                    'For i = 0 To rst.fields.Count - 1
                    '    MsgBox CStr(rst.fields(i).Name) & ": " & CStr(rst.fields(i).Value)
                    'Next i
                    
                    If strResult = "" Then
                        strResult = CStr(rst.fields("msisdn").value) & " - " & CStr(rst.fields("nom").value) & " " & CStr(rst.fields("prenom").value)
                    Else
                        strResult = strResult & "," & CStr(rst.fields("msisdn").value) & " - " & CStr(rst.fields("nom").value) & " " & CStr(rst.fields("prenom").value)
                    End If
                    rst.MoveNext
                Loop Until rst.EOF
            End If
        End If
    End If
    
    pickedNums = strResult
    
PickNumsResume:
    'PickNums = strResult
    Set objCmd = Nothing
    CloseDB
Exit Sub
PickNumsError:
    strResult = ""
    ShowErrorMessages Err, mstrModuleName, "PickNums"
    Resume PickNumsResume
End Sub

Public Sub AddTirageDone(resultsRequest As Integer, resultsResponse As Integer, _
laureatesRequest As Integer, laureatesResponse As Integer, _
resultsLibRequest As Integer, resultsLibResponse As Integer, _
resultsMooRequest As Integer, resultsMooResponse As Integer, _
var As String)
    Dim rst As New ADODB.Recordset
    Dim objCmd As ADODB.Command
    'Dim Rs As New ADODB.RecordSet
    'Dim StrSproc As String
    Dim i As Long
    Dim strResult As String

    On Error GoTo AddTDError
    strResult = ""
    If OpenDB() Then
        Set objCmd = New ADODB.Command
        
        With objCmd
            .ActiveConnection = mobjConn
            .CommandType = adCmdStoredProc
            .CommandText = "sp_addTirageDone"
            .Parameters.Append .CreateParameter("@resultsRequest", adInteger, adParamInput, , resultsRequest)
            .Parameters.Append .CreateParameter("@resultsResponse", adInteger, adParamInput, , resultsResponse)
            .Parameters.Append .CreateParameter("@laureatesRequest", adInteger, adParamInput, , laureatesRequest)
            .Parameters.Append .CreateParameter("@laureatesResponse", adInteger, adParamInput, , laureatesResponse)
            .Parameters.Append .CreateParameter("@resultsLibRequest", adInteger, adParamInput, , resultsLibRequest)
            .Parameters.Append .CreateParameter("@resultsLibResponse", adInteger, adParamInput, , resultsLibResponse)
            .Parameters.Append .CreateParameter("@resultsMooRequest", adInteger, adParamInput, , resultsMooRequest)
            .Parameters.Append .CreateParameter("@resultsMooResponse", adInteger, adParamInput, , resultsMooResponse)
            .Parameters.Append .CreateParameter("@var", adChar, adParamInput, 500, var)
        End With
        
        Set rst = objCmd.Execute
        
    End If
    
AddTDResume:
    'PickNums = strResult
    Set objCmd = Nothing
    CloseDB
Exit Sub
AddTDError:
    strResult = ""
    ShowErrorMessages Err, mstrModuleName, "AddTirageDone"
    Resume AddTDResume
End Sub

Public Sub StartTirage(libelle As String, descriptionTirage As String, _
resultsRequest As Integer, laureatesRequest As Integer, _
resultsLibRequest As Integer, resultsMooRequest As Integer, _
var As String)
    Dim rst As New ADODB.Recordset
    Dim objCmd As ADODB.Command
    'Dim Rs As New ADODB.RecordSet
    'Dim StrSproc As String
    Dim i As Long
    Dim strResult As String

    On Error GoTo StartTError
    strResult = ""
    If OpenDB() Then
        Set objCmd = New ADODB.Command
        
        With objCmd
            .ActiveConnection = mobjConn
            .CommandType = adCmdStoredProc
            .CommandText = "sp_startTirage"
            .Parameters.Append .CreateParameter("@libelle", adChar, adParamInput, 50, libelle)
            .Parameters.Append .CreateParameter("@description_tirage", adChar, adParamInput, 100, descriptionTirage)
            .Parameters.Append .CreateParameter("@resultsRequest", adInteger, adParamInput, , resultsRequest)
            .Parameters.Append .CreateParameter("@laureatesRequest", adInteger, adParamInput, , laureatesRequest)
            .Parameters.Append .CreateParameter("@resultsLibRequest", adInteger, adParamInput, , resultsLibRequest)
            .Parameters.Append .CreateParameter("@resultsMooRequest", adInteger, adParamInput, , resultsMooRequest)
            .Parameters.Append .CreateParameter("@var", adChar, adParamInput, 500, var)
        End With
        
        Set rst = objCmd.Execute
        
    End If
    
StartTResume:
    Set objCmd = Nothing
    CloseDB
Exit Sub
StartTError:
    strResult = ""
    ShowErrorMessages Err, mstrModuleName, "StartTirage"
    Resume StartTResume
End Sub

Public Sub EndTirage(var As String)
    Dim rst As New ADODB.Recordset
    Dim objCmd As ADODB.Command
    'Dim Rs As New ADODB.RecordSet
    'Dim StrSproc As String
    Dim i As Long
    Dim strResult As String

    On Error GoTo EndTError
    strResult = ""
    If OpenDB() Then
        Set objCmd = New ADODB.Command
        
        With objCmd
            .ActiveConnection = mobjConn
            .CommandType = adCmdStoredProc
            .CommandText = "sp_endTirage"
            .Parameters.Append .CreateParameter("@var", adChar, adParamInput, 500, var)
        End With
        
        Set rst = objCmd.Execute
        
    End If
    
EndTResume:
    Set objCmd = Nothing
    CloseDB
Exit Sub
EndTError:
    strResult = ""
    ShowErrorMessages Err, mstrModuleName, "StartTirage"
    Resume EndTResume
End Sub

Function GetRecordset(strSQL As String) As ADODB.Recordset
    'Return a disconnected recordset from the database.
    Dim rst As ADODB.Recordset
    On Error GoTo GRError
    
    If OpenDB() Then
        Set rst = New ADODB.Recordset
        rst.CursorLocation = adUseClient
        rst.Open strSQL, mobjConn, adOpenDynamic
        Set GetRecordset = rst
        rst.ActiveConnection = Nothing
    End If
GRResume:
    CloseDB
    Exit Function
GRError:
    ShowErrorMessages Err, mstrModuleName, "GetRecordset"
    Resume GRResume
End Function

Public Function GetRecordsetToArray(strSQL As String, ByRef recordCount As Long) As Variant
    'Return recordset data into an array. If, for any reason, the recordset
    ' does not return any data, this function simply creates a single element
    ' array and populates it with "No matching records...".
    Dim rst As ADODB.Recordset
    Dim arrData As Variant
    Dim lngX As Long
    
    recordCount = 0
    
    On Error GoTo GRAError
        
        If OpenDB() Then
            Set rst = New ADODB.Recordset
            With rst
                .CursorLocation = adUseClient
                .LockType = adLockReadOnly
                .Open strSQL, mobjConn, adOpenStatic
                If Not (rst.EOF) Then
                    'Disconnect the recordset
                    .ActiveConnection = Nothing
                    'Get the field count
                    lngX = .fields.Count
                    recordCount = CLng(.recordCount)
                    arrData = .GetRows()
                Else
                    'Recordset is empty; create dummy array record
                    ReDim arrData(0, 0)
                    arrData(0, 0) = "No matching records found in the database"
                End If
            End With
        Else
        
            'Connection failed for some reason; in order not to crash the
            ' calling routine, create dummy array record.
            ReDim arrData(0, 0)
            arrData(0, 0) = "Could not open database connection"
        End If
GRAResume:
    GetRecordsetToArray = arrData
    CloseDB
    Exit Function
GRAError:
    ShowErrorMessages Err, mstrModuleName, "GetRecordsetToArray"
    Resume GRAResume
End Function

Public Function GetMultipleRecordsets(strSQL As String) As ADODB.Recordset
    'Allows for multiple recordsets to be returned to the calling routine.
    'IMPORTANT - Although this method is similar to the GetRecordSet function,
    ' there are significant differences.
    ' 1. Do not disconnect the Recordset (i.e. don't set ActiveConnection to Nothing).
    ' 2. You must not close the database connection! It must remain open so that
    ' the calling routine can retrieve all of the resultsets in the returned Recordset.
    ' IT IS UP TO THE CALLING ROUTINE TO CLOSE THE CONNECTION WHEN DONE.
    Dim rst As ADODB.Recordset
    
    On Error GoTo GMRError
        If OpenDB() Then
            Set rst = New ADODB.Recordset
            rst.Open strSQL, mobjConn
            Set GetMultipleRecordsets = rst
        End If
GMRResume:
    'REMEMBER! Do NOT close the connection here. Calling routine must close it.
    Exit Function
GMRError:
    ShowErrorMessages Err, mstrModuleName, "GetMultipleRecordsets"
    Resume GMRResume
End Function

Public Function ExecuteActionQuery(SQLToExecute As String, Optional ByRef lastIdentify As Variant) As Boolean
    'Execute an action query, or stored procedure, which is not expected
    ' to return a resultset.
    Dim blnReturn As Boolean
    blnReturn = True
    On Error GoTo EAQError
    If OpenDB() Then
        mobjConn.Execute SQLToExecute
        If Not IsMissing(lastIdentify) Then
            lastIdentify = mobjConn.Execute("select @@identity").fields(0)
        End If
    End If
EAQResume:
    CloseDB
    ExecuteActionQuery = blnReturn
    Exit Function
EAQError:
    blnReturn = False
    ShowErrorMessages Err, mstrModuleName, "ExecuteActionQuery"
    Resume EAQResume
End Function

Public Property Get SQLConnection() As ADODB.Connection
    'Return the actual connection object, if requested by
    ' calling routine.
    If Not mobjConn Is Nothing Then
        On Error Resume Next
        Set SQLConnection = mobjConn
        On Error GoTo 0
    End If
End Property










Public Function TruncateTable(sName) As Boolean
    TruncateTable = ExecuteActionQuery("TRUNCATE TABLE " & sName & "")
End Function

Public Sub LoadCsvIntoTable_count()
    Dim rstParticipant As ADODB.Recordset
    Dim rstUrne As ADODB.Recordset
    Dim rstUrneLib As ADODB.Recordset
    Dim rstUrneMoo As ADODB.Recordset
    Dim i As Long, j As Long, K As Long
    Dim Rw As Long
    
    Dim strFile As String, str1 As String
    Dim strNow As String, dtNow As String
    Dim linesCount As Long
    Dim numsImported As Long
    
    Dim pctdone As Single
    
    Dim t As Double
    t = Time
    
    strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Selectionnez la base (.csv) de numéros...")
    If strFile = "Faux" Then
        Exit Sub
    Else
        
        linesCount = FileRowCount(strFile)
        
        Dim currNum As String, currMt As String
        dtNow = Date
            
        Do While i <= linesCount
                    
            pctdone = i / linesCount
            With ChargingDataBaseUFrm
                .LabelCaption.Caption = "Traitement en cours... " & i & " / " & linesCount
                .LabelProgress.Width = pctdone * (.FrameProgress.Width)
                .lblElapsedTime.Caption = "Temps traitement " & Format((Time - t), "hh:mm:ss")
            End With
                    
            DoEvents
                    
            i = i + 1
            numsImported = i
        Loop
        MsgBox CStr(numsImported) & " numéros chargés avec succès.", vbInformation, GetAppName
    End If
End Sub

Public Sub LoadCsvIntoTable()
    Dim rstParticipant As ADODB.Recordset
    Dim rstUrne As ADODB.Recordset
    Dim rstUrneLib As ADODB.Recordset
    Dim rstUrneMoo As ADODB.Recordset
    Dim i As Long, j As Long, K As Long
    Dim Rw As Long
    
    Dim strFile As String, str1 As String
    Dim strNow As String, dtNow As String
    Dim linesCount As Long
    Dim numsImported As Long
    
    Dim rqst As String
    
    Dim pctdone As Single
    
    Dim t As Double
    t = Time
    
    strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Selectionnez la base (.csv) de numéros...")
    If strFile = "Faux" Then
        Exit Sub
    Else
        
        linesCount = FileRowCount(strFile)
        
        If OpenDB() Then
            Dim currNum As String, currMt As String, currTypeconso As String
            dtNow = Date
            
            Open strFile For Input As #1
                i = 0
                numsImported = 0
                Do While Not EOF(1)
                    Line Input #1, str1
                    
                    pctdone = i / linesCount
                    With ChargingDataBaseUFrm
                        .LabelCaption.Caption = "Traitement en cours... " & i & " / " & linesCount
                        .LabelProgress.Width = pctdone * (.FrameProgress.Width)
                        .lblElapsedTime.Caption = "Temps traitement " & Format((Time - t), "hh:mm:ss")
                    End With
                    
                    DoEvents
                    
                    currNum = Split(str1, strColSep)(0)
                    currTypeconso = "ALL" 'Split(str1, strColSep)(1)
                    currMt = Split(str1, strColSep)(1)
                    
                    If Left(currNum, 3) = "241" Then
                        currNum = Replace(currNum, "241", "", 1, 3)
                    End If
                    
                    currNum = Right(currNum, 8)
                    
                    If (Left(currNum, 2) = "05" Or Left(currNum, 2) = "02" Or Left(currNum, 2) = "06") And (Len(currNum) = 8) Then
                        rqst = "INSERT INTO participant(msisdn,montant,creation_date,typeconso) VALUES('" & currNum & "'," & currMt & ",'" & CStr(dtNow) & "','" & currTypeconso & "');"
                        'MsgBox rqst
                        mobjConn.Execute rqst
                        numsImported = numsImported + 1
                    End If
                    i = i + 1
                Loop
            Close #1
            
            ' Clean up
        
            CloseDB
            MsgBox CStr(numsImported) & " numéros chargés avec succès.", vbInformation, GetAppName
        End If
    End If
End Sub


Public Function CSVFileLinesCount(ByVal strPathtoTextFile As String) As Long
    Dim CSV_FILE As String
    Dim arrPathtoTextFile() As String
     'set reference to Microsoft ActiveX Data Objects Library (Tools>References...)
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    
    If strPathtoTextFile = "Faux" Then
        CSVFileLinesCount = -1
    Else
        arrPathtoTextFile = Split(strPathtoTextFile, "\")
        CSV_FILE = arrPathtoTextFile(UBound(arrPathtoTextFile))
        
        strPathtoTextFile = Replace(strPathtoTextFile, CSV_FILE, "")
        
        Set cnn = New ADODB.Connection
        Set rst = New ADODB.Recordset
        
        cnn.Open "Provider=" & DbProvider & ";" & _
        "Data Source=" & strPathtoTextFile & ";" & _
        "Extended Properties=""text;HDR=No;FMT=CSVDelimited"""
        
        rst.Open "SELECT * FROM " & CSV_FILE, cnn, adOpenStatic, adLockOptimistic, adCmdText
        
        If Not rst.EOF Then CSVFileLinesCount = CLng(rst.recordCount)
         
        rst.Close 'Close ADO objects
        cnn.Close
         
        Set rst = Nothing
        Set cnn = Nothing
    End If
End Function


Public Function FileRowCount(filePath As String, Optional RowDelimiter As String = vbCr) As Long
' Returns the row count of a text file, including the header row
' Returns - 1 on error

' Unicode-compliant, works on UTF-8, UTF-16, ASCII, with or without a Byte order Marker.
' Reads a typical 30Mb file over the network in 200-300ms. Hint: always copy to a local folder.

' If you're scanning files for use with a SQL driver, use basSQL.TableRowCount: it's 20x slower,
' but it returns a proper test of the file's usability as a SQL 'table'

' Nigel Heffernan    Excellerando.Blogspot.com  2015

' Unit test:
' s=Timer : for i = 0 to 99 : n=FileRowCount("C:\Temp\MyFile.csv") : Next i : Print Format(n,"#,##0") & " rows in " & FORMAT((Timer-s)/i,"0.000") & " sec"

' Network performance on a good day: reads ~ 150 MB/second, plus an overhead of 70 ms for each file
' Local-drive performance:                 ~ 4.5 GB/second, plus an overhead of  4 ms for each file

    On Error GoTo FRCError
    
    Dim fileRowCountResutl As Long
    Dim hndFile     As Long
    Dim lngRowCount As Long
    Dim lngOffset   As Long
    Dim lngFileLen  As Long

    Const CHUNK_SIZE As Long = 8192

    Dim strChunk As String * CHUNK_SIZE
    
    Dim arrBytes(CHUNK_SIZE) As Byte
    
    fileRowCountResutl = 0
    
    If Len(Dir(filePath, vbNormal)) < 1 Then
        fileRowCountResutl = -1
        FileRowCount = -1
        Exit Function
    End If

    ' trap the error of a folder path without a filename:
    If FileName(filePath) = "" Then
        fileRowCountResutl = -1
        FileRowCount = -1
        Exit Function
    End If

 
    hndFile = FreeFile
    Open filePath For Binary Access Read Shared As #hndFile


    lngFileLen = LOF(hndFile)


    lngOffset = 1
    Do Until EOF(hndFile)
        Get #hndFile, , strChunk
        fileRowCountResutl = fileRowCountResutl + UBound(Split(strChunk, RowDelimiter))
    Loop
    
    Close #hndFile
    
    Erase arrBytes

FRCResume:
    FileRowCount = fileRowCountResutl
    Exit Function
FRCError:
    Dim str As String
    str = InputBox("Erreur. Entrer nombre de lignes du fichier:", "Base Participant")
    If str <> "" Then
        fileRowCountResutl = CLng(str)
    End If
    Resume FRCResume
End Function



Public Function FileName(path As String) As String
' Strip the folder and path from a file's path string, leaving only the file name

' This does not check for the existence or accessibility of the file:
' all we're doing here is string-handling

' Nigel Heffernan    Excellerando.Blogspot.com  2011

    Dim strPath   As String
    Dim arrPath() As String

    Const BACKSLASH As String * 1 = "\"

    strPath = Trim(path)

    If strPath = "" Then Exit Function
    If Right$(strPath, 1) = BACKSLASH Then Exit Function

    arrPath = Split(strPath, BACKSLASH)

    If UBound(arrPath) = 0 Then          ' does not contain "\"
        FileName = path
    Else
        FileName = arrPath(UBound(arrPath))
    End If

    Erase arrPath

End Function








Public Sub LoadCsvIntoTable_new()
    Dim rstParticipant As ADODB.Recordset
    Dim rstUrne As ADODB.Recordset
    Dim rstUrneLib As ADODB.Recordset
    Dim rstUrneMoo As ADODB.Recordset
    Dim i As Long, j As Long, K As Long
    Dim Rw As Long
    
    Dim strFile As String, str1 As String
    Dim strNow As String, dtNow As String
    Dim linesCount As Long
    Dim numsImported As Long
    
    Dim pctdone As Single
    
    Dim t As Double
    t = Time
    
    strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Selectionnez la base (.csv) de numéros...")
    If strFile = "Faux" Then
        Exit Sub
    Else
        
        linesCount = CSVFileLinesCount(strFile)
        
        If OpenDB() Then
        
            
            
            ' // send Errors to the Screen
            Spinner.EchoErrors = True

            ' // exit if the Spinner is already Running
            '    useful to capture repeated Button Clicks
            If Spinner.Running Then Exit Sub

            ' // fade the Spinner in
            ' // [Spinner__Number] fails in addins
            Spinner.FadeIn PixelBuddhaSpinner:=3, _
                   Duration:=3000, _
                   Disable:=CTRLBreak, _
                   Position:=ApplicationCenter, _
                   WaitForDuration:=3000

            ' // complete a dummy task
            Dim lngTask As Long
            
            For lngTask = 1 To linesCount
            
                Dim currNum As String, currMt As String
                dtNow = Date
            
                Open strFile For Input As #1
                    i = 0
                    numsImported = 0
                    Do While Not EOF(1)
                        Line Input #1, str1
            
                        ' // use the alternative DoEvents Function which is faster
                        Spinner.MyDoEvents

                        ' // or the regular DoEvents
                        'Spinner.Indicator i / linesCount
                        'DoEvents
         
                        ' // output some Text & Values
                        '[Running__Task].Value = "processing, please wait... " & Format(lngTask / 20000, "0%")
                        Spinner.Indicator "processing, please wait... " & Format(lngTask / linesCount, "0%")
                        
                        currNum = Split(str1, ",")(0)
                        currMt = Split(str1, ",")(1)
                    
                        If Left(currNum, 3) = "241" Then
                            currNum = Replace(currNum, "241", "", 1, 3)
                        End If
                    
                        If (Left(currNum, 2) = "05" Or Left(currNum, 2) = "02" Or Left(currNum, 2) = "06") And (Len(currNum) = 8) Then
                            mobjConn.Execute "INSERT INTO participant(msisdn,montant,creation_date) VALUES('" & currNum & "'," & currMt & ",'" & CStr(dtNow) & "')"
                            numsImported = numsImported + 1
                        End If
                        i = i + 1
                    Loop
                Close #1

                Next lngTask
    
                ' // output some Text
                '[Running__Task].Value = "complete, fading out..."
    
                ' // fade the Spinner out if it is Running
                If Spinner.Running Then Spinner.FadeOut Duration:=3000
    
                ' // output some Text
                '[Running__Task].Value = "task complete"
                Spinner.Indicator "task complete"
            
            
        
            
                    
                    
            
                ' Clean up
        
                CloseDB
                MsgBox CStr(numsImported) & " numéros chargés avec succès.", vbInformation, GetAppName
        End If
    End If
End Sub
