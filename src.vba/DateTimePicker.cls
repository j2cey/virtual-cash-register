Option Explicit
Private Type tagINITCOMMONCONTROLSEX
    dwSize     As Long
    dwICC      As Long
End Type
Private Type SYSTEMTIME
    Year       As Integer
    Month      As Integer
    DayOfWeek  As Integer
    Day        As Integer
    Hour       As Integer
    Minute     As Integer
    Second     As Integer
    Milliseconds As Integer
End Type
Private Type SYSTEMTIMERANGE
    MinYear    As Integer
    MinMonth   As Integer
    MinDayOfWeek As Integer
    MinDay     As Integer
    MinHour    As Integer
    MinMinute  As Integer
    MinSecond  As Integer
    MinMilliseconds As Integer
    MaxYear    As Integer
    MaxMonth   As Integer
    MaxDayOfWeek As Integer
    MaxDay     As Integer
    MaxHour    As Integer
    MaxMinute  As Integer
    MaxSecond  As Integer
    MaxMilliseconds As Integer
End Type

Private Const DATETIMEPICK_CLASS = "SysDateTimePick32"
Private Const ICC_DATE_CLASSES = &H100
Private Const DTS_SHORTDATEFORMAT = &H0
Private Const DTS_LONGDATEFORMAT = &H4
Private Const GDT_VALID = 0
Private Const GDTR_MIN = &H1
Private Const GDTR_MAX = &H2
Private Const WM_SETFONT = &H30
Private Const WM_GETFONT = &H31
Private Const DTM_FIRST = &H1000
Private Const DTM_GETSYSTEMTIME = (DTM_FIRST + 1)
Private Const DTM_SETSYSTEMTIME = (DTM_FIRST + 2)
Private Const DTM_GETRANGE = (DTM_FIRST + 3)
Private Const DTM_SETRANGE = (DTM_FIRST + 4)
Private Const DTM_SETFORMAT = (DTM_FIRST + 5)
Private Const DTM_SETMCCOLOR = (DTM_FIRST + 6)
Private Const DTM_GETMCCOLOR = (DTM_FIRST + 7)
Private Const DTM_GETMONTHCAL = (DTM_FIRST + 8)
Private Const DTM_SETMCFONT = (DTM_FIRST + 9)
Private Const DTM_GETMCFONT = (DTM_FIRST + 10)
Private Const MCSC_BACKGROUND = 0
Private Const MCSC_TEXT = 1
Private Const MCSC_TITLEBK = 2
Private Const MCSC_TITLETEXT = 3
Private Const MCSC_MONTHBK = 4
Private Const MCSC_TRAILINGTEXT = 5
Private Const POINT_PER_INCH = 72
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const GWL_HINSTANCE As Long = (-6)
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const FW_DONTCARE = 0
Private Const FW_BOLD = 100 ' 700
Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SHIFTJIS_CHARSET = 128
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_DONTCARE = 0
Private Const DEFAULT_GUI_FONT = 20

#If Win64 Then
    Private Declare PtrSafe Function InitCommonControlsEx Lib "ComCtl32" (LPINITCOMMONCONTROLSEX As Any) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare PtrSafe Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
    Private Declare PtrSafe Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal nWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal pszFaceName As String) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function InitCommonControlsEx Lib "ComCtl32" (LPINITCOMMONCONTROLSEX As Any) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
    Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
    Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal nWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal pszFaceName As String) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
#End If

Private lngPixelsX As Long
Private lngPixelsY As Long
Private lnghWnd_Excel As Long
Private lnghInstance As Long
Private strThunder As String
Private blnCreate As Boolean
Private MyForm As MSForms.UserForm
Private lnghWnd_Form As Long
Private lnghWnd_Sub As Long
Private colBaseCtrl As Collection
Private lnghwndDTP() As Long
Private lnghwndFONT() As Long
Private lnghwndDefaultFONT As Long
Private Const cstMask As Long = &H7FFFFFFF


Private Sub Class_Initialize()
    blnCreate = False
    Set colBaseCtrl = New Collection
    Call GetLogPixelsXY
    If (Val(Application.Version) <= 8) Then
        strThunder = "ThunderXFrame"
    Else
        strThunder = "ThunderDFrame"
    End If
    lnghwndDefaultFONT = GetStockObject(DEFAULT_GUI_FONT)
End Sub

Private Sub Class_Terminate()
    Me.Destroy
End Sub

Public Sub Add(NewCtrl As MSForms.ComboBox)
    If (blnCreate = False) Then
        colBaseCtrl.Add NewCtrl
    Else
    End If
End Sub

Public Sub Create(ByVal NewForm As MSForms.UserForm, Optional ByVal NewFormat As String = "", Optional ByVal NewMinDate As Date = #1/1/1601#, Optional ByVal NewMaxDate As Date = #12/31/9999#, Optional ByVal ForeColor As Long = -1, Optional ByVal BackColor As Long = -1, Optional ByVal TitleFore As Long = -1, Optional ByVal TitleBack As Long = -1, Optional ByVal Trailing As Long = -1, Optional ByVal DefaultFONT As Variant = True)
    Dim i      As Integer
    Dim intFONT As Integer
    If (VarType(DefaultFONT) = vbBoolean) Then
        If (DefaultFONT = True) Then
            intFONT = 1
        Else
            intFONT = 0
        End If
    Else
        Select Case DefaultFONT
            Case 0 To 3
                intFONT = DefaultFONT
            Case Else
                intFONT = 1
        End Select
    End If
    If (blnCreate = True) Then
    ElseIf (colBaseCtrl.count < 1) Or (NewForm Is Nothing) Then
        blnCreate = False
    Else
        Set MyForm = NewForm
        lnghWnd_Form = FindWindow(strThunder, MyForm.Caption)
        If (lnghWnd_Form = 0) Then
            blnCreate = False
            Exit Sub
        End If
        lnghWnd_Excel = FindWindow("XLMAIN", Application.Caption)
        lnghInstance = GetWindowLong(lnghWnd_Excel, GWL_HINSTANCE)
        lnghWnd_Sub = FindWindowEx(lnghWnd_Form, 0&, vbNullString, vbNullString)
        ReDim lnghwndDTP(1 To colBaseCtrl.count)
        ReDim lnghwndFONT(1 To colBaseCtrl.count)
        For i = 1 To colBaseCtrl.count
            Call Create_Sub(i, NewFormat, NewMinDate, NewMaxDate, intFONT)
        Next i
        blnCreate = True
        If (NewMinDate <= NewMaxDate) Then
            Me.MinDate(0) = NewMinDate
            Me.MaxDate(0) = NewMaxDate
        Else
            Me.MinDate(0) = #1/1/1601#
            Me.MaxDate(0) = #12/31/9999#
        End If
        Me.CalendarForeColor(0) = ForeColor
        Me.CalendarBackColor(0) = BackColor
        Me.CalendarTitleForeColor(0) = TitleFore
        Me.CalendarTitleBackColor(0) = TitleBack
        Me.CalendarTrailingForeColor(0) = Trailing
    End If
End Sub

Public Sub Destroy()
    Dim i      As Integer
    If (blnCreate = False) Then
    Else
        If (colBaseCtrl.count > 0) Then
            For i = 1 To colBaseCtrl.count
                If (IsWindow(lnghwndDTP(i)) <> 0) Then
                    Call DestroyWindow(lnghwndDTP(i))
                End If
                If (lnghwndFONT(i) <> 0) Then
                    Call DeleteObject(lnghwndFONT(i))
                End If
            Next i
        End If
    End If
    Set MyForm = Nothing
    Set colBaseCtrl = Nothing
    Erase lnghwndDTP
    Erase lnghwndFONT
    blnCreate = False
End Sub

Private Sub Create_Sub(ByVal Index As Integer, ByVal NewFormat As String, ByVal NewMinDate As Date, ByVal NewMaxDate As Date, ByVal NewFONT As Integer)
    Dim icce   As tagINITCOMMONCONTROLSEX
    Dim lngResult As Long
    Dim dblTop As Double
    Dim dblLeft As Double
    Dim dblHeight As Double
    Dim dblWidth As Double
    Dim dblFontHeight As Double
    Dim intWeight As Integer
    
    dblLeft = colBaseCtrl(Index).Left * lngPixelsX / POINT_PER_INCH
    dblTop = colBaseCtrl(Index).Top * lngPixelsY / POINT_PER_INCH
    dblWidth = colBaseCtrl(Index).Width * lngPixelsX / POINT_PER_INCH
    dblHeight = colBaseCtrl(Index).Height * lngPixelsY / POINT_PER_INCH
    dblFontHeight = colBaseCtrl(Index).Font.Size * lngPixelsY / POINT_PER_INCH
    colBaseCtrl(Index).Enabled = False
    If (IsWindow(lnghwndDTP(Index)) <> 0) Then
        Call DestroyWindow(lnghwndDTP(Index))
    End If
    With icce
        .dwICC = ICC_DATE_CLASSES
        .dwSize = Len(icce)
    End With
    lngResult = InitCommonControlsEx(icce)
    lnghwndDTP(Index) = CreateWindowEX(0&, DATETIMEPICK_CLASS, vbNullString, WS_CHILD Or WS_VISIBLE Or DTS_SHORTDATEFORMAT, dblLeft, dblTop, dblWidth, dblHeight, lnghWnd_Sub, 0&, lnghInstance, vbNullString)
    If (NewFONT = 1) Then
    Else
        If (colBaseCtrl(Index).Font.Bold = True) Then
            intWeight = FW_BOLD
        Else
            intWeight = FW_DONTCARE
        End If
        lnghwndFONT(Index) = CreateFont(CLng(dblFontHeight * (-1)), 0, 0, 0, intWeight, 0, 0, 0, SHIFTJIS_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, (DEFAULT_PITCH + FF_DONTCARE), colBaseCtrl(Index).Font.Name)
        Select Case NewFONT
            Case 0
                lngResult = SendMessage(lnghwndDTP(Index), WM_SETFONT, lnghwndFONT(Index), ByVal 0&)
            Case 2
                lngResult = SendMessage(lnghwndDTP(Index), WM_SETFONT, lnghwndFONT(Index), ByVal 0&)
                lngResult = SendMessage(lnghwndDTP(Index), DTM_SETMCFONT, lnghwndDefaultFONT, ByVal 0&)
            Case 3
                lngResult = SendMessage(lnghwndDTP(Index), DTM_SETMCFONT, lnghwndFONT(Index), ByVal 0&)
        End Select
    End If
End Sub

Private Sub GetLogPixelsXY()
    Dim lngHWnd As Long
    Dim lngDC  As Long
    lngHWnd = GetDesktopWindow()
    lngDC = GetDC(lngHWnd)
    lngPixelsX = GetDeviceCaps(lngDC, LOGPIXELSX)
    lngPixelsY = GetDeviceCaps(lngDC, LOGPIXELSY)
    ReleaseDC lngHWnd, lngDC
End Sub

Public Property Get count() As Integer
    count = colBaseCtrl.count
End Property

Public Property Let DateFormat(Optional ByVal Index As Integer = 0, ByVal NewFormat As String)
    Dim lngResult As Long
    Dim i      As Integer
    If (blnCreate = True) And (NewFormat <> "") Then
        Select Case Index
            Case 0
                For i = 1 To colBaseCtrl.count
                    lngResult = SendMessage(lnghwndDTP(i), DTM_SETFORMAT, GDT_VALID, ByVal NewFormat)
                Next i
            Case 1 To colBaseCtrl.count
                lngResult = SendMessage(lnghwndDTP(Index), DTM_SETFORMAT, GDT_VALID, ByVal NewFormat)
            Case Else
        End Select
    Else
    End If
End Property

Public Property Get Value(Optional ByVal Index As Integer = 0) As Date
    Dim st     As SYSTEMTIME
    Dim lngResult As Long
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                lngResult = SendMessage(lnghwndDTP(1), DTM_GETSYSTEMTIME, 0, st)
                Value = DateSerial(st.Year, st.Month, st.Day)
            Case 1 To colBaseCtrl.count
                lngResult = SendMessage(lnghwndDTP(Index), DTM_GETSYSTEMTIME, 0, st)
                Value = DateSerial(st.Year, st.Month, st.Day)
            Case Else
                Value = CDate(0)
        End Select
    Else
        Value = CDate(0)
    End If
End Property

Public Property Let Value(Optional ByVal Index As Integer = 0, ByVal newValue As Date)
    Dim lngResult As Long
    Dim st     As SYSTEMTIME
    Dim i      As Integer
    If (blnCreate = True) Then
        With st
            .Year = Year(newValue)
            .Month = Month(newValue)
            .Day = Day(newValue)
            .DayOfWeek = Weekday(newValue) - 1
            .Hour = 0
            .Minute = 0
            .Second = 0
        End With
        Select Case Index
            Case 0
                For i = 1 To colBaseCtrl.count
                    lngResult = SendMessage(lnghwndDTP(i), DTM_SETSYSTEMTIME, GDT_VALID, st)
                Next i
            Case 1 To colBaseCtrl.count
                lngResult = SendMessage(lnghwndDTP(Index), DTM_SETSYSTEMTIME, GDT_VALID, st)
            Case Else
        End Select
    Else
    End If
End Property

Public Property Get Enabled(Optional ByVal Index As Integer = 0) As Boolean
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                Enabled = CBool(IsWindowEnabled(lnghwndDTP(1)))
            Case 1 To colBaseCtrl.count
                Enabled = CBool(IsWindowEnabled(lnghwndDTP(Index)))
            Case Else
                Enabled = False
        End Select
    Else
        Enabled = False
    End If
End Property

Public Property Let Enabled(Optional ByVal Index As Integer = 0, ByVal blnNewEnabled As Boolean)
    Dim lngResult As Long
    Dim i      As Integer
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                For i = 1 To colBaseCtrl.count
                    lngResult = EnableWindow(lnghwndDTP(i), blnNewEnabled)
                Next i
            Case 1 To colBaseCtrl.count
                lngResult = EnableWindow(lnghwndDTP(Index), blnNewEnabled)
            Case Else
        End Select
    Else
    End If
End Property

Public Property Get MinDate(Optional ByVal Index As Integer = 0) As Date
    Dim lngResult As Long
    Dim st     As SYSTEMTIMERANGE
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                lngResult = SendMessage(lnghwndDTP(1), DTM_GETRANGE, GDT_VALID, st)
                MinDate = DateSerial(st.MinYear, st.MinMonth, st.MinDay)
            Case 1 To colBaseCtrl.count
                lngResult = SendMessage(lnghwndDTP(Index), DTM_GETRANGE, GDT_VALID, st)
                MinDate = DateSerial(st.MinYear, st.MinMonth, st.MinDay)
            Case Else
                MinDate = CDate(0)
        End Select
    Else
        MinDate = CDate(0)
    End If
End Property

Public Property Let MinDate(Optional ByVal Index As Integer = 0, ByVal dtmNewValue As Date)
    Dim i      As Integer
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                For i = 1 To colBaseCtrl.count
                    Call MinDate_Sub(i, dtmNewValue)
                Next i
            Case 1 To colBaseCtrl.count
                Call MinDate_Sub(Index, dtmNewValue)
            Case Else
        End Select
    Else
    End If
End Property

Private Sub MinDate_Sub(ByVal Index As Integer, ByVal dtmNewValue As Date)
    Dim lngResult As Long
    Dim dtmMax As Date
    Dim st     As SYSTEMTIMERANGE
    lngResult = SendMessage(lnghwndDTP(Index), DTM_GETRANGE, GDT_VALID, st)
    With st
        .MinYear = Year(dtmNewValue)
        .MinMonth = Month(dtmNewValue)
        .MinDay = Day(dtmNewValue)
        .MinDayOfWeek = Weekday(dtmNewValue) - 1
        .MinHour = 0
        .MinMinute = 0
        .MinSecond = 0
    End With
    If (st.MaxYear = 0) Then
        lngResult = SendMessage(lnghwndDTP(Index), DTM_SETRANGE, GDTR_MIN, st)
    Else
        dtmMax = DateSerial(st.MaxYear, st.MaxMonth, st.MaxDay)
        If (dtmMax < dtmNewValue) Then
            With st
                .MaxYear = .MinYear
                .MaxMonth = .MinMonth
                .MaxDay = .MinDay
                .MaxDayOfWeek = .MinDayOfWeek
                .MaxHour = 0
                .MaxMinute = 0
                .MaxSecond = 0
            End With
        End If
        lngResult = SendMessage(lnghwndDTP(Index), DTM_SETRANGE, GDTR_MIN Or GDTR_MAX, st)
    End If
End Sub

Public Property Get MaxDate(Optional ByVal Index As Integer = 0) As Date
    Dim lngResult As Long
    Dim st     As SYSTEMTIMERANGE
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                lngResult = SendMessage(lnghwndDTP(1), DTM_GETRANGE, GDT_VALID, st)
                MaxDate = DateSerial(st.MaxYear, st.MaxMonth, st.MaxDay)
            Case 1 To colBaseCtrl.count
                lngResult = SendMessage(lnghwndDTP(Index), DTM_GETRANGE, GDT_VALID, st)
                MaxDate = DateSerial(st.MaxYear, st.MaxMonth, st.MaxDay)
            Case Else
                MaxDate = CDate(0)
        End Select
    Else
        MaxDate = CDate(0)
    End If
End Property

Public Property Let MaxDate(Optional ByVal Index As Integer = 0, ByVal dtmNewValue As Date)
    Dim i      As Integer
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                For i = 1 To colBaseCtrl.count
                    Call MaxDate_Sub(i, dtmNewValue)
                Next i
            Case 1 To colBaseCtrl.count
                Call MaxDate_Sub(Index, dtmNewValue)
            Case Else
        End Select
    Else
    End If
End Property

Private Sub MaxDate_Sub(ByVal Index As Integer, ByVal dtmNewValue As Date)
    Dim lngResult As Long
    Dim dtmMin As Date
    Dim st     As SYSTEMTIMERANGE
    lngResult = SendMessage(lnghwndDTP(Index), DTM_GETRANGE, GDT_VALID, st)
    With st
        .MaxYear = Year(dtmNewValue)
        .MaxMonth = Month(dtmNewValue)
        .MaxDay = Day(dtmNewValue)
        .MaxDayOfWeek = Weekday(dtmNewValue) - 1
        .MaxHour = 0
        .MaxMinute = 0
        .MaxSecond = 0
    End With
    If (st.MinYear = 0) Then
        lngResult = SendMessage(lnghwndDTP(Index), DTM_SETRANGE, GDTR_MAX, st)
    Else
        dtmMin = DateSerial(st.MinYear, st.MinMonth, st.MinDay)
        If (dtmMin > dtmNewValue) Then
            With st
                .MinYear = .MaxYear
                .MinMonth = .MaxMonth
                .MinDay = .MaxDay
                .MinDayOfWeek = .MaxDayOfWeek
                .MinHour = 0
                .MinMinute = 0
                .MinSecond = 0
            End With
        End If
        lngResult = SendMessage(lnghwndDTP(Index), DTM_SETRANGE, GDTR_MAX Or GDTR_MIN, st)
    End If
End Sub

Public Property Get CalendarBackColor(Optional ByVal Index As Integer = 0) As Long
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                CalendarBackColor = SendMessage(lnghwndDTP(1), DTM_GETMCCOLOR, MCSC_MONTHBK, ByVal 0&)
            Case 1 To colBaseCtrl.count
                CalendarBackColor = SendMessage(lnghwndDTP(Index), DTM_GETMCCOLOR, MCSC_MONTHBK, ByVal 0&)
            Case Else
                CalendarBackColor = -1
        End Select
    Else
        CalendarBackColor = -1
    End If
End Property

Public Property Let CalendarBackColor(Optional ByVal Index As Integer = 0, ByVal NewColor As Long)
    Dim lngResult As Long
    Dim lngRGB As Long
    Dim i      As Integer
    If (blnCreate = True) Then
        Select Case NewColor
            Case vbScrollBars To vbInfoBackground
                lngRGB = GetSysColor(NewColor And cstMask)
            Case &H0& To &HFFFFFF
                lngRGB = NewColor
            Case Else
                Exit Property
        End Select
        Select Case Index
            Case 0
                For i = 1 To colBaseCtrl.count
                    lngResult = SendMessage(lnghwndDTP(i), DTM_SETMCCOLOR, MCSC_MONTHBK, ByVal lngRGB)
                Next i
            Case 1 To colBaseCtrl.count
                lngResult = SendMessage(lnghwndDTP(Index), DTM_SETMCCOLOR, MCSC_MONTHBK, ByVal lngRGB)
            Case Else
        End Select
    Else
    End If
End Property

Public Property Get CalendarForeColor(Optional ByVal Index As Integer = 0) As Long
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                CalendarForeColor = SendMessage(lnghwndDTP(1), DTM_GETMCCOLOR, MCSC_TEXT, ByVal 0&)
            Case 1 To colBaseCtrl.count
                CalendarForeColor = SendMessage(lnghwndDTP(Index), DTM_GETMCCOLOR, MCSC_TEXT, ByVal 0&)
            Case Else
                CalendarForeColor = -1
        End Select
    Else
        CalendarForeColor = -1
    End If
End Property

Public Property Let CalendarForeColor(Optional ByVal Index As Integer = 0, ByVal NewColor As Long)
    Dim lngResult As Long
    Dim lngRGB As Long
    Dim i      As Integer
    If (blnCreate = True) Then
        Select Case NewColor
            Case vbScrollBars To vbInfoBackground
                lngRGB = GetSysColor(NewColor And cstMask)
            Case &H0& To &HFFFFFF
                lngRGB = NewColor
            Case Else
                Exit Property
        End Select
        Select Case Index
            Case 0
                For i = 1 To colBaseCtrl.count
                    lngResult = SendMessage(lnghwndDTP(i), DTM_SETMCCOLOR, MCSC_TEXT, ByVal lngRGB)
                Next i
            Case 1 To colBaseCtrl.count
                lngResult = SendMessage(lnghwndDTP(Index), DTM_SETMCCOLOR, MCSC_TEXT, ByVal lngRGB)
            Case Else
        End Select
    Else
    End If
End Property

Public Property Get CalendarTrailingForeColor(Optional ByVal Index As Integer = 0) As Long
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                CalendarTrailingForeColor = SendMessage(lnghwndDTP(1), DTM_GETMCCOLOR, MCSC_TRAILINGTEXT, ByVal 0&)
            Case 1 To colBaseCtrl.count
                CalendarTrailingForeColor = SendMessage(lnghwndDTP(Index), DTM_GETMCCOLOR, MCSC_TRAILINGTEXT, ByVal 0&)
            Case Else
                CalendarTrailingForeColor = -1
        End Select
    Else
        CalendarTrailingForeColor = -1
    End If
End Property

Public Property Let CalendarTrailingForeColor(Optional ByVal Index As Integer = 0, ByVal NewColor As Long)
    Dim lngResult As Long
    Dim lngRGB As Long
    Dim i      As Integer
    If (blnCreate = True) Then
        Select Case NewColor
            Case vbScrollBars To vbInfoBackground
                lngRGB = GetSysColor(NewColor And cstMask)
            Case &H0& To &HFFFFFF
                lngRGB = NewColor
            Case Else
                Exit Property
        End Select
        Select Case Index
            Case 0
                For i = 1 To colBaseCtrl.count
                    lngResult = SendMessage(lnghwndDTP(i), DTM_SETMCCOLOR, MCSC_TRAILINGTEXT, ByVal lngRGB)
                Next i
            Case 1 To colBaseCtrl.count
                lngResult = SendMessage(lnghwndDTP(Index), DTM_SETMCCOLOR, MCSC_TRAILINGTEXT, ByVal lngRGB)
            Case Else
        End Select
    Else
    End If
End Property

Public Property Get CalendarTitleBackColor(Optional ByVal Index As Integer = 0) As Long
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                CalendarTitleBackColor = SendMessage(lnghwndDTP(1), DTM_GETMCCOLOR, MCSC_TITLEBK, ByVal 0&)
            Case 1 To colBaseCtrl.count
                CalendarTitleBackColor = SendMessage(lnghwndDTP(Index), DTM_GETMCCOLOR, MCSC_TITLEBK, ByVal 0&)
            Case Else
                CalendarTitleBackColor = -1
        End Select
    Else
        CalendarTitleBackColor = -1
    End If
End Property

Public Property Let CalendarTitleBackColor(Optional ByVal Index As Integer = 0, ByVal NewColor As Long)
    Dim lngResult As Long
    Dim lngRGB As Long
    Dim i      As Integer
    If (blnCreate = True) Then
        Select Case NewColor
            Case vbScrollBars To vbInfoBackground
                lngRGB = GetSysColor(NewColor And cstMask)
            Case &H0& To &HFFFFFF
                lngRGB = NewColor
            Case Else
                Exit Property
        End Select
        Select Case Index
            Case 0
                For i = 1 To colBaseCtrl.count
                    lngResult = SendMessage(lnghwndDTP(i), DTM_SETMCCOLOR, MCSC_TITLEBK, ByVal lngRGB)
                Next i
            Case 1 To colBaseCtrl.count
                lngResult = SendMessage(lnghwndDTP(Index), DTM_SETMCCOLOR, MCSC_TITLEBK, ByVal lngRGB)
            Case Else
        End Select
    Else
    End If
End Property

Public Property Get CalendarTitleForeColor(Optional ByVal Index As Integer = 0) As Long
    If (blnCreate = True) Then
        Select Case Index
            Case 0
                CalendarTitleForeColor = SendMessage(lnghwndDTP(1), DTM_GETMCCOLOR, MCSC_TITLETEXT, ByVal 0&)
            Case 1 To colBaseCtrl.count
                CalendarTitleForeColor = SendMessage(lnghwndDTP(Index), DTM_GETMCCOLOR, MCSC_TITLETEXT, ByVal 0&)
            Case Else
                CalendarTitleForeColor = -1
        End Select
    Else
        CalendarTitleForeColor = -1
    End If
End Property

Public Property Let CalendarTitleForeColor(Optional ByVal Index As Integer = 0, ByVal NewColor As Long)
    Dim lngResult As Long
    Dim lngRGB As Long
    Dim i      As Integer
    If (blnCreate = True) Then
        Select Case NewColor
            Case vbScrollBars To vbInfoBackground
                lngRGB = GetSysColor(NewColor And cstMask)
            Case &H0& To &HFFFFFF
                lngRGB = NewColor
            Case Else
                Exit Property
        End Select
        Select Case Index
            Case 0
                For i = 1 To colBaseCtrl.count
                    lngResult = SendMessage(lnghwndDTP(i), DTM_SETMCCOLOR, MCSC_TITLETEXT, ByVal lngRGB)
                Next i
            Case 1 To colBaseCtrl.count
                lngResult = SendMessage(lnghwndDTP(Index), DTM_SETMCCOLOR, MCSC_TITLETEXT, ByVal lngRGB)
            Case Else
        End Select
    Else
    End If
End Property
