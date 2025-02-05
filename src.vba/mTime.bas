'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : mTime
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/23
' Purpose   : Manage all Time related Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function    : GetUuid
'   Purpose     : Get a Unique Identifier according to current Date and Time.
'   Arguments   :
'
'   Returns     : String
'
'   Date        Developer           Action
'   ---------------------------------------------------------------------------------------
'   2025/01/23      Jude Parfait        Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (Id As Any) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long
#End If

' Retrieves the current system date and time in Coordinated Universal Time (UTC) format.
' To retrieve the current system date and time in local time, use the GetLocalTime function.
'Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SystemTime)
'Private Declare Sub GetSystemTime Lib “kernel32” (lpSystemTime As SystemTime)
'Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SystemTime)

Public Function Now_System() As Double
    Dim st As SYSTEMTIME
    GetSystemTime st
    Now_System = DateSerial(st.wYear, st.wMonth, st.wDay) + _
        TimeSerial(st.wHour, st.wMinute, st.wSecond) + _
        st.wMilliseconds
        'st.wMilliseconds / 86400000#
End Function

Function GetTodayMilliseconds() As Long
    Dim SysTime As SYSTEMTIME
    
    GetSystemTime SysTime
    
    GetTodayMilliseconds = Hour(Now) * 3600000 + Minute(Now) * 60000 + Second(Now) * 1000 + SysTime.wMilliseconds
End Function

Function Now_Timer() As Double
    Now_Timer = CDbl(Int(Now)) + CDbl(Timer() / 86400#)
End Function


Sub CompareCurrentTimeFunctions()
    ' Compare precision of different methods to get current time.
    Me.Range("A1:D1000").NumberFormat = "yyyy/mm/dd h:mm:ss.000"

    Dim d As Double
    Dim i As Long
    For i = 2 To 1000
        ' 1) Excel NOW() formula returns same value until delay of ~10 milliseconds. (local time)
        Me.Cells(1, 1).Formula = "=Now()"
        d = Me.Cells(1, 1)
        Me.Cells(i, 1) = d

        ' 2) VBA Now() returns same value until delay of ~1 second. (local time)
        d = Now
        Me.Cells(i, 2) = d

        ' 3) VBA Timer returns same value until delay of ~5 milliseconds. (local time)
        Me.Cells(i, 3) = Now_Timer
        
        ' 4) System time is precise down to 1 millisecond. (UTC)
        Me.Cells(i, 4) = Now_System
    Next i
End Sub

Public Function GetUuid() As String
    GetUuid = Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & GetTodayMilliseconds()
End Function

' ----------------------------------------------------------------
' Procedure  : CreateGUID
' Author     : Dan (webmaster@1stchoiceav.com)
' Source     : http://allapi.mentalis.org/apilist/CDB74B0DFA5C75B7C6AFE60D3295A96F.html
' Adapted by : Mike Wolfe
' Republished: https://nolongerset.com/createguid/
' Date       : 8/5/2022
' ----------------------------------------------------------------
Public Function CreateGUID() As String
    Const S_OK As Long = 0
    Dim Id(0 To 15) As Byte
    Dim Cnt As Long, Guid As String
    If CoCreateGuid(Id(0)) = S_OK Then
        For Cnt = 0 To 15
            CreateGUID = CreateGUID & IIf(Id(Cnt) < 16, "0", "") + Hex$(Id(Cnt))
        Next Cnt
        CreateGUID = Left$(CreateGUID, 8) & "-" & _
                     Mid$(CreateGUID, 9, 4) & "-" & _
                     Mid$(CreateGUID, 13, 4) & "-" & _
                     Mid$(CreateGUID, 17, 4) & "-" & _
                     Right$(CreateGUID, 12)
    Else
        MsgBox "Error while creating GUID!"
    End If
End Function
