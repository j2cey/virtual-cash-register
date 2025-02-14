Option Explicit

Private m_id As Long
Private m_title As String
Private m_parent As CProgression

Private m_start_timestamp As Double
Private m_end_timestamp As Double
Private m_duration_timestamp As Double
Private m_duration_formatted As String

Private m_nb_todo As Long
Private m_nb_done As Long

Private m_sub_progressions As Collection


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CProgression
    Set Self = Me
End Property

' Id
Public Property Let Id(lId As Long)
    m_id = lId
End Property

Public Property Get Id() As Long
    Id = m_id
End Property

' Title
Public Property Let Title(sTitle As String)
    m_title = sTitle
End Property

Public Property Get Title() As String
    Title = m_title
End Property

' Parent
Public Property Set Parent(oParent As CProgression)
    Set m_parent = oParent
End Property

Public Property Get Parent() As CProgression
    Set Parent = m_parent
End Property

' StartTimestamp
Public Property Let StartTimestamp(dblStartTimestamp As Double)
    m_start_timestamp = dblStartTimestamp
End Property

Public Property Get StartTimestamp() As Double
    StartTimestamp = m_start_timestamp
End Property

' EndTimestamp
Public Property Let EndTimestamp(dblEndTimestamp As Double)
    m_end_timestamp = dblEndTimestamp
End Property

Public Property Get EndTimestamp() As Double
    EndTimestamp = m_end_timestamp
End Property

' DurationTimestamp
Public Property Let DurationTimestamp(dblDurationTimestamp As Double)
    m_duration_timestamp = dblDurationTimestamp
End Property

Public Property Get DurationTimestamp() As Double
    DurationTimestamp = m_duration_timestamp
End Property

' DurationFormatted
Public Property Let DurationFormatted(strDurationFormatted As String)
    m_duration_formatted = strDurationFormatted
End Property

Public Property Get DurationFormatted() As String
    DurationFormatted = m_duration_formatted
End Property

' NbToDo
Public Property Let NbToDo(lngNbToDo As Long)
    m_nb_todo = lngNbToDo
End Property

Public Property Get NbToDo() As Long
    NbToDo = m_nb_todo
End Property

' NbDone
Public Property Let NbDone(lngNbDone As Long)
    m_nb_done = lngNbDone
End Property

Public Property Get NbDone() As Long
    NbDone = m_nb_done
End Property

' SubProgressions
Public Property Set SubProgressions(collSubProgressions As Collection)
    Set m_sub_progressions = collSubProgressions
End Property

Public Property Get SubProgressions() As Collection
    Set SubProgressions = m_sub_progressions
End Property


'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(sTitle As String, lngNbToDo As Long, Optional oParent As CProgression)
    Me.Title = sTitle
    Me.NbToDo = lngNbToDo
    
    If IsMissing(oParent) Or oParent Is Nothing Then
        Me.Id = 0
    Else
        oParent.AddSubProgression Me
    End If
End Sub

Public Sub AddSubProgression(subProgression As CProgression)
    Set subProgression.Parent = Me
    subProgression.Id = Me.SubProgressions.count + 1
    
    Me.SubProgressions.Add subProgression, CStr(subProgression.Id)
    
    Me.AddToDo subProgression.NbToDo, subProgression
End Sub

Public Function StartNewSubProgression(sTitle As String, lngNbToDo As Long) As CProgression
    Set StartNewSubProgression = StartNewProgression(sTitle, lngNbToDo, Me)
End Function

Public Sub Start()
    Me.StartTimestamp = time
    AddDone 0
End Sub

Public Sub AddToDo(lngNbToDo As Long, Optional subProgression As CProgression)
    Me.NbToDo = Me.NbToDo + lngNbToDo
    
    DisplayProgression subProgression
    
    If Not Me.Parent Is Nothing Then
        Me.Parent.AddToDo lngNbToDo, Me
    End If
End Sub

Public Function AddToDoLastSub(lngNbToDo As Long) As CProgression
    Dim subProgression As CProgression, lastIndex As String
    
    If Me.SubProgressions.count <= 0 Then
        Exit Function
    End If
    
    lastIndex = Me.SubProgressions.count
    Set subProgression = Me.SubProgressions.Item(lastIndex)
    subProgression.AddToDo lngNbToDo
    
    Set AddToDoLastSub = subProgression
End Function

Public Sub AddDone(lngNbDone As Long, Optional bEndIfAny As Boolean = False, Optional subProgression As CProgression)
    Me.NbDone = Me.NbDone + lngNbDone
    
    If Me.NbDone > Me.NbToDo Then
        Me.NbDone = Me.NbToDo
    End If
    
    SetIntervalTime
    
    'If Me.NbDone = Me.NbToDo And bEndIfAny Then
    '    SetIntervalTime
    'End If
    
    DisplayProgression subProgression
    
    If Not Me.Parent Is Nothing Then
        Me.Parent.AddDone lngNbDone, bEndIfAny, Me
    End If
End Sub

Public Function AddDoneLastSub(lngNbDone As Long, Optional bEndIfAny As Boolean = False) As CProgression
    Dim subProgression As CProgression, lastIndex As String
    
    If Me.SubProgressions.count <= 0 Then
        Exit Function
    End If
    
    lastIndex = Me.SubProgressions.count
    Set subProgression = Me.SubProgressions.Item(lastIndex)
    subProgression.AddDone lngNbDone, bEndIfAny
    
    Set AddDoneLastSub = subProgression
End Function



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Private Sub SetIntervalTime()
    Me.EndTimestamp = time
    Me.DurationTimestamp = Me.EndTimestamp - Me.StartTimestamp
    Me.DurationFormatted = Format(Me.DurationTimestamp, "hh:mm:ss")
    
    'Me.Title = Me.Title & "... FIN !"
End Sub

Private Sub DisplayProgression(Optional subProgression As CProgression)
    Dim pctdone As Single
    
    MainUFrm.ProgressContainerFrme.Visible = True
    
    pctdone = Me.NbDone / Me.NbToDo
    With MainUFrm
        .ProgressMsgLbl.Caption = Me.Title
        .ProgressBarFrme.Caption = CStr(Round(pctdone, 2) * 100) & " %"
        .LabelProgress.Width = pctdone * (.ProgressBarFrme.Width)
        .ProgressTimeLbl.Caption = Me.DurationFormatted
        
        '.Repaint
    End With
    
    DoEvents
    
    If Me.NbDone >= Me.NbToDo Then
        'Application.Wait (Now + TimeValue("00:00:01"))
        
        If Me.Parent Is Nothing And (IsMissing(subProgression) Or subProgression Is Nothing) Then
            MainUFrm.ProgressContainerFrme.Visible = False
        End If
    End If
End Sub


'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************


Private Sub Class_Initialize()
    m_start_timestamp = 0
    m_end_timestamp = 0
    m_duration_timestamp = 0
    m_duration_formatted = "00:00:00"
    
    m_nb_todo = 0
    m_nb_done = 0
    
    Set m_sub_progressions = New Collection
End Sub
