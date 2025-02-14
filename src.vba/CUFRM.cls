'Build 000
'*****************************************************************************************************************************
'
' Authors:  Jude Parfait NGOM NZE, jud10parfait@gmail.com
'
' (c)2016-2017, all rights reserved to the authors
'
'*****************************************************************************************************************************

'-------------------------------------------------------------------------
' Module    : CUFRM
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 18-07-2017
' Purpose   : Manage user form
'-------------------------------------------------------------------------
Option Explicit

Implements IUCONTROL


Private WithEvents m_oUFrm As MSForms.UserForm
Private m_currCtl As CUIControl
Private m_currUFrm As CUFRM

' list of controls for this control
Private m_subColl As CUIControlList

Private m_collUfrm As Collection

Private Const C_STATUS_INIT = 0
Private Const C_STATUS_ADD = 1
Private Const C_STATUS_UPD_DEL = 2

Private m_sName As String
Private m_iStatus As Integer ' Form's Status
Private m_sAction As String

Private m_vLockedevents As String

Private m_ImageBtnNormalSize As Variant ' Width, Height
Private m_ImageBtnMoveSize As Variant ' Width, Height

Private Const C_NOTIFY_SUCCESS = 1
Private Const C_NOTIFY_ERROR = -1
Private Const C_NOTIFY_WARNING = 2

Private Const mcGWL_STYLE = (-16)
Private Const mcWS_SYSMENU = &H80000

'Windows API calls to handle windows
'Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
'    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
'Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
'    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long


    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Private Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long


    Private Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long


    Private Declare Function DrawMenuBar Lib "user32" _
               (ByVal hWnd As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16 ' reference to the standard style bits
Private Const WS_CAPTION As Long = &HC00000 ' remove TitleBar
Private Declare PtrSafe Function DrawMenuBar2 Lib "user32" (ByVal hwnd As LongPtr) As Long



'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CUFRM
    Set Self = Me
End Property

Public Property Get uFrm() As MSForms.UserForm
' Created   : 18-07-2017
' Purpose   : Return the userform object of the current class object
'----------------------------------------------------------------------
    Set uFrm = m_oUFrm
End Property

Public Property Get CtlsList() As CUIControlList
' Created   : 18-07-2017
' Purpose   : Return CUIControl collection of userform controls
'----------------------------------------------------------------------
    Set CtlsList = m_subColl
End Property

Public Property Let Status(iStatus As Integer)
    
    m_iStatus = iStatus
    
    If m_iStatus = C_STATUS_INIT Then
        m_sAction = "Initialization"
    ElseIf m_iStatus = C_STATUS_ADD Then
        m_sAction = "Add new " & m_sName
    ElseIf m_iStatus = C_STATUS_UPD_DEL Then
        m_sAction = "Update/Delete " & m_sName
    Else
        m_sAction = "Missing Form status"
    End If
    
End Property

Public Property Get Status() As Integer
    Status = m_iStatus
End Property

Public Property Let ImageBtnNormalSize(vImageBtnNormalSize) ' Width, Height
    m_ImageBtnNormalSize = vImageBtnNormalSize
End Property

Public Property Get ImageBtnNormalSize() As Variant
    ImageBtnNormalSize = m_ImageBtnNormalSize
End Property

Public Property Let ImageBtnMoveSize(vImageBtnMoveSize) ' Width, Height
    m_ImageBtnMoveSize = vImageBtnMoveSize
End Property

Public Property Get ImageBtnMoveSize() As Variant
    ImageBtnMoveSize = m_ImageBtnMoveSize
End Property




'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Sub HideBar()
    Dim Style As Long, Menu As Long, hWndForm As Long
    hWndForm = FindWindow("ThunderDFrame", Me.uFrm.Caption)
    
    Style = GetWindowLong(hWndForm, &HFFF0)
    Style = Style And Not &HC00000
    SetWindowLong hWndForm, &HFFF0, Style
    
    DrawMenuBar hWndForm
    
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Purpose: Set a UserForm to have no title bar ... obviously, no icon or 'red X' can be shown
' ---------------------------------------------------------------------------------------------------------------------
Public Sub SetNoTitleBar()
    Dim hwnd As LongPtr
    On Error GoTo errExit
    hwnd = FindWindow("ThunderDFrame", Me.uFrm.Caption)
    If hwnd <> 0 Then
        SetWindowLongPtr hwnd, GWL_STYLE, (GetWindowLongPtr(hwnd, GWL_STYLE) And Not WS_CAPTION)
        DrawMenuBar2 hwnd
    End If
errExit:
End Sub

Public Sub Init(uFrm As MSForms.UserForm, Optional imgbtnsave As String = "", _
Optional imgbtndelete As String = "", Optional imgbtncancel As String = "", Optional imgbtnsearch As String = "", _
Optional imgbtnclose As String = "", Optional sMainList As String = "")
' Created   : 18-07-2017
' Purpose   : Initialize a class object
'
' Return
'
' UFrm          : the userform
' vMainSource   : the main source
' imgbtnsave    : the save button name
' imgbtndelete  : the delete button name
' imgbtncancel  : the cancel button name
' imgbtnsearch  : the search button name
' imgbtnclose   : the close button name
' sMainList     : the main list name
'----------------------------------------------------------------------
    
    Set m_oUFrm = uFrm
    
    If Not sMainList = "" Then
        Call Me.AddCtl(m_oUFrm.Controls(sMainList))
    End If
    
    If Not imgbtnsave = "" Then
        Call Me.AddCtl(m_oUFrm.Controls(imgbtnsave))
    End If
    
    If Not imgbtndelete = "" Then
        Call Me.AddCtl(m_oUFrm.Controls(imgbtndelete))
    End If
    
    If Not imgbtncancel = "" Then
        Call Me.AddCtl(m_oUFrm.Controls(imgbtncancel))
    End If
    
    If Not imgbtnsearch = "" Then
        Call Me.AddCtl(m_oUFrm.Controls(imgbtnsearch))
    End If
    
    If Not imgbtnclose = "" Then
        Call Me.AddCtl(m_oUFrm.Controls(imgbtnclose))
    End If
    
    RemoveCloseButton m_oUFrm
    
End Sub

Public Function AddCtl(ctl As MSForms.Control) As CUIControl
' Created   : 28-07-2017
' Purpose   : Add a control to the current object
'
' Return
'
' Ctl               : the control
'----------------------------------------------------------------------
    
    Dim i As Integer, oCurrCtl As CUIControl
    
    Set oCurrCtl = New CUIControl
    
    oCurrCtl.SetCtl ctl, Me
    
    m_subColl.AddCtlByUCTL oCurrCtl
    
    Set AddCtl = oCurrCtl
End Function

Public Sub AddUserForm(uFrm As CUFRM)
    m_collUfrm.Add uFrm
End Sub

Public Sub AddUCtlCriterion(sCtlName As String, sField As String, Val, bClearbefor As Boolean)
' Created   : 31-07-2017
' Purpose   : Add a criterion to a control
'
' Return
'
' sCtlName          : control name
' sField            : field criteria
' val               : criteria value
' bClearbefor       : determine whether criteria have to be cleared before
'----------------------------------------------------------------------
    Set m_currCtl = New CUIControl
    Set m_currCtl = m_subColl(sCtlName)
    
End Sub

Public Sub ClearUCtlCriterion(sCtlName As String)
' Created   : 31-07-2017
' Purpose   : clear control's criteria
'
' Return
'
' sCtlName          : control name
'----------------------------------------------------------------------
    Set m_currCtl = New CUIControl
    Set m_currCtl = m_subColl(sCtlName)
    
End Sub

Public Sub FileSizeCtl(sCtlName As String, sFileSizeCtl As String)
    Set m_currCtl = New CUIControl
    Set m_currCtl = m_subColl(sCtlName)
    
    m_currCtl.FileSizeCtl = sFileSizeCtl
End Sub

Public Sub RemoveCloseButton(objForm As Object)
    Dim lngStyle As LongPtr
    Dim lngHWnd As LongPtr

    Dim lpClassName As String
    lpClassName = vbNullString
    If Val(Application.Version) >= 9 Then
       lpClassName = "ThunderDFrame"
    Else
       lpClassName = "ThunderXFrame"
    End If

    lngHWnd = FindWindow(lpClassName, objForm.Caption)
    lngStyle = GetWindowLongPtr(lngHWnd, mcGWL_STYLE)

    If lngStyle And mcWS_SYSMENU > 0 Then
        SetWindowLongPtr lngHWnd, mcGWL_STYLE, (lngStyle And Not mcWS_SYSMENU)
    End If
End Sub



'*****************************************************************************************************************************
'    Private functions and subs
'*****************************************************************************************************************************

Private Sub LockEvent(eventName As String)
' Add an event name in the locked events list
'   eventName:  The event name

    If IsEventLocked(eventName) Then
        ' Event already locked
    Else
        ' Add new event to the locked list
        If m_vLockedevents = "" Or IsEmpty(m_vLockedevents) Or IsNull(m_vLockedevents) Then
            m_vLockedevents = eventName
        Else
            m_vLockedevents = m_vLockedevents & "|" & eventName
        End If
        
    End If
    
End Sub

Private Sub UnlockEvent(eventName As String)
' Withdraw an event name from the locked events list
'   eventName:  The event name

    Dim pos As Integer
    pos = EventLocked_Position(eventName)
    
    If pos > -1 Then
        Dim lockedevents_tab As Variant
        lockedevents_tab = Split(m_vLockedevents, "|")
        lockedevents_tab(pos) = " "
        
        m_vLockedevents = Trim(Join(lockedevents_tab, "|"))
        
        If m_vLockedevents = "|" Then
            m_vLockedevents = ""
        End If
        
    Else
        ' This event is not in the locked list
    End If
End Sub

Private Function IsEventLocked(eventName As String) As Boolean
' Check whether an event name is in the locked events list
' Returns True if found or False if not
'   eventName:  The event name

    IsEventLocked = (EventLocked_Position(eventName) > -1)
End Function

Private Function EventLocked_Position(eventName As String) As Integer
' Find an event name prosition from the locked events list
' Returns the position of the event if found (-1 if not found)
'   eventName:  The event name

    Dim pos As Integer
    
    pos = -1
    If m_vLockedevents = "" Then
        ' No event in the locked list
    Else
        Dim lockedevents_tab As Variant, i As Integer
        lockedevents_tab = Split(m_vLockedevents, "|")
        
        For i = 0 To UBound(lockedevents_tab)
            If lockedevents_tab(i) = eventName Then
                pos = i
                Exit For
            End If
        Next i
    End If
    
    EventLocked_Position = pos
End Function

Public Sub ImageBtnSetMoveSize(img As MSForms.Image)
    Me.CtlsList.ImageBtnApplyNormalSizeAll
    img.Width = CDbl(m_ImageBtnMoveSize(0))
    img.Height = CDbl(m_ImageBtnMoveSize(1))
End Sub

Private Function GetCtlName(sCtlName As String) As String
    GetCtlName = Split(sCtlName, ":")(0)
End Function

Private Function GetCtlValueType(sCtlName As String) As String
    GetCtlValueType = Split(sCtlName, ":")(1)
End Function

Public Function ConfirmAction(sMsg As String) As Boolean
    ' message box result for confirmation
    Dim MsgBoxResult As Integer
    
    MsgBoxResult = MsgBox(sMsg, vbYesNo + vbDefaultButton2 + vbExclamation, GetAppName)
    ConfirmAction = (MsgBoxResult = vbYes)
End Function

Public Sub Notify(heads, msgs, iNotifyType As Integer)
    Dim msgboxStyle As VbMsgBoxStyle, sMsg As String, i As Integer, j As Integer
    
    If Not IsEmpty(heads) Then
        
        Dim headsMsgs As Variant
        
        ' Get distinct heads
        For i = 0 To UBound(heads)
            If i = 0 Then
                headsMsgs = Arrays.AddToArray(headsMsgs, heads(i))
            Else
                If Arrays.IsInArray(headsMsgs, heads(i)) = -1 Then
                    headsMsgs = Arrays.AddToArray(headsMsgs, heads(i))
                End If
            End If
        Next i
        
        ' Add messages to relevant head
        For i = 0 To UBound(msgs)
            For j = 0 To UBound(headsMsgs)
                
                Dim currhead As String
                currhead = Split(headsMsgs(j), vbNewLine)(0)
                
                If heads(i) = currhead Then
                    headsMsgs(j) = headsMsgs(j) & vbNewLine & msgs(i)
                End If
            Next j
        Next i
        
        ' build the final message
        For i = 0 To UBound(headsMsgs)
            If i = 0 Then
                sMsg = headsMsgs(i)
            Else
                sMsg = sMsg & vbNewLine & vbNewLine & headsMsgs(i)
            End If
        Next i
        
        If iNotifyType = C_NOTIFY_SUCCESS Then
            msgboxStyle = vbInformation
        ElseIf iNotifyType = C_NOTIFY_ERROR Then
            msgboxStyle = vbCritical
        ElseIf iNotifyType = C_NOTIFY_WARNING Then
            msgboxStyle = vbExclamation
        Else
            msgboxStyle = vbOKOnly
        End If
    Else
        For i = 0 To UBound(msgs)
            If i = 0 Then
                sMsg = msgs(i)
            Else
                sMsg = sMsg & vbNewLine & msgs(i)
            End If
        Next i
    End If
    
    MsgBox sMsg, msgboxStyle, GetAppName
End Sub



'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************

Private Sub Class_Initialize()
    Set m_subColl = New CUIControlList
    Set m_collUfrm = New Collection
    Me.Status = C_STATUS_ADD
    
    Me.ImageBtnNormalSize = Array(18, 18)
    Me.ImageBtnMoveSize = Array(20, 20)
End Sub

Private Sub m_oUFrm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.CtlsList.ImageBtnApplyNormalSizeAll
End Sub


'*****************************************************************************************************************************
'    Interfaces Implementation
'*****************************************************************************************************************************

Private Sub IUCONTROL_ImageBtnSetMoveSize(img As MSForms.Image)
    Me.ImageBtnSetMoveSize img
End Sub



