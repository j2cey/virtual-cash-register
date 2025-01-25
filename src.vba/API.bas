Option Explicit

'AUTHOR: Ricardo Camisa
'Email: rich.7.2014@gmail.com
'Contact: +244 925341780


Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const GWL_STYLE As Long = (-16)
Private Const WS_CAPTION = 55000000
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const IDC_HAND = 32649&
Public MyForm As Long
Public Style As Long
Public Const CURRENT_STYLE As Long = (-16)
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public CbtStyle         As clsCss
Public UI               As UI
Public Colbtn           As New Collection
Public dpFrom           As DateTimePicker

'-----------------------------------------------------------------------------------------------------------
#If VBA7 Then
    'Author: Ricardo Camisa
    'API qui nous permet de déplacer le formulaire et de le libérer ---------------------------------
    Public Declare PtrSafe Sub ReleaseCapture Lib "user32" ()
    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            lParam As Any) As Long
    '----------------------------------------------------------------------------------------------------
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwNilliseconds As Long)
    Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Public Declare PtrSafe Function MoveWindow Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    'Author: Ricardo Camisa
    'API qui nous permet de déplacer le formulaire et de le libérer ---------------------------------
    Public Declare Sub ReleaseCapture Lib "user32" ()
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            lParam As Any) As Long
    '----------------------------------------------------------------------------------------------------
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwNilliseconds As Long)
    Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Public Declare Function MoveWindow Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Private lngPixelsX As Long
Private lngPixelsY As Long
Private strThunder As String
Private blnCreate As Boolean
Private lnghWnd_Form As Long
Private lnghWnd_Sub As Long
Private colBaseCtrl As Collection
Private Const cstMask As Long = &H7FFFFFFF

Public Function MouseCursor(CursorType As Long)
  Dim lngRet As Long
  lngRet = LoadCursorBynum(0&, CursorType)
  lngRet = SetCursor(lngRet)
End Function

'Méthode qui permet de déplacer le formulaire
Public Sub MoveForm(Form As Object, obj As Object, Button As Integer)
    Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long
    If Val(Application.Version) < 9 Then
        lngMyHandle = FindWindowA("ThunderXFrame", Form.Caption)
    Else
        lngMyHandle = FindWindowA("ThunderDFrame", Form.Caption)
    End If
    
    If Button = 1 Then
        With obj
            Call ReleaseCapture
            Call SendMessage(lngMyHandle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End With
    End If
End Sub

'Dans l'événement de zone de texte Lors du déplacement de la souris, tapez:
'=MouseCursor(32649) => changer la forme de la main
Public Function LittleHand()
    Call MouseCursor(IDC_HAND)
End Function

Public Sub RemoveAll(objForm As Object)
    MyForm = FindWindowA(vbNullString, objForm.Caption)
    Style = Style Or WS_CAPTION
    MoveWindow MyForm, CURRENT_STYLE, (Style)
End Sub

Sub ShowExcel()
Application.Visible = True
Application.ScreenUpdating = True
End Sub

Public Sub CursorHand(MyForm As Object)
    Dim ctrl As Control
    Dim strx As String
    
    strx = ""
    
    For Each ctrl In MyForm.Controls
        If ctrl.Tag = "btn" And TypeName(ctrl) = "CommandButton" Then
            Set CbtStyle = New clsCss
            Set CbtStyle.btnButton = ctrl
            Colbtn.Add CbtStyle
        ElseIf ctrl.Tag = "btn" And TypeName(ctrl) = "Label" Then
            Set CbtStyle = New clsCss
            Set CbtStyle.btnLabel = ctrl
            Colbtn.Add CbtStyle
        ElseIf ctrl.Tag = "btn" And TypeName(ctrl) = "Image" Then
            Set CbtStyle = New clsCss
            Set CbtStyle.btnImage = ctrl
            Colbtn.Add CbtStyle
        ElseIf ctrl.Tag = "btn" And TypeName(ctrl) = "CheckBox" Then
            Set CbtStyle = New clsCss
            Set CbtStyle.btnCheckBox = ctrl
            Colbtn.Add CbtStyle
        ElseIf ctrl.Tag = "Menus" And TypeName(ctrl) = "Frame" Then
            Set UI = New UI
            Set UI.Menus = ctrl
            Colbtn.Add UI
        ElseIf ctrl.Tag = "ContainerMenus" And TypeName(ctrl) = "Frame" Then
            Set UI = New UI
            Set UI.ContainerMenus = ctrl
            Colbtn.Add UI
        End If
    Next
End Sub
