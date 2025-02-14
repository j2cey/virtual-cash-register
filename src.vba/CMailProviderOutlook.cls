Option Explicit


Implements IMAILPROVIDER

Private m_mail_parameter As CMailParameter
Private m_email_app As Outlook.Application
Private m_email_item As Outlook.MailItem



'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CMailProviderOutlook
    Set Self = Me
End Property

' MailParameter
Public Property Set MailParameter(oMailParameter As CMailParameter)
    Set m_mail_parameter = oMailParameter
End Property

Public Property Get MailParameter() As CMailParameter
    Set MailParameter = m_mail_parameter
End Property

' EmailApp
Public Property Set EmailApp(oEmailApp As Outlook.Application)
    Set m_email_app = oEmailApp
End Property

Public Property Get EmailApp() As Outlook.Application
    Set EmailApp = m_email_app
End Property

' EmailItem
Public Property Set EmailItem(oEmailItem As Outlook.MailItem)
    Set m_email_item = oEmailItem
End Property

Public Property Get EmailItem() As Outlook.MailItem
    Set EmailItem = m_email_item
End Property



'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Sub SendMail(ByVal sSubject As String, ByVal sMailBody As String, ByVal sMailTo As String, ByVal sMailCC As String)
    
    Set Me.EmailApp = New Outlook.Application
    Set Me.EmailItem = EmailApp.CreateItem(olMailItem)
    
    Me.EmailItem.To = sMailTo
    Me.EmailItem.CC = sMailCC
    
    Me.EmailItem.Subject = sSubject
    Me.EmailItem.HTMLBody = sMailBody
    
    Me.EmailItem.Send
End Sub

Public Sub SendMailToMany(ByVal sSubject As String, ByVal sMailBody As String, recipientsArr As Variant, Optional attachementFile As String)
    Dim i As Integer, sTo As String
    Dim strFileExists As String
    
    If IsEmpty(recipientsArr) Then
        Exit Sub
    End If
    
    Set Me.EmailApp = New Outlook.Application
    Set Me.EmailItem = EmailApp.CreateItem(olMailItem)
    
    'Me.EmailItem.SentOnBehalfOfName = Me.MailParameter.MailSenderName
    
    'Me.EmailItem.To = sMailTo
    'Me.EmailItem.CC = sMailCC
    
    Me.EmailItem.Subject = sSubject
    Me.EmailItem.HTMLBody = sMailBody
    
    'For i = 0 To UBound(recipientsArr)
    '    Me.EmailItem.Recipients.Add (CStr(recipientsArr(i)))
    'Next i
    
    For i = 0 To UBound(recipientsArr)
        sTo = sTo & ";" & CStr(recipientsArr(i)(1))
    Next
    
    sTo = Right(sTo, Len(sTo) - 1)
    Me.EmailItem.To = sTo
    
    If Not IsMissing(attachementFile) And Not IsEmpty(attachementFile) Then
        If (Not attachementFile = "") Then
            Me.EmailItem.Attachments.Add (attachementFile)
        End If
    End If
    
    Me.EmailItem.Send
End Sub



'*****************************************************************************************************************************
'    Interfaces Implementation
'*****************************************************************************************************************************

'    Public Properties
'******************************************************************

' MailParameter
Public Property Set IMAILPROVIDER_MailParameter(RHS As CMailParameter)
    Set m_mail_parameter = RHS
End Property

Public Property Get IMAILPROVIDER_MailParameter() As CMailParameter
    Set IMAILPROVIDER_MailParameter = m_mail_parameter
End Property


'    Public functions and subs
'******************************************************************

Private Sub IMAILPROVIDER_SendMail(ByVal sSubject As String, ByVal sMailBody As String, ByVal sMailTo As String, Optional ByVal sMailCC As String)
    Me.SendMail sSubject, sMailBody, sMailTo, sMailCC
End Sub

Private Sub IMAILPROVIDER_SendMailTomany(ByVal sSubject As String, ByVal sMailBody As String, recipientsArr As Variant, Optional attachementFile As String)
    Me.SendMailToMany sSubject, sMailBody, recipientsArr, attachementFile
End Sub