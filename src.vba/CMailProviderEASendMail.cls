Option Explicit

Implements IMAILPROVIDER


Private m_mail_parameter As CMailParameter
Private m_SmtpObject As EASendMailObjLib.Mail
Private m_BodyFormat As Integer

'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CMailProviderEASendMail
    Set Self = Me
End Property

' MailParameter
Public Property Set MailParameter(oMailParameter As CMailParameter)
    Set m_mail_parameter = oMailParameter
End Property

Public Property Get MailParameter() As CMailParameter
    Set MailParameter = m_mail_parameter
End Property

' SmtpObject
Public Property Set SmtpObject(oSmtpObject As EASendMailObjLib.Mail)
    Set m_SmtpObject = oSmtpObject
End Property

Public Property Get SmtpObject() As EASendMailObjLib.Mail
    Set SmtpObject = m_SmtpObject
End Property

' BodyFormat
Public Property Let BodyFormat(iBodyFormat As Integer)
    Let m_BodyFormat = iBodyFormat
End Property

Public Property Get BodyFormat() As Integer
    BodyFormat = m_BodyFormat
End Property


'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Sub SendMail(ByVal sSubject As String, ByVal sMailBody As String, ByVal sMailTo As String, ByVal sMailCC As String)
    
    If Me.MailParameter Is Nothing Then
        Exit Sub
    End If
    
    'Sender = "john@company.com"
    'Name = "John"
    'Address = "jane@company.com"
    'Subject = "Test subject"
    'body = "Test body"
    
    Me.BodyFormat = 0
    
    Set Me.SmtpObject = New EASendMailObjLib.Mail
    Me.SmtpObject.LicenseCode = "TryIt" ' Here goes your license code for the software; for now, we are using the trial version

    ' Please change the server address, username, and password to the ones you will be using
    Me.SmtpObject.ServerAddr = Me.MailParameter.MailServerAddress
    Me.SmtpObject.UserName = Me.MailParameter.MailUserName
    Me.SmtpObject.Password = Me.MailParameter.MailUserPassword

    Me.SmtpObject.ServerPort = Me.MailParameter.MailServerPort '587

    ' Using TryTLS,
    ' If the SMTP server supports TLS, then a TLS connection is used; otherwise, a normal TCP connection is used.
    ' https://www.emailarchitect.net/easendmail/sdk/?ct=connecttype
    Me.SmtpObject.ConnectType = 4

    Me.SmtpObject.FromAddr = Me.MailParameter.MailSender
    Me.SmtpObject.AddRecipient Me.MailParameter.MailSenderName, sMailTo, 0

    Me.SmtpObject.Subject = sSubject
    Me.SmtpObject.BodyFormat = Me.BodyFormat
    Me.SmtpObject.BodyText = sMailBody

    Me.SmtpObject.Asynchronous = 1
    Me.SmtpObject.SendMail
    Set Me.SmtpObject = Nothing
    
End Sub

Public Sub SendMailToMany(ByVal sSubject As String, ByVal sMailBody As String, recipientsArr As Variant, Optional attachementFile As String)
End Sub



'*****************************************************************************************************************************
'    Interfaces Implementation
'*****************************************************************************************************************************

'    Public Properties
'******************************************************************

' MailParameter
Public Property Set IMAILPROVIDER_MailParameter(RHS As MailParameter)
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