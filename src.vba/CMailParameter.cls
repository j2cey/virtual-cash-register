'Build 000
'*****************************************************************************************************************************
'
' Authors:  Jude Parfait NGOM NZE, jud10parfait@gmail.com
'
' (c)2024, all rights reserved to the authors
'
'*****************************************************************************************************************************
'-------------------------------------------------------------------------
' Module    : clsMailParameter
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 12-07-2024
' Purpose   : Manage Mail Provider Parameters
'-------------------------------------------------------------------------
Option Explicit


Private m_provider_name As String
Private m_mail_sender As String
Private m_mail_sender_name As String
Private m_mail_server_address As String
Private m_mail_server_port As String
Private m_mail_user_name As String
Private m_mail_user_password As String
Private m_provider_description As String



'*****************************************************************************************************************************
'    Public events
'*****************************************************************************************************************************



'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As CMailParameter
    Set Self = Me
End Property

' ProviderName
Public Property Let ProviderName(sProviderName As String)
    m_provider_name = sProviderName
End Property

Public Property Get ProviderName() As String
    ProviderName = m_provider_name
End Property

' MailSender
Public Property Let MailSender(sSender As String)
    m_mail_sender = sSender
End Property

Public Property Get MailSender() As String
    MailSender = m_mail_sender
End Property

' MailSenderName
Public Property Let MailSenderName(sSenderName As String)
    m_mail_sender_name = sSenderName
End Property

Public Property Get MailSenderName() As String
    If m_mail_sender_name = "@Nom-Application" Then
        MailSenderName = GetAppName
    Else
        MailSenderName = m_mail_sender_name
    End If
End Property

' MailServerAddress
Public Property Let MailServerAddress(sServerAddress As String)
    m_mail_server_address = sServerAddress
End Property

Public Property Get MailServerAddress() As String
    MailServerAddress = m_mail_server_address
End Property

' MailServerPort
Public Property Let MailServerPort(sServerPort As String)
    m_mail_server_port = sServerPort
End Property

Public Property Get MailServerPort() As String
    MailServerPort = m_mail_server_port
End Property

' MailUserName
Public Property Let MailUserName(sUserName As String)
    m_mail_user_name = sUserName
End Property

Public Property Get MailUserName() As String
    MailUserName = m_mail_user_name
End Property

' MailUserPassword
Public Property Let MailUserPassword(sUserPassword As String)
    m_mail_user_password = sUserPassword
End Property

Public Property Get MailUserPassword() As String
    MailUserPassword = m_mail_user_password
End Property

' ProviderDescription
Public Property Let ProviderDescription(sProviderDescription As String)
    m_provider_description = sProviderDescription
End Property

Public Property Get ProviderDescription() As String
    ProviderDescription = m_provider_description
End Property



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init()
    
End Sub



'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************


Private Sub Class_Initialize()
    
End Sub

