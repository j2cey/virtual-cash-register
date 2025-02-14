'Build 000
'*****************************************************************************************************************************
'
' Authors:  Jude Parfait NGOM NZE, jud10parfait@gmail.com
'
' (c)2024, all rights reserved to the authors
'
'*****************************************************************************************************************************
'-------------------------------------------------------------------------
' Module    : clsMail
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 12-07-2024
' Purpose   : Manage E-Mail
'-------------------------------------------------------------------------
Option Explicit


Private m_MailProvider As IMAILPROVIDER
Private m_sSubject As String



'*****************************************************************************************************************************
'    Public events
'*****************************************************************************************************************************



'*****************************************************************************************************************************
'   Public Properties
'*****************************************************************************************************************************

Public Property Get Self() As clsMail
    Set Self = Me
End Property

' MailProvider
Public Property Set MailProvider(oMailProvider As IMAILPROVIDER)
    Set m_MailProvider = oMailProvider
End Property

Public Property Get MailProvider() As IMAILPROVIDER
    Set MailProvider = m_MailProvider
End Property

' MailSubject
Public Property Let MailSubject(sSubject As String)
    Let m_sSubject = sSubject
End Property

Public Property Get MailSubject() As String
    EmailItem = m_sSubject
End Property



'*****************************************************************************************************************************
'    Public functions and subs
'*****************************************************************************************************************************

Public Sub Init(ByVal sProvider As String, Optional ByVal oMailParameter As CMailParameter)
    If sProvider = "Outlook" Then
        Set Me.MailProvider = New clsMailProviderOutlook
    ElseIf sProvider = "EASendMail" Then
        Set Me.MailProvider = New clsMailProviderEASendMail
    End If
    
    If Not Me.MailProvider Is Nothing Then
        Set Me.MailProvider.MailParameter = oMailParameter
    End If
End Sub


Public Sub SendMail(ByVal sSubject As String, ByVal sMailBody As String, ByVal sMailTo As String, Optional ByVal sMailCC As String)
    Me.MailProvider.SendMail sSubject, sMailBody, sMailTo, sMailCC
End Sub

Public Sub SendMailToMany(ByVal sSubject As String, ByVal sMailBody As String, recipientsArr As Variant, Optional attachementFile As String)
    
    If IsEmpty(recipientsArr) Then
        Exit Sub
    End If
    
    Dim finalBody As String

    finalBody = HtmlHeader & sMailBody & HtmlFooter

    If (Not IsMissing(attachementFile)) And (Not IsEmpty(attachementFile)) And (Not attachementFile = "") Then
        Me.MailProvider.SendMailToMany sSubject, finalBody, recipientsArr, attachementFile
    Else
        Me.MailProvider.SendMailToMany sSubject, finalBody, recipientsArr
    End If
End Sub

Public Sub SetAttachement()

End Sub

'*****************************************************************************************************************************
'    Private functions and subs
'*****************************************************************************************************************************

Private Function HtmlHeader_Old() As String
    Dim sHeader As String

    sHeader = "<!DOCTYPE html><html><body>"
    sHeader = sHeader & "<div style=""font-family:'Segoe UI', Calibri, Arial, Helvetica; font-size: 14px; max-width: 768px; color:#0b5394;"">"
    
    HtmlHeader = sHeader
End Function

Private Function HtmlFooter_Old() As String
    Dim sFooter As String

    sFooter = "<br><br><br><br>"
    sFooter = sFooter & "<p style=""margin: 0; padding-top:0;"">Cordialement, Regards.</p>"
    sFooter = sFooter & "<br>"
    If Not GetAppLogo = "" Then
        sFooter = sFooter & "<p style=""margin : 0; padding-top:0;""><img border=0 src='" & GetAppLogo & "'width='162px' height='63px'></p>"
    End If
    sFooter = sFooter & "<p style=""margin: 0; padding-top:0;""><B>" & GetAppName & "</B></p>"
    sFooter = sFooter & "</div></body></html>"
    
    HtmlFooter = sFooter
End Function


Private Function HtmlHeader_Old2() As String
    Dim sHeader As String

    sHeader = "<!doctype html>"
    sHeader = sHeader & "<html lang=""en"">"
    sHeader = sHeader & "<head>"
    sHeader = sHeader & "    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">"
    sHeader = sHeader & "    <meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    sHeader = sHeader & "    <title>Simple Transactional Email</title>"
    sHeader = sHeader & "    <style media=""all"" type=""text/css"">"
    sHeader = sHeader & "@media all {"
    sHeader = sHeader & ".btn-primary table td:hover {"
    sHeader = sHeader & "    background-color: #ec0867 !important;"
    sHeader = sHeader & "}"

    sHeader = sHeader & ".btn-primary a:hover {"
    sHeader = sHeader & "    background-color: #ec0867 !important;"
    sHeader = sHeader & "    border-color: #ec0867 !important;"
    sHeader = sHeader & "}"
    sHeader = sHeader & "}"
    sHeader = sHeader & "@media only screen and (max-width: 640px) {"
    sHeader = sHeader & ".main p,"
    sHeader = sHeader & ".main td,"
    sHeader = sHeader & ".main span {"
    sHeader = sHeader & "    font-size: 16px !important;"
    sHeader = sHeader & "}"
    sHeader = sHeader & " "
    sHeader = sHeader & ".wrapper {"
    sHeader = sHeader & "    padding: 8px !important;"
    sHeader = sHeader & "}"

    sHeader = sHeader & ".content {"
    sHeader = sHeader & "    padding: 0 !important;"
    sHeader = sHeader & "}"

    sHeader = sHeader & ".container {"
    sHeader = sHeader & "   padding: 0 !important;"
    sHeader = sHeader & "    padding-top: 8px !important;"
    sHeader = sHeader & "    width: 100% !important;"
    sHeader = sHeader & "}"

    sHeader = sHeader & ".main {"
    sHeader = sHeader & "    border-left-width: 0 !important;"
    sHeader = sHeader & "    border-radius: 0 !important;"
    sHeader = sHeader & "    border-right-width: 0 !important;"
    sHeader = sHeader & "}"

    sHeader = sHeader & ".btn table {"
    sHeader = sHeader & "    max-width: 100% !important;"
    sHeader = sHeader & "    width: 100% !important;"
    sHeader = sHeader & "}"

    sHeader = sHeader & ".btn a {"
    sHeader = sHeader & "    font-size: 16px !important;"
    sHeader = sHeader & "    max-width: 100% !important;"
    sHeader = sHeader & "    width: 100% !important;"
    sHeader = sHeader & "}"
    sHeader = sHeader & "}"
    sHeader = sHeader & "@media all {"
    sHeader = sHeader & ".ExternalClass {"
    sHeader = sHeader & "    width: 100%;"
    sHeader = sHeader & "}"

    sHeader = sHeader & ".ExternalClass,"
    sHeader = sHeader & ".ExternalClass p,"
    sHeader = sHeader & ".ExternalClass span,"
    sHeader = sHeader & ".ExternalClass font,"
    sHeader = sHeader & ".ExternalClass td,"
    sHeader = sHeader & ".ExternalClass div {"
    sHeader = sHeader & "    line-height: 100%;"
    sHeader = sHeader & "}"

    sHeader = sHeader & ".apple-link a {"
    sHeader = sHeader & "    color: inherit !important;"
    sHeader = sHeader & "    font-family: inherit !important;"
    sHeader = sHeader & "    font-size: inherit !important;"
    sHeader = sHeader & "    font-weight: inherit !important;"
    sHeader = sHeader & "    line-height: inherit !important;"
    sHeader = sHeader & "    text-decoration: none !important;"
    sHeader = sHeader & "}"

    sHeader = sHeader & "#MessageViewBody a {"
    sHeader = sHeader & "    color: inherit;"
    sHeader = sHeader & "    text-decoration: none;"
    sHeader = sHeader & "    font-size: inherit;"
    sHeader = sHeader & "    font-family: inherit;"
    sHeader = sHeader & "    font-weight: inherit;"
    sHeader = sHeader & "    line-height: inherit;"
    sHeader = sHeader & "}"
    sHeader = sHeader & "}"
    sHeader = sHeader & "</style>"
    sHeader = sHeader & "</head>"
    sHeader = sHeader & "<body style=""font-family: Helvetica, sans-serif; -webkit-font-smoothing: antialiased; font-size: 16px; line-height: 1.3; -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100%; background-color: #f4f5f6; margin: 0; padding: 0;"">"
    sHeader = sHeader & "    <table role=""presentation"" border=""0"" cellpadding=""0"" cellspacing=""0"" class=""body"" style=""border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #f4f5f6; width: 100%;"" width=""100%"" bgcolor=""#f4f5f6"">"
    sHeader = sHeader & "    <tr>"
    sHeader = sHeader & "        <td style=""font-family: Helvetica, sans-serif; font-size: 16px; vertical-align: top;"" valign=""top"">&nbsp;</td>"
    sHeader = sHeader & "        <td class=""container"" style=""font-family: Helvetica, sans-serif; font-size: 16px; vertical-align: top; max-width: 600px; padding: 0; padding-top: 24px; width: 600px; margin: 0 auto;"" width=""600"" valign=""top"">"
    sHeader = sHeader & "        <div class=""content"" style=""box-sizing: border-box; display: block; margin: 0 auto; max-width: 600px; padding: 0;"">"

    sHeader = sHeader & "            <!-- START CENTERED WHITE CONTAINER -->"
    sHeader = sHeader & "            <span class=""preheader"" style=""color: transparent; display: none; height: 0; max-height: 0; max-width: 0; opacity: 0; overflow: hidden; mso-hide: all; visibility: hidden; width: 0;"">This is preheader text. Some clients will show this text as a preview.</span>"
    sHeader = sHeader & "            <table role=""presentation"" border=""0"" cellpadding=""0"" cellspacing=""0"" class=""main"" style=""border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; background: #ffffff; border: 1px solid #eaebed; border-radius: 16px; width: 100%;"" width=""100%"">"

    sHeader = sHeader & "            <!-- START MAIN CONTENT AREA -->"
    sHeader = sHeader & "            <tr>"
    sHeader = sHeader & "                <td class=""wrapper"" style=""font-family: Helvetica, sans-serif; font-size: 16px; vertical-align: top; box-sizing: border-box; padding: 24px;"" valign=""top"">"
    sHeader = sHeader & "                <p style=""font-family: Helvetica, sans-serif; font-size: 16px; font-weight: normal; margin: 0; margin-bottom: 16px;"">Hi there</p>"
    sHeader = sHeader & "                <p style=""font-family: Helvetica, sans-serif; font-size: 16px; font-weight: normal; margin: 0; margin-bottom: 16px;"">Sometimes you just want to send a simple HTML email with a simple design and clear call to action. This is it.</p>"
    'sHeader = sHeader & "                <table role=""presentation"" border=""0"" cellpadding=""0"" cellspacing=""0"" class=""btn btn-primary"" style=""border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; box-sizing: border-box; width: 100%; min-width: 100%;"" width=""100%"">"
    'sHeader = sHeader & "                    <tbody>"
    'sHeader = sHeader & "                    <tr>"
    'sHeader = sHeader & "                        <td align=""left"" style=""font-family: Helvetica, sans-serif; font-size: 16px; vertical-align: top; padding-bottom: 16px;"" valign=""top"">"
    'sHeader = sHeader & "                        <table role=""presentation"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: auto;"">"
    'sHeader = sHeader & "                            <tbody>"
    'sHeader = sHeader & "                            <tr>"
    'sHeader = sHeader & "                                <td style=""font-family: Helvetica, sans-serif; font-size: 16px; vertical-align: top; border-radius: 4px; text-align: center; background-color: #0867ec;"" valign=""top"" align=""center"" bgcolor=""#0867ec""> <a href=""http://htmlemail.io"" target=""_blank"" style=""border: solid 2px #0867ec; border-radius: 4px; box-sizing: border-box; cursor: pointer; display: inline-block; font-size: 16px; font-weight: bold; margin: 0; padding: 12px 24px; text-decoration: none; text-transform: capitalize; background-color: #0867ec; border-color: #0867ec; color: #ffffff;"">Call To Action</a> </td>"
    'sHeader = sHeader & "                            </tr>"
    'sHeader = sHeader & "                            </tbody>"
    'sHeader = sHeader & "                        </table>"
    'sHeader = sHeader & "                        </td>"
    'sHeader = sHeader & "                    </tr>"
    'sHeader = sHeader & "                    </tbody>"
    'sHeader = sHeader & "                </table>"
    sHeader = sHeader & "                <p style=""font-family: Helvetica, sans-serif; font-size: 16px; font-weight: normal; margin: 0; margin-bottom: 16px;"">This is a really simple email template. It's sole purpose is to get the recipient to click the button with no distractions.</p>"
    sHeader = sHeader & "                <p style=""font-family: Helvetica, sans-serif; font-size: 16px; font-weight: normal; margin: 0; margin-bottom: 16px;"">Good luck! Hope it works.</p>"
    sHeader = sHeader & "                </td>"
    sHeader = sHeader & "            </tr>"

    sHeader = sHeader & "            <!-- END MAIN CONTENT AREA -->"
    sHeader = sHeader & "            </table>"

    sHeader = sHeader & "            <!-- START FOOTER -->"
    sHeader = sHeader & "            <div class=""footer"" style=""clear: both; padding-top: 24px; text-align: center; width: 100%;"">"
    sHeader = sHeader & "            <table role=""presentation"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: separate; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%;"" width=""100%"">"
    sHeader = sHeader & "                <tr>"
    sHeader = sHeader & "                <td class=""content-block"" style=""font-family: Helvetica, sans-serif; vertical-align: top; color: #9a9ea6; font-size: 16px; text-align: center;"" valign=""top"" align=""center"">"
    sHeader = sHeader & "                    <span class=""apple-link"" style=""color: #9a9ea6; font-size: 16px; text-align: center;"">Company Inc, 7-11 Commercial Ct, Belfast BT1 2NB</span>"
    sHeader = sHeader & "                    <br> Don't like these emails? <a href=""http://htmlemail.io/blog"" style=""text-decoration: underline; color: #9a9ea6; font-size: 16px; text-align: center;"">Unsubscribe</a>."
    sHeader = sHeader & "                </td>"
    sHeader = sHeader & "                </tr>"
    sHeader = sHeader & "                <tr>"
    sHeader = sHeader & "                <td class=""content-block powered-by"" style=""font-family: Helvetica, sans-serif; vertical-align: top; color: #9a9ea6; font-size: 16px; text-align: center;"" valign=""top"" align=""center"">"
    sHeader = sHeader & "                    Powered by <a href=""http://htmlemail.io"" style=""color: #9a9ea6; font-size: 16px; text-align: center; text-decoration: none;"">HTMLemail.io</a>"
    sHeader = sHeader & "                </td>"
    sHeader = sHeader & "                </tr>"
    sHeader = sHeader & "            </table>"
    sHeader = sHeader & "            </div>"

    sHeader = sHeader & "            <!-- END FOOTER -->"
                
    sHeader = sHeader & "<!-- END CENTERED WHITE CONTAINER --></div>"
    sHeader = sHeader & "        </td>"
    sHeader = sHeader & "        <td style=""font-family: Helvetica, sans-serif; font-size: 16px; vertical-align: top;"" valign=""top"">&nbsp;</td>"
    sHeader = sHeader & "    </tr>"
    sHeader = sHeader & "    </table>"
    sHeader = sHeader & "</body>"
    sHeader = sHeader & "</html>"
    
    HtmlHeader = sHeader
End Function



Private Function HtmlHeader_Old3() As String
    Dim sHeader As String

    sHeader = "<!doctype html>"
    sHeader = sHeader & "<!-- Complete Email template -->"
    
    sHeader = sHeader & "<body style=""background-color:grey"">"
    sHeader = sHeader & "    <table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"""
    sHeader = sHeader & "        width=""550"" bgcolor=""white"" style=""border:2px solid black"">"
    sHeader = sHeader & "        <tbody>"
    sHeader = sHeader & "            <tr>"
    sHeader = sHeader & "                <td align=""center"">"
    sHeader = sHeader & "                    <table align=""center"" border=""0"" cellpadding=""0"""
    sHeader = sHeader & "                        cellspacing=""0"" class=""col-550"" width=""550"">"
    sHeader = sHeader & "                        <tbody>"
    sHeader = sHeader & "                            <tr>"
    sHeader = sHeader & "                                <td align=""center"" style=""background-color: #4cb96b;"
    sHeader = sHeader & "                                        height: 50px;"">"
    
    sHeader = sHeader & "                                    <a href=""#"" style=""text-decoration: none;"">"
    sHeader = sHeader & "                                        <p style=""color:white;"
    sHeader = sHeader & "                                                font-weight:bold;"">"
    sHeader = sHeader & "                                            GeeksforGeeks"
    sHeader = sHeader & "                                        </p>"
    sHeader = sHeader & "                                    </a>"
    sHeader = sHeader & "                                </td>"
    sHeader = sHeader & "                            </tr>"
    sHeader = sHeader & "                        </tbody>"
    sHeader = sHeader & "                    </table>"
    sHeader = sHeader & "                </td>"
    sHeader = sHeader & "            </tr>"
    sHeader = sHeader & "            <tr style=""height: 300px;"">"
    sHeader = sHeader & "                <td align=""center"" style=""border: none;"
    sHeader = sHeader & "                        border-bottom: 2px solid #4cb96b; "
    sHeader = sHeader & "                        padding-right: 20px;padding-left:20px"">"

    sHeader = sHeader & "                   <p style=""font-weight: bolder;font-size: 42px;"
    sHeader = sHeader & "                            letter-spacing: 0.025em;"
    sHeader = sHeader & "                            color:black;"">"
    sHeader = sHeader & "                        Hello Geeks!"
    sHeader = sHeader & "                        <br> Check out our latest Blogs"
    sHeader = sHeader & "                    </p>"
    sHeader = sHeader & "               </td>"
    sHeader = sHeader & "            </tr>"

    sHeader = sHeader & "            <tr style=""display: inline-block;"">"
    sHeader = sHeader & "                <td style=""height: 150px;"
    sHeader = sHeader & "                        padding: 20px;"
    sHeader = sHeader & "                        border: none; "
    sHeader = sHeader & "                        border-bottom: 2px solid #361B0E;"
    sHeader = sHeader & "                        background-color: white;"">"
              
    sHeader = sHeader & "                    <h2 style=""text-align: left;"
    sHeader = sHeader & "                            align-items: center;"">"
    sHeader = sHeader & "                        Design Patterns : A Must Skill to "
    sHeader = sHeader & "                    have for Software Developers in 2019"
    sHeader = sHeader & "                </h2>"
    sHeader = sHeader & "                    <p class=""data"""
    sHeader = sHeader & "                    style=""text-align: justify-all;"
    sHeader = sHeader & "                            align-items: center;"
    sHeader = sHeader & "                            font-size: 15px;"
    sHeader = sHeader & "                            padding-bottom: 12px;"">"
    sHeader = sHeader & "                        Design Patterns….??? I think you have heard this name "
    sHeader = sHeader & "                    before in programming… Yes, you might have heard "
    sHeader = sHeader & "                    this name before in programming if you are…"
    sHeader = sHeader & "                    </p>"
    sHeader = sHeader & "                    <p>"
    sHeader = sHeader & "                        <a href=""https://www.geeksforgeeks.org/design-patterns-a-must-skill-to-have-for-software-developers-in-2019/"""

    sHeader = sHeader & "                        style=""text-decoration: none; "
    sHeader = sHeader & "                                color:black; "
    sHeader = sHeader & "                                border: 2px solid #4cb96b; "
    sHeader = sHeader & "                                padding: 10px 30px;"
    sHeader = sHeader & "                                font-weight: bold;""> "
    sHeader = sHeader & "                        Read More "
    sHeader = sHeader & "                    </a>"
    sHeader = sHeader & "                    </p>"
    sHeader = sHeader & "                </td>"
    sHeader = sHeader & "            </tr>"
    sHeader = sHeader & "            <tr style=""border: none; "
    sHeader = sHeader & "            background-color: #4cb96b; "
    sHeader = sHeader & "            height: 40px; "
    sHeader = sHeader & "            color:white; "
    sHeader = sHeader & "            padding-bottom: 20px; "
    sHeader = sHeader & "            text-align: center;"">"
                    
    sHeader = sHeader & "<td height=""40px"" align=""center"">"
    sHeader = sHeader & "    <p style=""color:white; "
    sHeader = sHeader & "    line-height: 1.5em;"">"
    sHeader = sHeader & "    GeeksforGeeks"
    sHeader = sHeader & "    </p>"
    sHeader = sHeader & "    <a href=""#"" "
    sHeader = sHeader & "    style=""border:none;"
    sHeader = sHeader & "        text-decoration: none; "
    sHeader = sHeader & "        padding: 5px;""> "
       
    sHeader = sHeader & "    <img height=""30"" "
    sHeader = sHeader & "    src=""https://extraaedgeresources.blob.core.windows.net/demo/salesdemo/EmailAttachments/icon-twitter_20190610074030.png"""
    sHeader = sHeader & "    width=""30"" /> "
    sHeader = sHeader & "    </a> "
        
    sHeader = sHeader & "    <a href=""#"""
    sHeader = sHeader & "    style=""border:none;"
    sHeader = sHeader & "    text-decoration: none; "
    sHeader = sHeader & "    padding: 5px;""> "
        
    sHeader = sHeader & "    <img height=""30"" "
    sHeader = sHeader & "    src="
    sHeader = sHeader & """https://extraaedgeresources.blob.core.windows.net/demo/salesdemo/EmailAttachments/icon-linkedin_20190610074015.png"""
    sHeader = sHeader & "width=""30"" /> "
    sHeader = sHeader & "    </a>"
        
    sHeader = sHeader & "    <a href=""#"" "
    sHeader = sHeader & "    style=""border:none;"
    sHeader = sHeader & "    text-decoration: none;"
    sHeader = sHeader & "    padding: 5px;""> "
        
    sHeader = sHeader & "    <img height=""20"""
    sHeader = sHeader & "    src="
    sHeader = sHeader & """https://extraaedgeresources.blob.core.windows.net/demo/salesdemo/EmailAttachments/facebook-letter-logo_20190610100050.png"""
    sHeader = sHeader & "        width=""24"" "
    sHeader = sHeader & "        style=""position: relative; "
    sHeader = sHeader & "            padding-bottom: 5px;"" />"
    sHeader = sHeader & "    </a>"
    sHeader = sHeader & "</td>"
    sHeader = sHeader & "</tr>"
    sHeader = sHeader & "<tr>"
    sHeader = sHeader & "<td style=""font-family:'Open Sans', Arial, sans-serif;"
    sHeader = sHeader & "        font-size:11px; line-height:18px; "
    sHeader = sHeader & "        color:#999999;"" "
    sHeader = sHeader & "    valign=""top"""
    sHeader = sHeader & "    align=""center"">"
    sHeader = sHeader & "<a href=""#"" "
    sHeader = sHeader & "target=""_blank"" "
    sHeader = sHeader & "style=""color:#999999; "
    sHeader = sHeader & "        text-decoration:underline;"">PRIVACY STATEMENT</a> "
    sHeader = sHeader & "        | <a href=""#"" target=""_blank"" "
    sHeader = sHeader & "        style=""color:#999999; text-decoration:underline;"">TERMS OF SERVICE</a> "
    sHeader = sHeader & "        | <a href=""#"""
    sHeader = sHeader & "        target=""_blank"" "
    sHeader = sHeader & "        style=""color:#999999; text-decoration:underline;"">RETURNS</a><br>"
    sHeader = sHeader & "                © 2012 GeeksforGeeks. All Rights Reserved.<br>"
    sHeader = sHeader & "                If you do not wish to receive any further "
    sHeader = sHeader & "                emails from us, please"
    sHeader = sHeader & "                <a href=""#"""
    sHeader = sHeader & "                target=""_blank"""
    sHeader = sHeader & "                style=""text-decoration:none; "
    sHeader = sHeader & "                        color:#999999;"">unsubscribe</a>"
    sHeader = sHeader & "            </td>"
    sHeader = sHeader & "            </tr>"
    sHeader = sHeader & "            </tbody></table></td>"
    sHeader = sHeader & "        </tr>"
    sHeader = sHeader & "        <tr>"
    sHeader = sHeader & "        <td class=""em_hide"""
    sHeader = sHeader & "        style=""line-height:1px;"
    sHeader = sHeader & "                min-width:700px;"
    sHeader = sHeader & "                background-color:#ffffff;"">"
    sHeader = sHeader & "            <img alt="""" "
    sHeader = sHeader & "            src=""images/spacer.gif"" "
    sHeader = sHeader & "            style=""max-height:1px; "
    sHeader = sHeader & "            min-height:1px; "
    sHeader = sHeader & "            display:block; "
    sHeader = sHeader & "            width:700px; "
    sHeader = sHeader & "            min-width:700px;"" "
    sHeader = sHeader & "            width=""700"""
    sHeader = sHeader & "            border=""0"" "
    sHeader = sHeader & "            height=""1"">"
    sHeader = sHeader & "            </td>"
    sHeader = sHeader & "        </tr>"
    sHeader = sHeader & "        </tbody>"
    sHeader = sHeader & "    </table>"
    sHeader = sHeader & "</body>"
    sHeader = sHeader & "</html>"

End Function

Private Function HtmlHeader_FULL() As String
    Dim sHeader As String

    sHeader = "<!DOCTYPE html>"
    sHeader = sHeader & "<html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">"
    sHeader = sHeader & "<head>"
    sHeader = sHeader & "    <meta charset=""UTF-8"">"
    sHeader = sHeader & "    <meta name=""viewport"" content=""width=device-width,initial-scale=1"">"
    sHeader = sHeader & "    <meta name=""x-apple-disable-message-reformatting"">"
    sHeader = sHeader & "    <title></title>"
    sHeader = sHeader & "    <!--[if mso]>"
    sHeader = sHeader & "    <noscript>"
    sHeader = sHeader & "    <xml>"
    sHeader = sHeader & "        <o:OfficeDocumentSettings>"
    sHeader = sHeader & "            <o:PixelsPerInch>96</o:PixelsPerInch>"
    sHeader = sHeader & "        </o:OfficeDocumentSettings>"
    sHeader = sHeader & "    </xml>"
    sHeader = sHeader & "    </noscript>"
    sHeader = sHeader & "    <![endif]-->"

    sHeader = sHeader & "</head>"
    sHeader = sHeader & "<body style=""margin:0;padding:0;"">"
    sHeader = sHeader & "    <table role=""presentation"" style=""width: 100%;border-collapse: collapse;border: 0;border-spacing: 0;background: #ffffff;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "        <tr>"
    sHeader = sHeader & "            <td align=""center"" style=""padding: 0;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                <table role=""presentation"" style=""width: 602px;border-collapse: collapse;border: 1px solid #cccccc;border-spacing: 0;text-align: left;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                    <tr>"
    sHeader = sHeader & "                        <td align=""center"" style=""padding: 40px 0 30px 0;background: #ffffff;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                            <img alt=""logo"" src=""data:image/jpg;base64," & ConvertFileToBase64(GetAppLogo) & """ style=""width: 162px; border-radius: 4px;"">"
    sHeader = sHeader & "                        </td>"
    sHeader = sHeader & "                    </tr>"
    sHeader = sHeader & "                    <tr>"
    sHeader = sHeader & "                        <td style=""padding: 36px 30px 42px 30px;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                            <table role=""presentation"" style=""width: 100%;border-collapse: collapse;border: 0;border-spacing: 0;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                                <tr>"
    sHeader = sHeader & "                                    <td style=""padding: 0 0 36px 0;color: #153643;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                                        <h1 style=""font-size:24px;margin:0 0 20px 0;font-family:Arial,sans-serif;"">" & GetAppName & "</h1>"
    sHeader = sHeader & "                                        <p style=""margin:0 0 12px 0;font-size:14px;line-height:22px;font-family:Arial,sans-serif;"">"
    sHeader = sHeader & "                                                                Bravo ! Vous avez cree avec succes votre e-sim chez Moov-Africa."
    sHeader = sHeader & "                                        </p>"
    sHeader = sHeader & "                                    </td>"
    sHeader = sHeader & "                                </tr>"
    sHeader = sHeader & "                                <tr>"
    sHeader = sHeader & "                                    <td style=""padding: 0;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                                        <table role=""presentation"" style=""width: 100%;border-collapse: collapse;border: 0;border-spacing: 0;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                                            <tr>"
    sHeader = sHeader & "                                                <td style=""width: 260px;padding: 0;vertical-align: top;color: #153643;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                                                    <p style=""margin:0 0 12px 0;font-size:10px;line-height:20px;font-family:Arial,sans-serif;"">"
    sHeader = sHeader & "                                                        <dl>"
    sHeader = sHeader & "                                                            <dt><small> <strong>Nom/Raison Sociale:</strong> </small></dt>"
    sHeader = sHeader & "                                                            <dd> <small>client->nom </small></dd>"
                                                                
    sHeader = sHeader & "                                                            <dt><small> <strong>Numero Telephone:</strong> </small></dt>"
    sHeader = sHeader & "                                                            <dd><small>client->telephone</small></dd>"
                                                                
    sHeader = sHeader & "                                                            <dt><small> <strong>Code PIN:</strong> </small></dt>"
    sHeader = sHeader & "                                                            <dd><small>client->pin</small></dd>"
                                                                
    sHeader = sHeader & "                                                            <dt><small> <strong>Code PUK:</strong> </small></dt>"
    sHeader = sHeader & "                                                            <dd><small>client->puk</small></dd>"
    sHeader = sHeader & "                                                            <dt><small> <strong>Code Activation:</strong> </small></dt>"
    sHeader = sHeader & "                                                            <dd><small>$client->ac</small></dd>"
    sHeader = sHeader & "                                                        </dl>"
    sHeader = sHeader & "                                                    </p>"
    sHeader = sHeader & "                                                </td>"
                                                                                                                                            
    sHeader = sHeader & "                                                <td style=""width: 20px;padding: 0;font-size: 0;line-height: 0;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                                                    &nbsp;"
    sHeader = sHeader & "                                                </td>"
    sHeader = sHeader & "                                                <td style=""width: 260px;padding: 0;vertical-align: top;color: #153643;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                                                    <p style=""margin:0 0 12px 0;font-size:10px;line-height:20px;font-family:Arial,sans-serif;"">"
    sHeader = sHeader & "                                                        Scannez ce QR code pour telecharger votre profile E-SIM."
    sHeader = sHeader & "                                                    </p>"
    sHeader = sHeader & "                                                    <p style=""margin:0 0 25px 0;font-size:10px;line-height:20px;font-family:Arial,sans-serif; text-align:center;"">"
    'sHeader = sHeader & "                                                        <img src=""{{ $message-&gt;embed( $client-&gt;qrcodeimage_fullpath ) }}"" alt="""" width=""200"" style=""height:auto;display:block;"">"
    sHeader = sHeader & "                                                    </p>"
    sHeader = sHeader & "                                                </td>"
    sHeader = sHeader & "                                            </tr>"
    sHeader = sHeader & "                                        </table>"
    sHeader = sHeader & "                                    </td>"
    sHeader = sHeader & "                                </tr>"
    sHeader = sHeader & "                            </table>"
    sHeader = sHeader & "                        </td>"
    sHeader = sHeader & "                    </tr>"
    sHeader = sHeader & "                    <tr>"
    sHeader = sHeader & "                        <td style=""padding: 10px;background: #FF5733;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                            <table role=""presentation"" style=""width:100%;border-collapse:collapse;border:0;border-spacing:0;font-size:9px;font-family:Arial,sans-serif;"">"
    sHeader = sHeader & "                                <tr>"
    sHeader = sHeader & "                                    <td style=""padding: 0;width: 50%;font-family: Arial, sans-serif;"" align=""left"">"
    sHeader = sHeader & "                                        <p style=""margin:0;font-size:8px;line-height:16px;font-family:Arial,sans-serif;color:#ffffff;"">"
    sHeader = sHeader & "                                            &reg; MAGT/DSI - " & Year(Now)
    sHeader = sHeader & "                                        </p>"
    sHeader = sHeader & "                                    </td>"
    sHeader = sHeader & "                                    <td style=""padding: 0;width: 50%;font-family: Arial, sans-serif;"" align=""right"">"
    sHeader = sHeader & "                                        <table role=""presentation"" style=""border-collapse: collapse;border: 0;border-spacing: 0;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                                            <tr>"
    sHeader = sHeader & "                                                <td style=""padding: 0 0 0 10px;width: 38px;font-family: Arial, sans-serif;"">"
    'sHeader = sHeader & "                                                    <a href=""#"" style=""color:#ffffff;""><img src=""{{ $message-&gt;embed( public_path() . '/img/facebook.png' ) }}"" alt=""Twitter"" width=""38"" style=""height:auto;display:block;border:0;""></a>"
    sHeader = sHeader & "                                                </td>"
    sHeader = sHeader & "                                                <td style=""padding: 0 0 0 10px;width: 38px;font-family: Arial, sans-serif;"">"
    'sHeader = sHeader & "                                                    <a href=""#"" style=""color:#ffffff;""><img src=""{{ $message-&gt;embed( public_path() . '/img/twiter.png' ) }}"" alt=""Facebook"" width=""38"" style=""height:auto;display:block;border:0;""></a>"
    sHeader = sHeader & "                                                </td>"
    sHeader = sHeader & "                                            </tr>"
    sHeader = sHeader & "                                        </table>"
    sHeader = sHeader & "                                    </td>"
    sHeader = sHeader & "                                </tr>"
    sHeader = sHeader & "                            </table>"
    sHeader = sHeader & "                        </td>"
    sHeader = sHeader & "                    </tr>"
    sHeader = sHeader & "                </table>"
    sHeader = sHeader & "            </td>"
    sHeader = sHeader & "        </tr>"
    sHeader = sHeader & "    </table>"
    sHeader = sHeader & "</body>"
    sHeader = sHeader & "</html>"
                                                                                                
    HtmlHeader = sHeader
End Function

Private Function HtmlHeader() As String
    Dim sHeader As String

    sHeader = "<!DOCTYPE html>"
    sHeader = sHeader & "<html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">"
    sHeader = sHeader & "<head>"
    sHeader = sHeader & "    <meta charset=""UTF-8"">"
    sHeader = sHeader & "    <meta name=""viewport"" content=""width=device-width,initial-scale=1"">"
    sHeader = sHeader & "    <meta name=""x-apple-disable-message-reformatting"">"
    sHeader = sHeader & "    <title></title>"
    sHeader = sHeader & "    <!--[if mso]>"
    sHeader = sHeader & "    <noscript>"
    sHeader = sHeader & "    <xml>"
    sHeader = sHeader & "        <o:OfficeDocumentSettings>"
    sHeader = sHeader & "            <o:PixelsPerInch>96</o:PixelsPerInch>"
    sHeader = sHeader & "        </o:OfficeDocumentSettings>"
    sHeader = sHeader & "    </xml>"
    sHeader = sHeader & "    </noscript>"
    sHeader = sHeader & "    <![endif]-->"

    sHeader = sHeader & "</head>"
    sHeader = sHeader & "<body style=""margin:0;padding:0;"">"
    sHeader = sHeader & "    <table role=""presentation"" style=""width: 100%;border-collapse: collapse;border: 0;border-spacing: 0;background: #ffffff;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "        <tr>"
    sHeader = sHeader & "            <td align=""center"" style=""padding: 0;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                <table role=""presentation"" style=""width: 602px;border-collapse: collapse;border: 1px solid #cccccc;border-spacing: 0;text-align: left;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                    <tr>"
    sHeader = sHeader & "                        <td align=""center"" style=""padding: 40px 0 30px 0;background: #ffffff;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                            <img alt=""logo"" src=""data:image/jpg;base64, " & ConvertFileToBase64(GetAppLogo) & """ style=""width: 162px; height:auto;display:block;"">"
    sHeader = sHeader & "                        </td>"
    sHeader = sHeader & "                    </tr>"
    sHeader = sHeader & "                    <tr>"
    sHeader = sHeader & "                        <td style=""padding: 36px 30px 42px 30px;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                            <table role=""presentation"" style=""width: 100%;border-collapse: collapse;border: 0;border-spacing: 0;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                                <tr>"
    sHeader = sHeader & "                                    <td style=""padding: 0 0 36px 0;color: #153643;font-family: Arial, sans-serif;"">"
    sHeader = sHeader & "                                        <p style=""margin:0 0 12px 0;font-size:14px;line-height:22px;font-family:Arial,sans-serif;"">"
                                                                                                
    HtmlHeader = sHeader
End Function

Private Function HtmlFooter() As String
    Dim sFooter As String

    sFooter = "                                        </p>"
    sFooter = sFooter & "                                    </td>"
    sFooter = sFooter & "                                </tr>"
    sFooter = sFooter & "                                <tr>"
    sFooter = sFooter & "                                    <td style=""padding: 0;font-family: Arial, sans-serif;"">"
    sFooter = sFooter & "                                    </td>"
    sFooter = sFooter & "                                </tr>"
    sFooter = sFooter & "                                <tr>"
    sFooter = sFooter & "                                    <td style=""padding: 0;font-family: Arial, sans-serif;"">"
    sFooter = sFooter & "                                    </td>"
    sFooter = sFooter & "                                </tr>"
    sFooter = sFooter & "                                <tr>"
    sFooter = sFooter & "                                    <td style=""padding: 0;font-family: Arial, sans-serif;"">"
    sFooter = sFooter & "                                        <p style=""margin:0 0 12px 0;font-size:14px;line-height:22px;font-family:Arial,sans-serif;"">"
    sFooter = sFooter & "                                        Cordialement, Regards."
    sFooter = sFooter & "                                        <h1 style=""font-size:20px;margin:0 0 20px 0;font-family:Arial,sans-serif;"">" & GetAppName & "</h1>"
    sFooter = sFooter & "                                        </p>"
    sFooter = sFooter & "                                    </td>"
    sFooter = sFooter & "                                </tr>"
    sFooter = sFooter & "                            </table>"
    sFooter = sFooter & "                        </td>"
    sFooter = sFooter & "                    </tr>"
    sFooter = sFooter & "                    <tr>"
    sFooter = sFooter & "                        <td style=""padding: 10px;background: #FF5733;font-family: Arial, sans-serif;"">"
    sFooter = sFooter & "                            <table role=""presentation"" style=""width:100%;border-collapse:collapse;border:0;border-spacing:0;font-size:9px;font-family:Arial,sans-serif;"">"
    sFooter = sFooter & "                                <tr>"
    sFooter = sFooter & "                                    <td style=""padding: 0;width: 50%;font-family: Arial, sans-serif;"" align=""left"">"
    sFooter = sFooter & "                                        <p style=""margin:0;font-size:8px;line-height:16px;font-family:Arial,sans-serif;color:#ffffff;"">"
    sFooter = sFooter & "                                            &reg; MAGT/DSI - 2024"
    sFooter = sFooter & "                                        </p>"
    sFooter = sFooter & "                                    </td>"
    sFooter = sFooter & "                                    <td style=""padding: 0;width: 50%;font-family: Arial, sans-serif;"" align=""right"">"
    sFooter = sFooter & "                                        <table role=""presentation"" style=""border-collapse: collapse;border: 0;border-spacing: 0;font-family: Arial, sans-serif;"">"
    sFooter = sFooter & "                                            <tr>"
    sFooter = sFooter & "                                                <td style=""padding: 0 0 0 10px;width: 38px;font-family: Arial, sans-serif;"">"
    'sFooter = sFooter & "                                                    <a href=""#"" style=""color:#ffffff;""><img src=""{{ $message-&gt;embed( public_path() . '/img/facebook.png' ) }}"" alt=""Twitter"" width=""38"" style=""height:auto;display:block;border:0;""></a>"
    sFooter = sFooter & "                                                </td>"
    sFooter = sFooter & "                                                <td style=""padding: 0 0 0 10px;width: 38px;font-family: Arial, sans-serif;"">"
    'sFooter = sFooter & "                                                    <a href=""#"" style=""color:#ffffff;""><img src=""{{ $message-&gt;embed( public_path() . '/img/twiter.png' ) }}"" alt=""Facebook"" width=""38"" style=""height:auto;display:block;border:0;""></a>"
    sFooter = sFooter & "                                                </td>"
    sFooter = sFooter & "                                            </tr>"
    sFooter = sFooter & "                                        </table>"
    sFooter = sFooter & "                                    </td>"
    sFooter = sFooter & "                                </tr>"
    sFooter = sFooter & "                            </table>"
    sFooter = sFooter & "                        </td>"
    sFooter = sFooter & "                    </tr>"
    sFooter = sFooter & "                </table>"
    sFooter = sFooter & "            </td>"
    sFooter = sFooter & "        </tr>"
    sFooter = sFooter & "    </table>"
    sFooter = sFooter & "</body>"
    sFooter = sFooter & "</html>"
                                                                                                
    HtmlFooter = sFooter
End Function

Private Function DefaultLogoBase64() As String
    Dim finalFile As String

    finalFile = "/9j/4AAQSkZJRgABAgEB9AH0AAD//gASTEVBRFRPT0xTIHYyMC4wAP/bAIQABQUFCAUIDAcHDAwJCQkMDQwMDAwNDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQEFCAgKBwoMBwc"
    finalFile = finalFile & "MDQwKDA0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0N/8QBogAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoLAQADAQEBAQEBAQEBAAAAAAA"
    finalFile = finalFile & "AAQIDBAUGBwgJCgsQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGV"
    finalFile = finalFile & "mZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+hEAAgECBAQDBAcFBAQAAQJ3AAE"
    finalFile = finalFile & "CAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJm"
    finalFile = finalFile & "aoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/8AAEQgAPwCiAwERAAIRAQMRAf/aAAwDAQACEQMRAD8A+tNb1u08PWj31+4ihjGST1J"
    finalFile = finalFile & "7Ko6knsBWtOnKrJQgrtkSkoLmlseURaj4p8enzdO/4kumE/JK4P2iRf7yjoAe3SvTcaGF0qfvKnVL4UcN6tbWHuR79WaA+DunXZ36tPPqMh6tLjr+FZ/XZx0pRUF5FfVYvWbcn5kv/Cl"
    finalFile = finalFile & "PC3/PqPzpfXq/8xf1an2/ET/hSvhYc/ZR+dH16v8AzB9Wp9vxEHwX8KngWyn6NT+vV/5g+rU+34jv+FKeF/8An1H5ml9er/zB9Wp9g/4Up4X/AOfUfmaPr1f+YPq1PsH/AApTwv8A8+o"
    finalFile = finalFile & "/Oj69X/mD6tT7B/wpTwv/AM+o/Oj69X/mD6tT7B/wpTwv/wA+o/Oj69X/AJg+rU+wf8KU8L/8+o/M0fXq/wDMH1an2D/hSnhf/n1H50fXq/8AMH1an2D/AIUp4X/59R+dH16v/MH1an2"
    finalFile = finalFile & "/Ej/4Ux4U6fZ1/wC+qf12v3/AX1en/TJP+FKeF/8An1H50vr1f+Yf1an2D/hSnhf/AJ9R+Zo+vV/5g+rU+wh+CvhYcm1AH1o+vV/5g+rU+34jP+FMeFP+fdP++qf12v3/AAF9Xp/0zvf"
    finalFile = finalFile & "Dnhyw8MWv2PTEEcGc4BzzXFVqyrS5qj1OmEFBWib9YGh4i1sfiD4mkS5+bSdDYKI/4JbnrlvUAdvavYv9VoJx/iVOvVRPN1r1Wn8EPxZ33jHxfZeB9PN5c44+WKJcAu2OFA9K4aFGWInyx+b7HXUqRoxu/kj5F1"
    finalFile = finalFile & "v41+I9VlJt5PsUTfdSIZOO2e+fWvqaeAowXvLmfdngTxdSb9x29B+geIPEupCS/v7+4t9NteZpGGC56iNM4yzUqlKjC0IQi5vZdvN+RVOdSScpyait3+i8yn4j+Leua23lWsz2tqvyoiffcDgM3fcepxV0sFSp6"
    finalFile = finalFile & "yScu76ehFTFTk+WL0Ry0PivxBpcgmW6uYmzkFt2D+fFdLo0pq3LFmCrVIaps+l/hb8YT4glXSNa2x3ZH7qXosvsf9o9q+exeC9kva0vh6rt/wAA9nD4r2nuT36P/M9Y8Y+MbHwXZNeXjDfgiKIH5nbsAOuM9TXm"
    finalFile = finalFile & "UKEsRLkht1fY7qtWNGN5fJHL/CfxVeeL9Ol1G+IDNM4RR0RAeB+FdGMoxoTUIdl82ZYebqRcn3ZzXxK+LsehSLpOjMs168ipI45WIE4Iz/erowuDdRe1q6Rtou5lWxCg+SD17jPiB8WB4W0+GysmE2qzxIWPaLcOWb39KeHwf"
    finalFile = finalFile & "tpuc1amm/mTWxHs4qKfvW18v+CXvEfxSj8JaBbSSMLjVbm2VlTuGZfvv6DJ4+lRTwjrVZJK1OMv6SKliPZ043d5tL8ep2fgnXZLrw9Dq2pyDc0Zllc8Ad65K9NRrOlTWzskdFKd6aqSfS580+OvjXqOszvb6M5tLJCVDr9+"
    finalFile = finalFile & "QevsD2xX0OHwEKaUqq5pduiPHrYuTfLT0X9bnlo1/W4z9p+0XQ77/nx/hXo+zpfDyx9NDh9pUvfX7z2j4c/Gu7s500/xA/n28hCpcH7yE8Dd/s15OJwEWnOgrNdOjPRoYt35am34o+nNf8S2XhzT31S7cCFV3Lg8vkZAX1J"
    finalFile = finalFile & "r52nSlVmqcVr+R7M6kaced7fmfF/i34v614kkYQSGytM/LHH1I7Fj1zX1tHBU6S95c0u7PnauLnN2jouyK3gnTfEnjS8FraXVxHApBmmJYKi98Z7ntVV5UcPHmlGN+i7joxqVZW1+/Y+4tC0hdDs47JHaXyxgu5yzHuTXx1"
    finalFile = finalFile & "SftJOdrX6I+ihHkSia9ZGh5H8GEMuhi+fmS9kMrn1OMV6mOdqvItoqyODCq0Lvdu7Pnz46a3LqPiBrNifKsl2qvbcec49a93L6ahS5+stTycZNym49tD1D4KfD+wbTl1y+jWee4JMe8ZCKOOB05NebjsTPndGDslvbqd2Eo"
    finalFile = finalFile & "RUVUau3sc7+0LqH2eS00i3AigKmZ1QBQWBwMgYzgV0ZbG6lVlq9lcxxsuVqEdFa/wCho/ArwNZXdq2u3qLNIzbYVYZVAOpx65rPMMRKMvYwdlbXzLwdCNueSu+h77rvhTTfENq9ndwRlHBAIUBlOOCCAOnWvEp1p0pKcGz1"
    finalFile = finalFile & "J041FyySPz71K1k8Mau8MbfPYz5jbvhW+X/69fcRarU03tJa/M+TlejN26P8mSeJvFF94suheai5dwoVVGdqgdwOmT3NFKjChHlgrBUqyras6jRviPdeHtAbQ9OHlTTOxefuFbsvuRXNPCxq1VWnqktEdEcS6dP2cdG27s4"
    finalFile = finalFile & "GyJa7hdiWZp0LMTkklhyTXbLSLS7M5YNymm/61Ot+IAxrp/3Iq5cN/C+bOqv8f3EPj/m9izk4s48Z7cdvSnh/hf8AiZNbeP8A27+p7L4n1SXTvh1ZQxEr9rARiOOBzj8a8qlBSxk2/s6no1JcuHgu5xXwW8GW3irU3lvlD2"
    finalFile = finalFile & "9kofyz0dicDPsDzXXjq8qEEoaOWl+xyYSkqkm59NfU+z5dFsZoDavBEYSu3bsXGOnpXySqST5lJ39T6Pki1ayt6Hwr8V/CcPhHWntrXi3nXzo1/uZPT8DX2eDrOvT5pbrR+Z8xiaao1Pd9V6Mn8YeJ7jVtB0qwkYlYY8tz9"
    finalFile = finalFile & "5hwCfXilRoqFWpNdX9w6tVunCHZHYfBLwBZ+JWl1XUVE0Vs4RIj0LYzuP0rkx+IlRtTp6Nq7fkdGDoRnectbdPPc+j9f8RaL8PLMyzCO3GMpFGAHkI6AAdfqa8CnSqYqVld929kexOcKC1svJbsj+H3i6Txrpx1J4xCrOVR"
    finalFile = finalFile & "B1CgnGfeniaKw8/Zp301HSqOonJ9zu64joPI/g2xttJk0t+JdPlMLjuDjNepjtaiqLaaujgwrfI4PeLsz58+Oehzabr7XrA+TeruVu24cYz617uX1FKkodYnk4ym4zcu+p6P8GfiRp1npq6Lqcq20luT5bOcK6nnr65rz8d"
    finalFile = finalFile & "hZyn7Wmrp726HZhcRBQ5JNK2xg/Hp9P1YW2qWNzFM0WYnRGy2GOQ30Fb5dz0+anOLV9U2ZY3lnacZJ9LfiZ3we+Jtt4XRtJ1QlLZ23RS9kPcN7GrxuElWftaXxLddyMLiVBck9vyPcfEfxg0HRrVpra4S7mKny44jnLY43e"
    finalFile = finalFile & "gz1rx6WCq1JWlFxXVs9OeKpwV4tN+R8UStd+JtRZo0Mt1eykhF5xuPA+gr69ctCGrtGKPmrOtPTqz6Sf4Y23g3wpdXN0om1GWLLOf+WeR91fp3r5/61KvXjGOkE/vPb+rqjSba9633HiHw68ET+N70QrlLSEbp5ccAf3QfU"
    finalFile = finalFile & "17GJxCw0L/aeyPMo0HWm+iX4GBc2yWWsG2h/wBXDeBFz1wr4FbJuVPme7jf8DHlUall3/U3viD/AMhz/gEVY4b+F82bV/j+4g8fc3sX/XpH/Kqw/wAL/wATJrbx9I/qes63dabqvgC2tVuYheWiqwi3fNnPK49cV5dOM4Yu"
    finalFile = finalFile & "UuV8suvQ75ShLDxXMrx6Hnfww8cjwNqRmmBa0uVCTAdVAPDAexrvxeH+swsviWqOLDVvYy12ejPreb4r+G4bX7X9sjbjIjB+fPpj1r5dYOs5cvI/Xoe+8TSSvzL06nxp498WSeNtWe+ClUP7qBP4tucDI9Sa+sw9FYamodd"
    finalFile = finalFile & "36nz1ar7ed1t+h6B498Bz6P4Y02+CHzIUAuQB90NyCfauLD4hTr1IX0b90661BxpQlbW1mZXwn+IyeCJ5ILwM1lckMSvJR+zY9Mda1xmFeISlD4l+KMsLiPYtxls9yL4x+J7LxXqkNzp8nmwxw7T6Bvp60YGjKhBxmrO4Yu"
    finalFile = finalFile & "rGpJcj6H0D8B/+RbT/AH2rw8w/jfI9fCfB8z2qvIPQPEdbaX4d6+2uBWfSdUIW72jPkzZ4lIH8OOtexC2Kpex2qQ+HzXY86SdCp7RfBL4vJ9zv9f8AD+l+P9N8mbbNBKN0UqYJU9mU/wBK4adSeFndaNbpnVOEa0bPboz5Y"
    finalFile = finalFile & "174Da3p0hNgUu4c/Kc4kx2yK+kp5hTkvfvF/geJPBTi7w1Xl/kdB8Ofgrcy3T3PiWPFvGNqQlj87H+I+wrDE45KKjh37z69jajhLu9VadvMr+LfgHf2szzaG6z27EkRPwyZ7D1AqqOYxaSrKz7rYmrgmnent9zOIsPg14lv"
    finalFile = finalFile & "JREbdYBnBd+APeuyWOoxV+a/kjlWEqSdmmvU+n/h38LLPwQn2iQi4v3HzSkfc9VT0Hv1r53E4uWI91aQ7d/U9uhh40Vf7X5HV+NtGn8QaTNp1sQsk4C5PYHqfwrloTVKopy2RvVi5xcY7sZ4Q8JW3g7Sk021UblT94/d5CO"
    finalFile = finalFile & "ST9elOtWlXm5y76LshU6apR5Vv19T5Kvvhb4kl1eS6S2/dNd+YD/sb85/Kvp44uiqajza8tvwPFeGnz81nv8Arc0/GXwz8Q6lq32m2t98W2MZ9x1rOhiqUKfLKVnqVVw85Sukxni74Y+Ib+7ikt7feq20aE+jDqKdHF0YRa"
    finalFile = finalFile & "cre82TVw05W0ey2OVPwi8TjkWpz9a6frtH+Y5/qs9rOx39z8A72fTYLqzcJfFMzwScLu77T61wrMYqcoyXuX0aOp4JuKcfitqnocPF8GfE0kvlfZ1XnG4n5frmux46ilfm+Rz/AFSpe2v3nungD4HwaDMmoay4ubmM5SMcx"
    finalFile = finalFile & "ofX3NeNice6icKWkXu+rPSoYRU/enq+3Q92v7CDUrd7S6QSQyqVZT0INeNGTg1KLs0em4qS5XsfKfij9n68t5Xl0OVZYGJIikOCuew9QK+lo5jFpKsrPujxKuCd70/8meet8IPE6EqLUfUdK7vrtH+Y4/qlS+zPqv4SeHr3"
    finalFile = finalFile & "wzoi2eoqI5gxOB6GvmsZUjVqc0NrHvYeDpxs+56fXnHWV7u0hv4mt7hFlikBVlYAgg+xqotxalF2aE0mrPY8il+HmqeG5Gn8JXhhic5NnN80Xr8rHJUewr1FiYVVy4qF2vtLc4HRnT1oSt/dew8+LfF2nfu7zSo5iON8Dsw"
    finalFile = finalFile & "Pvij2OHlrCq15NC9rWjpKmn6MT/hP/EX/AEBZfzNH1aj/AM/kL29X/n0w/wCE+8RD/mCyfmaPq1H/AJ/Ift6v/Pth/wAJ94iH/MGk/M0fVqP/AD+Qe3q/8+2H/CfeIv8AoCyfmaPq1H/n8g9vV/59sP8AhPvEX/QGk/M0fV"
    finalFile = finalFile & "qP/P5B7er/AM+2H/CfeIv+gLJ+Zo+rUf8An8g9vV/59sP+E+8Rf9AaX8zR9Wo/8/kHt6v/AD7Yf8J94i/6A0v5mj6tR/5/IPb1f+fbD/hPvEX/AEBpfzNH1aj/AM/kHt6v/Pth/wAJ94i/6A0v5mj6tR/5/IPb1f8An2w/4"
    finalFile = finalFile & "T7xF/0BpfzNH1aj/wA/kL29X/n2xf8AhPvEX/QGk/M/40fVqP8Az+Q/b1f+fbE/4T7xF/0BZPzNH1aj/wA/kHt6v/Pth/wn3iL/AKAsn5mj6tR/5/IPb1f+fbD/AIT7xF/0BpPzNH1aj/z+Qe3q/wDPth/wn/iL/oCy/maP"
    finalFile = finalFile & "q1H/AJ/IXt6v/PtnYaZ4s3WS3OsRfYJnP+pY5bHrXJOj7zjSfMl1OuNTS81yvsbQ8Q6aQD9oi5/2hWXsp/ysvnj3Rm+MoidOknWRoWhG5SvckgY/WtKD99Rte5FX4W07WOSxfaG9lDbu93c3GJZVYnbsxnn0FdXuVOdySjF"
    finalFile = finalFile & "aL1MPehypat6v0Nv/AITdTEqbFW8YkeWzYQBSQTu9OKx+r6t393v1NPa7L7XY29J8QpqVi16Rs8rcHHUZXPQ9xWM6bhJQ77GkZ80ebscZ4W8Q3kFneT6lIZpEmPkKeCQ2Sij9K661KLlCNNWVtf1OenOSjJzd3fQj0DxDqN"
    finalFile = finalFile & "lp4nvnE9xcSSOvmHYkaKfuZ9QOKdSlCU+WCtFJLTdvuFOclG8tW29+iN5fGMl0sK2cHmTTI7MCcImzrlvQ9vWsPYKN3OVkmvV3NPa3sorV3/Ali8WS3qxxWUHmXL7t6E4WPb1y3v8Aw+tJ0VG7nK0ej6saqXsorX8iKLxk8"
    finalFile = finalFile & "9oJ1hEcpkePbKdqAxnDEt6enrTdBKVr3Vk9N9QVW6vazvbXyGxeM5LtIUtYPMnnaRcZ+RfL6kt/dPY0Ogo3cnZK3rqJVb2UVq7/AICN44Xy1jCKt4zMpjdsICnU7vQ9qf1fW9/d7rfUPa9Ptdhr+ON6RLDGvnyFlIdiseVH"
    finalFile = finalFile & "IVu5Paj6vZu70Xbf7g9rsktfwJLnxoExDEiC4SPzZkkbaEX0B7t6CkqHVv3b2TXUbq9EtbXdzdbU5r7Svt1gmZZI9yKeuf8AH09ax5FCpyTeieppzOUOaG9tDgtL1JYpEupbqeK4iy11BMCDKOwRTwOemK7Zx0cVFOL+Frp"
    finalFile = finalFile & "6nLGVrNyaa3T6nR/8JdNasGvrcwRTLmE5yxPYMO2aw9inpCV2t/8AgG3tWviVk9jQvvFEdhHG8iNmRPMK9wOg/M1nGi5NpPZ2KdTlSbXmZ8Hiq8uQsUVr/pTqZBGzEARjux7MewrR0Yx1cvdWl/PyJVRvTl13t5DZvGm0JA"
    finalFile = finalFile & "sapeMpd4pW2hAp559T29aFQ3lf3dk11D2uyt73Z9Dnr6zuPExS/hTfHL9wH+Er8pU+ik85reLVG8G7Nb/P9TJp1PeS0Y//AIVhG/zM+GPJA6AnqB7Cj621okH1dHoGuaSdZgFvv8tdwLe4BziuGnP2b5ra9DqnHnVthv8AZ"
    finalFile = finalFile & "BF4LwP9yDyVXHT3o5/d5Ldbhy+9zeVjl38Fy20iT2sitKFdH8xQQQ7Z3DPcdMV0+3TTjJaXTVvIw9k07xeux0f9hbdNbTUfaZFKmQADknJOBXP7T31Utt0NuT3eT8THfwYpu1uhKQqQCMR4+XzAMCU+4rX2/u8tut7+XYz9"
    finalFile = finalFile & "l73Nfpa3n3KM3gq4XdHDONkqFTvUEIzdWVT6mrVdaNx1T6dSfZPZPf8AA1tL8LtYRSK0uZJIvLVgMbeMFgPfrWU6vM1poncuNPlT11tY4LVdRXw5dyWVnM0d1JEqH5MqzkYVt3b3ruhH2sVOSvFO++y7HLJ+zbjF2drfM37"
    finalFile = finalFile & "fwdeeRGwmBaSM+YGGVV3HLqDwTWDrxu9NE9PRdDVUnZa9NfmbmjeFn0xTvl3v5ZjUgYwSMFseprGdbn2Vle5rCny9eljPk8GSwSCa1kUu67JfMUH/AIEuehrRV01aS0W1vyI9k07xfrclufClyVWCKVDCVAYFAGVh1dD2J9"
    finalFile = finalFile & "qlVo7tO/r+DG6b2T09PxK58FTW8pFtKpilUK7SKGcYPJBPJzVe3TXvLVbW0QvZNPR6ee50mp6LJc2S2tnKbaSLaUdeOR2I9DXPCajLmkrp7o2lBuPLF2aOeuPB0+skz6nKvnqoERRR8pU5DE98+ldCrqn7tNadbmLpOes3r"
    finalFile = finalFile & "0JZvCt1qaM2oT7plXbDtHyoQchsdye9Sq0YO1OOnXz8hum5fE9enkIfCd3fOZr+43MYxEqqoCqobdn6mn7aMVaEetw9k27yfSxcvPD96l4bvT5xF5qLHIGGcYx8y5+nSojUjy8k43s7opwkpc0Xa6syldeEJ/tBnt5VJmXb"
    finalFile = finalFile & "K0ihmB/vLn+VWqytytbbWdiXSd7p773O0sbRbGBIE6IMZAxk9zj3rjk+ZuR0JcqsW6ko/9k="

    DefaultLogoBase64 = finalFile
End Function

'*****************************************************************************************************************************
'    Private events
'*****************************************************************************************************************************


Private Sub Class_Initialize()
    
End Sub

