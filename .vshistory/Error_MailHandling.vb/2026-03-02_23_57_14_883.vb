' ######################################################################
' ## Copyright (c) 2021 TimeShareIt GdbR
' ## by Thomas Steger
' ## File creation Date: 2021-8-6 07:29
' ## File update Date: 2021-8-23 12:15
' ## Filename: Mail.vb (F:\++++ Code Share\classes\Mail.vb)
' ## Project: ConDrop_Server
' ## Last User: stegert
' ######################################################################
'
'

Imports Microsoft.VisualBasic.ControlChars
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports System.DateTime
Imports DevComponents.DotNetBar
Imports Limilabs.Client.SMTP
Imports Limilabs.Mail
Imports Limilabs.Mail.Headers
Imports Limilabs.Mail.MIME
Imports log4net.Core
Imports Microsoft.Win32

Friend Class Error_MailHandling
    'Private Shared rb As Boolean = False
    Public Structure MailAttachement
        Public Name As String
        Public Path As String
        Public Size As Long
        Public Type As String
    End Structure
    Public Structure EmailTemplatePH
        Public Friendly As String
        Public Placeholder As String
    End Structure
    Public Structure MailTemplate
        Public TemplateName As String
        Public TemplateSubject As String
        Public TemplateText As String
        Public TemplateType As String
    End Structure

#Region "Email"
    Public Shared Function isEmail(ByVal email As String) As Boolean
        Dim emailRegex As New Regex("([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)", RegexOptions.IgnoreCase)
        Return emailRegex.IsMatch(email)
    End Function
    Public Shared Function isEmail(ByVal email As String, ByVal ReturnEmail As Boolean) As String
        Dim emailRegex As New Regex("([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)", RegexOptions.IgnoreCase)
        Dim ret As Match = emailRegex.Match(email)
        Return ret.Value
    End Function
    Public Shared Function SendSystemEmailUniversal(ByVal RegistryHiveValue As RegistryHive,
                                                    ByVal RegistryPath As String,
                                                    ByVal Header As String,
                                                    ByVal MessageText As String,
                                                    ByVal Optional MessageHtml As String = "",
                                                    ByVal Optional PathAttachment As List(Of String) = Nothing,
                                                    ByVal Optional SMTPCred As Error_Helper.SMTPCredentials = Nothing,
                                                    ByVal Optional Bitmaps As List(Of Bitmap) = Nothing) As Boolean
        Dim query As String = ""
        Dim ReturnVal As Boolean = False
        Dim builder As New MailBuilder()

        Try
            If SMTPCred.SMTP_SenderAddress Is Nothing Then
                SMTPCred.SMTP_Password = Error_cCrypt.DecryptString(Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_Password")))
                SMTPCred.SMTP_User = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_User"))
                SMTPCred.SMTP_Server = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_RelayServer"))
                SMTPCred.SMTP_NoSSL = Error_VarConvert.ConvertToBoolean(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_NoSSL"))
                SMTPCred.SMTP_SSL = Error_VarConvert.ConvertToBoolean(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_SSL"))
                SMTPCred.SMTP_STARTTLS = Error_VarConvert.ConvertToBoolean(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_STARTTLS"))
                SMTPCred.SMTP_SenderAddress = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_SenderAddress"))
                SMTPCred.SMTP_SenderName = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_SenderName"))
                SMTPCred.SMTP_RecipientAddress = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_RecipientAddress"))
                SMTPCred.SMTP_RecipientName = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_RecipientName"))
                SMTPCred.SMTP_CC = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_CC"))
                SMTPCred.SMTP_BCC = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_BCC"))
                SMTPCred.SMTP_NoSSL_Port = Error_VarConvert.ConvertToInteger(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_NoSSL_Port"))
                SMTPCred.SMTP_SSL_Port = Error_VarConvert.ConvertToInteger(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_SSL_Port"))
                SMTPCred.SMTP_STARTTLS_Port = Error_VarConvert.ConvertToInteger(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_STARTTLS_Port"))

            End If
            If SMTPCred.SMTP_SenderAddress <> "" Then
                builder.From.Add(New MailBox(SMTPCred.SMTP_SenderAddress, SMTPCred.SMTP_SenderName))
                builder.To.Add(New MailBox(SMTPCred.SMTP_RecipientAddress, SMTPCred.SMTP_RecipientName))
                If SMTPCred.SMTP_CC <> "" Then
                    builder.Cc.Add(New MailBox(SMTPCred.SMTP_CC))
                End If
                If SMTPCred.SMTP_BCC <> "" Then
                    builder.Bcc.Add(New MailBox(SMTPCred.SMTP_BCC))
                End If
                builder.Subject = Header
                builder.Text = MessageText
                If MessageHtml <> "" Then
                    builder.Html = MessageHtml
                End If

                If PathAttachment IsNot Nothing Then
                    For Each Anhang As String In PathAttachment
                        If Anhang <> "" AndAlso File.Exists(Anhang) Then
                            Dim attachment As MimeData = builder.AddAttachment(Anhang)
                        End If
                    Next
                End If
            End If

            Dim c As Integer = 0
            If Bitmaps IsNot Nothing AndAlso Bitmaps.Count > 0 Then

                Dim temp As String = Error_Helper.GetTempPath()

                For Each bMap As Bitmap In Bitmaps
                    c += 1
                    Dim filename As String = temp & "Screenshot " & c.ToString & ".jpg"
                    bMap.Save(filename, ImageFormat.Jpeg)
                    Dim attachment As MimeData = builder.AddAttachment(filename)
                    File.Delete(filename)
                Next
            End If

            If Not String.IsNullOrEmpty(SMTPCred.SMTP_Password) AndAlso
                                                       Not String.IsNullOrEmpty(SMTPCred.SMTP_User) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_SenderAddress) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_RecipientAddress) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_RecipientAddress) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_Server) Then
                Dim email As IMail = builder.Create()
                Using Smtp As New Smtp()
                    Try

                        If SMTPCred.SMTP_NoSSL Then
                            Smtp.Connect(SMTPCred.SMTP_Server, SMTPCred.SMTP_NoSSL_Port, False)
                            Smtp.UseBestLogin(SMTPCred.SMTP_User, SMTPCred.SMTP_Password)
                        ElseIf SMTPCred.SMTP_SSL Then
                            Smtp.ConnectSSL(SMTPCred.SMTP_Server, SMTPCred.SMTP_SSL_Port)
                            Smtp.UseBestLogin(SMTPCred.SMTP_User, SMTPCred.SMTP_Password)
                        ElseIf SMTPCred.SMTP_STARTTLS Then
                            Smtp.Connect(SMTPCred.SMTP_Server, SMTPCred.SMTP_STARTTLS_Port)
                            Smtp.StartTLS()
                            Smtp.UseBestLogin(SMTPCred.SMTP_User, SMTPCred.SMTP_Password)
                        End If
                        Dim result As ISendMessageResult = Smtp.SendMessage(email)
                        If result.Status = SendMessageStatus.Success Then
                            ReturnVal = True
                        Else
                            ReturnVal = False
                        End If

                    Catch ex As Exception
                        Call Error_Logger.writelog(Level.Error, "Fehler beim Versenden der Email: " & ex.Message, ex)
                        If Error_Helper.IsIDE() Then
                            MessageBoxEx.Show("Es ist ein Fehler beim Versenden einer System-Email aufgetreten:" & Environment.NewLine & "" & Environment.NewLine & ex.Message, "Mailversand Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        ReturnVal = False
                    End Try
                End Using
            End If
            Return ReturnVal

        Catch ex As Exception
            ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Error_Helper.IsIDE() Then Stop
            Return False
        End Try
    End Function
    Public Shared Function SendEmailUniversal(ByVal RegistryHiveValue As RegistryHive,
                                              ByVal RegistryPath As String,
                                              ByVal Header As String,
                                              ByVal MessageText As String,
                                              ByVal Optional MessageHtml As String = "",
                                              ByVal Optional PathAttachment As List(Of String) = Nothing,
                                              ByVal Optional SMTPCred As Error_Helper.SMTPCredentials = Nothing,
                                              ByVal Optional Bitmaps As List(Of Bitmap) = Nothing) As Boolean
        Dim query As String = ""
        Dim ReturnVal As Boolean = False
        Dim builder As New MailBuilder()

        Try
            If SMTPCred.SMTP_SenderAddress Is Nothing Then
                SMTPCred.SMTP_Password = Error_cCrypt.DecryptString(Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_Password")))
                SMTPCred.SMTP_User = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_User"))
                SMTPCred.SMTP_Server = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_RelayServer"))
                SMTPCred.SMTP_NoSSL = Error_VarConvert.ConvertToBoolean(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_NoSSL"))
                SMTPCred.SMTP_SSL = Error_VarConvert.ConvertToBoolean(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_SSL"))
                SMTPCred.SMTP_STARTTLS = Error_VarConvert.ConvertToBoolean(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_STARTTLS"))
                SMTPCred.SMTP_SenderAddress = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_SenderAddress"))
                SMTPCred.SMTP_SenderName = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_SenderName"))
                SMTPCred.SMTP_RecipientAddress = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_RecipientAddress"))
                SMTPCred.SMTP_RecipientName = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_RecipientName"))
                SMTPCred.SMTP_CC = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_CC"))
                SMTPCred.SMTP_BCC = Error_VarConvert.ConvertToString(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_BCC"))
                SMTPCred.SMTP_NoSSL_Port = Error_VarConvert.ConvertToInteger(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_NoSSL_Port"))
                SMTPCred.SMTP_SSL_Port = Error_VarConvert.ConvertToInteger(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_SSL_Port"))
                SMTPCred.SMTP_STARTTLS_Port = Error_VarConvert.ConvertToInteger(Error_Helper.RegistryReadValue(RegistryHiveValue, RegistryPath & "\email", "SMTP_STARTTLS_Port"))
            End If
            If SMTPCred.SMTP_SenderAddress <> "" Then
                builder.From.Add(New MailBox(SMTPCred.SMTP_SenderAddress, SMTPCred.SMTP_SenderName))
                builder.To.Add(New MailBox(SMTPCred.SMTP_RecipientAddress, SMTPCred.SMTP_RecipientName))
                If SMTPCred.SMTP_CC <> "" Then
                    builder.Cc.Add(New MailBox(SMTPCred.SMTP_CC))
                End If
                If SMTPCred.SMTP_BCC <> "" Then
                    builder.Bcc.Add(New MailBox(SMTPCred.SMTP_BCC))
                End If
                builder.Subject = Header
                builder.Text = MessageText
                If MessageHtml <> "" Then
                    builder.Html = MessageHtml
                End If

                If PathAttachment IsNot Nothing OrElse PathAttachment.Count > 0 Then
                    For Each Anhang As String In PathAttachment
                        If Anhang <> "" AndAlso File.Exists(Anhang) Then
                            Dim attachment As MimeData = builder.AddAttachment(Anhang)
                        End If
                    Next
                End If
            End If

            Dim c As Integer = 0
            If Bitmaps IsNot Nothing AndAlso Bitmaps.Count > 0 Then

                Dim temp As String = Error_Helper.GetTempPath()

                For Each bMap As Bitmap In Bitmaps
                    c += 1
                    Dim filename As String = temp & "Screenshot " & c.ToString & ".jpg"
                    bMap.Save(filename, ImageFormat.Jpeg)
                    Dim attachment As MimeData = builder.AddAttachment(filename)
                    File.Delete(filename)
                Next
            End If

            If Not String.IsNullOrEmpty(SMTPCred.SMTP_Password) AndAlso
                                                       Not String.IsNullOrEmpty(SMTPCred.SMTP_User) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_SenderAddress) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_RecipientAddress) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_RecipientAddress) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_Server) Then
                Dim email As IMail = builder.Create()
                Using Smtp As New Smtp()
                    Try

                        If SMTPCred.SMTP_NoSSL Then
                            Smtp.Connect(SMTPCred.SMTP_Server, SMTPCred.SMTP_NoSSL_Port, False)
                            Smtp.UseBestLogin(SMTPCred.SMTP_User, SMTPCred.SMTP_Password)
                        ElseIf SMTPCred.SMTP_SSL Then
                            Smtp.ConnectSSL(SMTPCred.SMTP_Server, SMTPCred.SMTP_SSL_Port)
                            Smtp.UseBestLogin(SMTPCred.SMTP_User, SMTPCred.SMTP_Password)
                        ElseIf SMTPCred.SMTP_STARTTLS Then
                            Smtp.Connect(SMTPCred.SMTP_Server, SMTPCred.SMTP_STARTTLS_Port)
                            Smtp.StartTLS()
                            Smtp.UseBestLogin(SMTPCred.SMTP_User, SMTPCred.SMTP_Password)
                        End If
                        Dim result As ISendMessageResult = Smtp.SendMessage(email)
                        If result.Status = SendMessageStatus.Success Then
                            ReturnVal = True
                        Else
                            ReturnVal = False
                        End If

                    Catch ex As Exception
                        Call Error_Logger.writelog(Level.Error, "Fehler beim Versenden der Email: " & ex.Message, ex)
                        If Error_Helper.IsIDE() Then
                            MessageBoxEx.Show("Es ist ein Fehler beim Versenden einer System-Email aufgetreten:" & Environment.NewLine & "" & Environment.NewLine & ex.Message, "Mailversand Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        ReturnVal = False
                    End Try
                End Using
            End If
            Return ReturnVal

        Catch ex As Exception
            ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Error_Helper.IsIDE() Then Stop
            Return False
        End Try
    End Function
#End Region

End Class