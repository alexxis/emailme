'для примера используем сервер mail.ru
Const Login = "mailbox@inbox.ru"
Const PW = "Password"
Const SMTPSrv = "smtp.mail.ru"
 
Set objMail = New MailNotification
'можно изменить порт, аутентификацию, таймаут подключения,
'кодировку и использовать SSL
'With objMail
' .Port = 465 'gmail
' .UseSSL = True
'End With
'можно добавить несколько вложений - вместо Nothing - какие-нибудь логи например
'arrLogs = Array("c:\temp\1.log","c:\temp\2.log")
If objMail.Send(SMTPSrv, Login, PW, Login, "Test address", "Test address", Nothing) Then
 'MsgBox "Сообщение отправлено", vbInformation
Else
 'MsgBox "Не удалось отправить сообщение", vbCritical
End If
Set objMail = Nothing
 
Class MailNotification
 Private m_Msg, m_Conf
 Private m_SMTPPort, m_SMTPAuth, m_SMTPUseSSL, m_SMTPTimeout, m_Charset
   
 Private Sub Class_Initialize()
  Set m_Msg = CreateObject("CDO.Message")
  Set m_Conf = CreateObject("CDO.Configuration")
  'значения по умолчанию
  m_SMTPPort = 465 'порт
  m_SMTPAuth = 1 'базовая аутентификация    
  m_SMTPUseSSL = True 'не использовать SSL
  m_SMTPTimeout = 60 'таймаут подключения
  m_Charset = "windows-1251" 'кодировка  
 End Sub
 Private Sub Class_Terminate()
  Set m_Msg = Nothing
  Set m_Conf = Nothing
 End Sub
  
 Public Property Let Port(i)
  m_SMTPPort = i
 End Property
 Public Property Let Auth(i)
  m_SMTPAuth = i
 End Property
 Public Property Let UseSSL(b)
  m_SMTPUseSSL = b
 End Property
 Public Property Let Timeout(i)
  m_SMTPTimeout = i
 End Property
 Public Property Let Charset(s)
  m_Charset = s
 End Property
  
 Public Function Send(sSMTPSrv, sLogin, sPW, sTo, sSubject, sBody, arrAttachment)
  On Error Resume Next
  With m_Conf.Fields
   'значение 1, которое используется по умолчанию – использовать каталог Pickup
   .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
   '1 - базовая аутентификация, 0 – без аутентификации (анонимно), 2 – аутентификация NTLM
   .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = m_SMTPAuth
   .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = sSMTPSrv
   .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = m_SMTPPort
   .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = sLogin
   .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = sPW
   'использовать ssl
   .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = m_SMTPUseSSL
   'таймаут
   .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = m_SMTPTimeout
   .Update
  End With
  With m_Msg
   .Configuration = m_Conf
   .From = sLogin
   .To = sTo  
   .Subject = sSubject
   .TextBody = sBody
   .Bodypart.Charset = m_Charset ' выставляем кодировку
   If IsArray(arrAttachment) Then
    For i = 0 To UBound(arrAttachment)
     .AddAttachment arrAttachment(i)
    Next
   End If
   .Send
  End With
  If Err.Number = 0 Then Send = True 
 End Function
End Class
