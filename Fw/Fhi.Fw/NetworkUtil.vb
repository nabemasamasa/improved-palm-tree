Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Text

''' <summary>
''' ネットワーク関連のユーティリティ
''' </summary>
''' <remarks></remarks>
Public Class NetworkUtil
    ''' <summary>
    ''' 処理中のマシンのIPアドレスを返す
    ''' </summary>
    ''' <returns>IPアドレス</returns>
    ''' <remarks></remarks>
    Public Shared Function GetIpAddressesAtMyComputer() As IPAddress()
        Return Dns.GetHostAddresses(Dns.GetHostName)
    End Function

    ''' <summary>
    ''' 処理中のマシンのIPアドレス（文字列）を返す
    ''' </summary>
    ''' <returns>IPアドレス（文字列）</returns>
    ''' <remarks></remarks>
    Public Shared Function GetIpAddressesAtMyComputerAsString() As String()
        Dim result As New List(Of String)
        For Each ip As IPAddress In GetIpAddressesAtMyComputer
            result.Add(ip.ToString)
        Next
        Return result.ToArray
    End Function

    ''' <summary>
    ''' IPv6か？を返す
    ''' </summary>
    ''' <param name="ip">IPアドレス</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsIPv6(ByVal ip As IPAddress) As Boolean
        If ip Is Nothing Then
            Return False
        End If
        Return ip.IsIPv6LinkLocal OrElse ip.IsIPv6Multicast OrElse ip.IsIPv6SiteLocal
    End Function

    ''' <summary>
    ''' IPv4か？を返す
    ''' </summary>
    ''' <param name="ip">IPアドレス</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsIPv4(ByVal ip As IPAddress) As Boolean
        If ip Is Nothing Then
            Return False
        End If
        Return Not IsIPv6(ip)
    End Function

    ''' <summary>
    ''' IPv4か？を返す
    ''' </summary>
    ''' <param name="ipAddress">IPアドレス</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsIPv4(ByVal ipAddress As String) As Boolean
        If StringUtil.IsEmpty(ipAddress) Then
            Return False
        End If
        Dim addresses As String() = Split(ipAddress, ".")
        If addresses.Length <> 4 Then
            Return False
        End If
        For Each address As String In addresses
            If Not IsNumeric(address) OrElse CInt(address) < 0 OrElse 255 < CInt(address) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' メールを送信する
    ''' </summary>
    ''' <param name="smtpServerHostName">SMTPサーバ名</param>
    ''' <param name="smtpServerPortNo">ポート番号</param>
    ''' <param name="fromAddress">メールアドレス（From）</param>
    ''' <param name="toAddresses">メールアドレス（To）</param>
    ''' <param name="smtpAuthUserName">認証ユーザ</param>
    ''' <param name="smtpAuthPassword">認証パスワード</param>
    ''' <param name="enableSSL">SSLの使用有無</param>
    ''' <param name="timeoutMillis">タイムアウト（ミリ秒）</param>
    ''' <param name="ccAddresses">メールアドレス（Cc）</param>
    ''' <param name="bccAddresses">メールアドレス（Bcc）</param>
    ''' <param name="priority">優先度</param>
    ''' <param name="subject">件名</param>
    ''' <param name="body">本文</param>
    ''' <param name="isBodyHtml">HTMLかどうか</param>
    ''' <param name="attachFiles">添付ファイル</param>
    ''' <param name="subjectEncoding">エンコード（件名）</param>
    ''' <param name="bodyEncoding">エンコード（本文）</param>
    ''' <remarks></remarks>
    Public Shared Sub SendMail(ByVal smtpServerHostName As String, ByVal smtpServerPortNo As Integer, _
                               ByVal fromAddress As String, ByVal toAddresses As String(), ByVal subject As String, ByVal body As String, _
                               Optional ByVal smtpAuthUserName As String = Nothing, Optional ByVal smtpAuthPassword As String = Nothing, _
                               Optional ByVal enableSSL As Boolean = False, Optional ByVal timeoutMillis As Integer = 100000, _
                               Optional ByVal ccAddresses As String() = Nothing, Optional ByVal bccAddresses As String() = Nothing, _
                               Optional ByVal priority As MailPriority = MailPriority.Normal, Optional ByVal isBodyHtml As Boolean = False, _
                               Optional ByVal attachFiles As FileInfo() = Nothing, _
                               Optional ByVal subjectEncoding As String = "UTF-8", Optional ByVal bodyEncoding As String = "UTF-8")

        'SMTPサーバを設定
        Dim aSmtpClient As New SmtpClient(smtpServerHostName, smtpServerPortNo)

        'SMTP認証の設定
        If smtpAuthUserName IsNot Nothing AndAlso smtpAuthPassword IsNot Nothing Then
            aSmtpClient.Credentials = New NetworkCredential(smtpAuthUserName, smtpAuthPassword)
        End If

        aSmtpClient.EnableSsl = enableSSL
        aSmtpClient.Timeout = timeoutMillis

        'メールの設定
        Using mail As New MailMessage

            mail.From = New MailAddress(fromAddress)

            If toAddresses IsNot Nothing Then
                For Each addr As String In toAddresses
                    mail.To.Add(addr)
                Next
            End If

            If ccAddresses IsNot Nothing Then
                For Each addr As String In ccAddresses
                    mail.CC.Add(addr)
                Next
            End If

            If bccAddresses IsNot Nothing Then
                For Each addr As String In bccAddresses
                    mail.Bcc.Add(addr)
                Next
            End If

            mail.Priority = priority

            mail.Subject = subject
            mail.Body = body
            mail.IsBodyHtml = isBodyHtml

            '添付ファイル
            If attachFiles IsNot Nothing Then
                For Each f As FileInfo In attachFiles
                    mail.Attachments.Add(New Attachment(f.FullName))
                Next
            End If
            mail.SubjectEncoding = Encoding.GetEncoding(subjectEncoding)
            mail.BodyEncoding = Encoding.GetEncoding(bodyEncoding)

            'メールを送信
            aSmtpClient.Send(mail)
        End Using
    End Sub
End Class
