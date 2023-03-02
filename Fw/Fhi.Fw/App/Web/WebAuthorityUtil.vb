Imports System.Security.Principal

Namespace App.Web
    ''' <summary>
    ''' Web認証系について汎用的な機能をまとめたクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class WebAuthorityUtil
#Region "Nested Classes"
        Private Class InvalidAuthenticationException : Inherits BaseException
            Public Sub New()
            End Sub
            Public Sub New(message As String)
                MyBase.New(message)
            End Sub
            Public Sub New(message As String, innerException As Exception)
                MyBase.New(message, innerException)
            End Sub
        End Class
#End Region

        ''' <summary>
        ''' 現在のログオンユーザーを取得する
        ''' </summary>
        ''' <param name="currentPrincipal"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLogonUser(currentPrincipal As IPrincipal) As String
            If Not TypeOf currentPrincipal.Identity Is WindowsIdentity Then
                Throw New InvalidAuthenticationException(String.Format("Windows認証以外({0})が設定されています。ログオンユーザーを取得できません。", currentPrincipal.Identity.GetType.Name))
            End If

            'impersonateする事でSystem.Environment.UserNameをWindows認証のユーザーに挿げ替えている
            'Environment.UserNameにしているのは、WindowsIdentityで取得するユーザー名はドメイン考慮されてないユーザー名の為(tky/tky00112 Or tky00112などブレがある)
            Dim impersonationContext As WindowsImpersonationContext = DirectCast(currentPrincipal.Identity, System.Security.Principal.WindowsIdentity).Impersonate()
            Try
                Return System.Environment.UserName
            Finally
                impersonationContext.Undo()
            End Try
        End Function

    End Class
End Namespace