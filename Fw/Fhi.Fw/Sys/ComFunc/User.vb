Namespace Sys.ComFunc

    Public Class User

#Region " ユーザー名取得 "

        ''' <summary>
        ''' ユーザー名を返します.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' アプリケーションフレームワークを無効にした場合,
        ''' My.User.InitializeWithWindowsUser() を呼び出さないと
        ''' ユーザー名が空で取得されます.
        ''' また, ドメイン名を含むため, ドメイン名を含まない形で
        ''' ユーザー名を返すようにラッピングした関数です.
        ''' </remarks>
        Public Shared Function GetName() As String
            My.User.InitializeWithWindowsUser()
            Dim bufUserName() As String = My.User.Name.Split("\"c)

            If bufUserName.Length > 1 Then
                Return bufUserName(1)
            Else
                Return bufUserName(0)
            End If
        End Function
#End Region

    End Class

End Namespace
