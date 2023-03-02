Namespace Db
    Public Class ConnectionUtil

        ''' <summary>
        ''' 接続文字列からサーバーの部分を返します.
        ''' </summary>
        ''' <param name="connectionStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetServer(ByVal connectionStr As String) As String
            Dim token() As String = connectionStr.Split(";"c)
            Dim ret As String = String.Empty

            For Each s As String In token
                If InStr(s.ToUpper, "SERVER=") > 0 OrElse _
                   InStr(s.ToUpper, "SQLSERVER:") > 0 Then

                    Dim dbToken() As String = s.Split("="c)

                    If dbToken.Length > 1 Then
                        ret = dbToken(1).Trim()
                        Exit For
                    End If
                End If
            Next

            Return ret
        End Function

        ''' <summary>
        ''' 接続文字列からサーバーの部分を返します.
        ''' </summary>
        ''' <param name="jdbcConnectionStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetServerJdbc(ByVal jdbcConnectionStr As String) As String
            Dim token() As String = jdbcConnectionStr.Split(";"c)
            Dim ret As String = String.Empty

            For Each s As String In token
                If InStr(s.ToUpper, "SERVER") > 0 Then
                    ret = s
                    Exit For
                End If
            Next

            Return ret
        End Function

        ''' <summary>
        ''' 接続文字列からデータベース名の部分を返します.
        ''' </summary>
        ''' <param name="connectionStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetDatabaseName(ByVal connectionStr As String) As String
            Dim token() As String = connectionStr.Split(";"c)
            Dim ret As String = String.Empty

            For Each s As String In token
                If InStr(s.ToUpper, "DATABASE") > 0 Then
                    Dim dbToken() As String = s.Split("="c)

                    If dbToken.Length > 1 Then
                        ret = dbToken(1).Trim()
                        Exit For
                    End If
                End If
            Next

            Return ret
        End Function

        ''' <summary>
        ''' DB接続失敗時のエラーメッセージを作成する "{0} データベースへ接続できません." "{1} の接続文字列を確認してください."
        ''' </summary>
        ''' <param name="databaseName">データベース名</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MakeInvalidDbConnectionMessage(ByVal databaseName As String, ByVal configFileName As String) As String
            Return String.Format("{0} データベースへ接続できません." & vbCrLf & "{1} の接続文字列を確認してください.", databaseName, configFileName)
        End Function

        ''' <summary>
        ''' DB接続失敗時のエラーメッセージを作成する "{0} データベースへ接続できません." " の接続文字列を確認してください."
        ''' </summary>
        ''' <param name="databaseName">データベース名</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MakeInvalidDbConnectionMessage(ByVal databaseName As String) As String
            Return String.Format("{0} データベースへ接続できません.", databaseName)
        End Function
    End Class
End Namespace