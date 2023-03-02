Namespace Db
    ''' <summary>
    ''' ロック要求のタイムアウト例外
    ''' </summary>
    ''' <remarks></remarks>
    Public Class LockTimeoutException : Inherits DatabaseException

        Private Const DEFAULT_MESSAGE As String = "データベースのロックタイムアウトです"

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            MyBase.New(DEFAULT_MESSAGE)
        End Sub
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="message">メッセージ</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="message">メッセージ</param>
        ''' <param name="innerException">元例外</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal message As String, ByVal innerException As Exception)
            MyBase.New(message, innerException)
        End Sub
    End Class

End Namespace