Namespace Db
    ''' <summary>
    ''' テーブル排他制御の例外クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class TableExclusionException : Inherits BaseException
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            MyBase.New()
        End Sub
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="exStr">メッセージ</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal exStr As String)
            MyBase.New(exStr)
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