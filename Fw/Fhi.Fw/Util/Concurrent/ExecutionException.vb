Namespace Util.Concurrent
    Public Class ExecutionException : Inherits Exception

        Public Sub New()
        End Sub

        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub

        Public Sub New(ByVal message As String, ByVal innerException As Exception)
            MyBase.New(message, innerException)
        End Sub

        Public Sub New(ByVal innerException As Exception)
            MyBase.New("", innerException)
        End Sub
    End Class
End Namespace