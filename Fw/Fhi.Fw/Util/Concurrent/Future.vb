Namespace Util.Concurrent
    ''' <summary>
    ''' JavaのFutureインターフェース相当
    ''' </summary>
    ''' <typeparam name="TResult">値の型</typeparam>
    ''' <remarks></remarks>
    Public Interface Future(Of TResult)

        Function Cancel(ByVal mayInterruptIfRunning As Boolean) As Boolean

        ReadOnly Property IsCancelled() As Boolean

        ReadOnly Property IsDone() As Boolean

        Function [Get]() As TResult

        Function [Get](ByVal timeout As Long, ByVal unit As TimeUnit) As TResult
    End Interface
End Namespace