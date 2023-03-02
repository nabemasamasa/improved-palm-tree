Namespace Util.Concurrent
    ''' <summary>
    ''' JavaのCallableインターフェース相当
    ''' </summary>
    ''' <typeparam name="TResult">値の型</typeparam>
    ''' <remarks></remarks>
    Public Interface Callable(Of TResult)
        Function [Call]() As TResult
    End Interface
End Namespace