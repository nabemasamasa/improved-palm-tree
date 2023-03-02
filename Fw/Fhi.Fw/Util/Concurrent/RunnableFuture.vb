Imports Fhi.Fw.Lang.Threading

Namespace Util.Concurrent
    ''' <summary>
    ''' JavaのRunnableFutureインターフェース相当
    ''' </summary>
    ''' <typeparam name="TResult">値の型</typeparam>
    ''' <remarks></remarks>
    Public Interface RunnableFuture(Of TResult) : Inherits Future(Of TResult), Runnable

    End Interface
End Namespace