Imports Fhi.Fw.Lang.Threading

Namespace Util.Concurrent
    ''' <summary>
    ''' 再利用されるスレッドの実行を抽象化したインターフェース
    ''' </summary>
    ''' <remarks>JavaのExecutorServiceインターフェース相当</remarks>
    Public Interface ExecutorService : Inherits Executor
        Sub Shutdown()
        Function IsShutdown() As Boolean
        Function IsTerminated() As Boolean
        Function Submit(Of T)(ByVal task As Callable(Of T)) As Future(Of T)
        Function Submit(Of T)(ByVal task As Runnable, ByVal result As T) As Future(Of T)
        Function Submit(Of Wildcard)(ByVal task As Runnable) As Future(Of Wildcard)
    End Interface
End Namespace