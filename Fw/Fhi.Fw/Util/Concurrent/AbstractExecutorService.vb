Imports Fhi.Fw.Lang.Threading

Namespace Util.Concurrent
    Public MustInherit Class AbstractExecutorService : Implements ExecutorService

        Protected Overridable Function NewTaskFor(Of T)(ByVal runnable As Runnable, ByVal value As T) As RunnableFuture(Of T)
            Return New FutureTask(Of T)(runnable, value)
        End Function

        Protected Overridable Function NewTaskFor(Of T)(ByVal callable As Callable(Of T)) As RunnableFuture(Of T)
            Return New FutureTask(Of T)(callable)
        End Function

        Public Function Submit(Of T)(ByVal task As Callable(Of T)) As Future(Of T) Implements ExecutorService.Submit
            If task Is Nothing Then
                Throw New NullReferenceException
            End If
            Dim ftask As RunnableFuture(Of T) = NewTaskFor(Of T)(task)
            Execute(ftask)
            Return ftask
        End Function

        Public Function Submit(Of T)(ByVal task As Runnable, ByVal result As T) As Future(Of T) Implements ExecutorService.Submit
            If task Is Nothing Then
                Throw New NullReferenceException
            End If
            Dim ftask As RunnableFuture(Of T) = NewTaskFor(Of T)(task, result)
            Execute(ftask)
            Return ftask
        End Function

        Public Function Submit(Of Wildcard)(ByVal task As Runnable) As Future(Of Wildcard) Implements ExecutorService.Submit
            If task Is Nothing Then
                Throw New NullReferenceException
            End If
            Dim ftask As RunnableFuture(Of Wildcard) = NewTaskFor(Of Wildcard)(task, Nothing)
            Execute(ftask)
            Return ftask
        End Function

        Public MustOverride Sub Execute(ByVal command As Runnable) Implements Executor.Execute

        Public MustOverride Sub Shutdown() Implements ExecutorService.Shutdown

        Public MustOverride Function IsShutdown() As Boolean Implements ExecutorService.IsShutdown

        Public MustOverride Function IsTerminated() As Boolean Implements ExecutorService.IsTerminated

    End Class
End Namespace