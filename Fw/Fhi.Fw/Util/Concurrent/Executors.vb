Imports System.Threading
Imports Fhi.Fw.Lang.Threading

Namespace Util.Concurrent

    Public Class Executors

#Region "DefaultThreadFactory"
        Friend Class DefaultThreadFactoryImpl : Implements ThreadFactory

            Friend Shared poolNumber As Integer = 1
            Friend threadNumber As Integer = 1
            Friend ReadOnly namePrefix As String

            Friend Sub New()
                SyncLock GetType(DefaultThreadFactoryImpl)
                    Me.namePrefix = String.Format("pool-{0}-thread-", poolNumber)
                    poolNumber += 1
                End SyncLock
            End Sub

            Public Function NewThread(ByVal r As Runnable) As Thread Implements ThreadFactory.NewThread
                Return NewThread(AddressOf r.Run)
            End Function

            Public Function NewThread(ByVal dlgtThreadStart As ThreadStart) As Thread Implements ThreadFactory.NewThread
                Dim hoge As New Thread(dlgtThreadStart)
                SyncLock Me
                    hoge.Name = namePrefix & threadNumber
                    threadNumber += 1
                End SyncLock
                If Not hoge.IsBackground Then
                    hoge.IsBackground = True
                End If
                If hoge.Priority <> ThreadPriority.Normal Then
                    hoge.Priority = ThreadPriority.Normal
                End If
                Return hoge
            End Function
        End Class
#End Region

#Region "RunnableAdapter"
        Friend Class RunnableAdapter(Of T) : Implements Callable(Of T)

            Friend ReadOnly task As Runnable
            Friend ReadOnly result As T

            Public Sub New(ByVal task As Runnable, ByVal result As T)
                Me.task = task
                Me.result = result
            End Sub

            Public Function [Call]() As T Implements Callable(Of T).[Call]
                task.Run()
                Return result
            End Function
        End Class
#End Region

#Region "CallableDlgtAdapter"
        Private Class CallableDlgtAdapter(Of TResult) : Implements Callable(Of TResult)

            Private ReadOnly dlgtCall As CallCallback(Of TResult)

            Public Sub New(ByVal dlgtCall As CallCallback(Of TResult))
                Me.dlgtCall = dlgtCall
            End Sub

            Public Function [Call]() As TResult Implements Callable(Of TResult).[Call]
                Return dlgtCall.Invoke
            End Function
        End Class
#End Region

        Public Delegate Function CallCallback(Of TResult)() As TResult

        Public Shared Function DefaultThreadFactory() As ThreadFactory
            Return New DefaultThreadFactoryImpl
        End Function

        ''' <summary>
        ''' DelegateをCallable実装のインスタンスにする
        ''' </summary>
        ''' <typeparam name="T">Callableの型</typeparam>
        ''' <param name="callee">Callable#Callとなるデリゲート</param>
        ''' <returns>Callable実装のインスタンス</returns>
        ''' <remarks></remarks>
        Public Shared Function Callable(Of T)(ByVal callee As CallCallback(Of T)) As Callable(Of T)
            If callee Is Nothing Then
                Throw New NullReferenceException
            End If
            Return New CallableDlgtAdapter(Of T)(callee)
        End Function

        ''' <summary>
        ''' 処理終了時にresultを返すCallable実装のインスタンスを返す
        ''' </summary>
        ''' <typeparam name="T">resultの型</typeparam>
        ''' <param name="task">処理</param>
        ''' <param name="result">戻り値</param>
        ''' <returns>Callable実装のインスタンス</returns>
        ''' <remarks></remarks>
        Public Shared Function Callable(Of T)(ByVal task As Runnable, ByVal result As T) As Callable(Of T)
            If task Is Nothing Then
                Throw New NullReferenceException
            End If
            Return New RunnableAdapter(Of T)(task, result)
        End Function
    End Class
End Namespace