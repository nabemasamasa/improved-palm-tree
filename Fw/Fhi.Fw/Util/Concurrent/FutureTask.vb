Imports System.Threading
Imports Fhi.Fw.Lang.Threading

Namespace Util.Concurrent
    ''' <summary>
    ''' JavaのFutureTaskクラス相当（簡易版）
    ''' </summary>
    ''' <typeparam name="TResult">値の型</typeparam>
    ''' <remarks></remarks>
    Public Class FutureTask(Of TResult) : Implements RunnableFuture(Of TResult)

        Private ReadOnly thisSync As Sync

        Public Sub New(ByVal callImpl As Executors.CallCallback(Of TResult))
            Me.New(Executors.Callable(callImpl))
            If callImpl Is Nothing Then
                Throw New NullReferenceException
            End If
        End Sub

        Public Sub New(ByVal aCallable As Callable(Of TResult))

            If aCallable Is Nothing Then
                Throw New NullReferenceException
            End If
            thisSync = New Sync(aCallable, Me)
        End Sub

        Public Sub New(ByVal runnable As Runnable, ByVal result As TResult)
            thisSync = New Sync(Executors.Callable(runnable, result), Me)
        End Sub

        ''' <summary>
        ''' 中断する
        ''' </summary>
        ''' <param name="mayInterruptIfRunning">処理中スレッドをInterruptするなら、true</param>
        ''' <returns>中断した場合、true</returns>
        ''' <remarks></remarks>
        Public Function Cancel(ByVal mayInterruptIfRunning As Boolean) As Boolean Implements Future(Of TResult).Cancel
            Return thisSync.InnerCancel(mayInterruptIfRunning)
        End Function

        Public ReadOnly Property IsCancelled() As Boolean Implements Future(Of TResult).IsCancelled
            Get
                Return thisSync.InnerIsCancelled
            End Get
        End Property

        Public ReadOnly Property IsDone() As Boolean Implements Future(Of TResult).IsDone
            Get
                Return thisSync.InnerIsDone
            End Get
        End Property

        Protected Overridable Sub Done()

        End Sub

        Protected Overridable Sub [Set](ByVal value As TResult)
            thisSync.InnerSet(value)
        End Sub

        ''' <summary>
        ''' 値を取得する
        ''' </summary>
        ''' <returns>値</returns>
        ''' <exception cref="CancellationException">取り消された場合（Celcelメソッドが呼ばれた場合）</exception> 
        ''' <exception cref="ExecutionException">Callableで例外が発生した場合</exception> 
        ''' <remarks></remarks>
        Public Function [Get]() As TResult Implements Future(Of TResult).[Get]
            Return thisSync.InnerGet()
        End Function

        ''' <summary>
        ''' 値を取得する
        ''' </summary>
        ''' <param name="timeout"></param>
        ''' <param name="unit"></param>
        ''' <returns>値</returns>
        ''' <exception cref="CancellationException">取り消された場合（Celcelメソッドが呼ばれた場合）</exception> 
        ''' <exception cref="ExecutionException">Callableで例外が発生した場合</exception> 
        ''' <remarks></remarks>
        Public Function [Get](ByVal timeout As Long, ByVal unit As TimeUnit) As TResult Implements Future(Of TResult).[Get]
            Return thisSync.InnerGet(CInt(unit.ToMillis(timeout)))
        End Function

        Public Sub Run() Implements Runnable.Run
            thisSync.InnerRun()
        End Sub

        ''' <summary>
        ''' 処理を非同期実行する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub RunAsync()
            Call (New Thread(AddressOf Run)).Start()
        End Sub

        Private Class Sync

            Public Const RUNNING As Integer = 1
            Public Const RAN As Integer = 2
            Public Const CANCELLED As Integer = 4

            Private ReadOnly caller As FutureTask(Of TResult)

            Private ReadOnly aCallable As Callable(Of TResult)
            Private result As TResult
            Private anException As Exception

            Private _state As Integer
            Private stateLock As New Object
            Private _runner As Object

            Public Sub New(ByVal aCallable As Callable(Of TResult), ByVal caller As FutureTask(Of TResult))
                Me.aCallable = aCallable
                Me.caller = caller
            End Sub

            Private Sub LogDebug(message As String, ParamArray args As Object())
                ' テスト時に使いたいのでコメントアウト
                'EzUtil.logDebug(message, args)
            End Sub

            Private Property State() As Integer
                Get
                    Return Thread.VolatileRead(_state)
                End Get
                Set(ByVal value As Integer)
                    LogDebug("{0}: State = {1}", Thread.CurrentThread.Name, value)
                    Thread.VolatileWrite(_state, value)
                End Set
            End Property

            Private Property Runner() As Thread
                Get
                    Return DirectCast(Thread.VolatileRead(_runner), Thread)
                End Get
                Set(ByVal value As Thread)
                    Thread.VolatileWrite(_runner, value)
                End Set
            End Property

            Public Function InnerIsCancelled() As Boolean
                Return State = CANCELLED
            End Function

            Private Function RanOrCancelled(ByVal state As Integer) As Boolean
                Return (state And (RAN Or CANCELLED)) <> 0
            End Function

            Public Function InnerIsDone() As Boolean
                Return RanOrCancelled(State) AndAlso Runner Is Nothing
            End Function

            Public Function InnerCancel(ByVal mayInterruptIfRunning As Boolean) As Boolean
                Try
                    LogDebug("{0}: InnerCancel({1}) SyncLocked", Thread.CurrentThread.Name, mayInterruptIfRunning.ToString)

                    SyncLock stateLock
                        If RanOrCancelled(State) Then
                            LogDebug("{0}: InnerCancel abort.", Thread.CurrentThread.Name)
                            Return False
                        End If

                        'TODO mayInterruptIfRunning=trueじゃないとキャンセルしないけど、ステータスだけ変わるのは！？
                        State = CANCELLED
                    End SyncLock

                    If mayInterruptIfRunning Then
                        LogDebug("{0}: runner {1} Nothing", Thread.CurrentThread.Name, If(Runner Is Nothing, "Is", "IsNot"))
                        If Runner IsNot Nothing Then
                            Runner.Interrupt()
                            LogDebug("{0}: runner('{1} Thread').Interrupt", Thread.CurrentThread.Name, Runner.Name)
                        End If
                    End If

                    SyncLock Me
                        caller.Done()
                        Monitor.PulseAll(Me)
                    End SyncLock

                    Return True
                Finally
                    LogDebug("{0}: InnerCancel({1}) SyncLock release", Thread.CurrentThread.Name, mayInterruptIfRunning.ToString)
                End Try

            End Function

            Public Function InnerGet() As TResult
                SyncLock Me
                    Try
                        LogDebug("{0}: InnerGet() SyncLocked", Thread.CurrentThread.Name)
                        While Not RanOrCancelled(State)
                            Monitor.Wait(Me)
                        End While

                        If State = CANCELLED Then
                            Throw New CancellationException
                        ElseIf anException IsNot Nothing Then
                            Throw New ExecutionException(anException.Message, anException)
                        End If
                        Return result
                    Finally
                        LogDebug("{0}: InnerGet() SyncLock release", Thread.CurrentThread.Name)
                    End Try
                End SyncLock
            End Function

            Public Function InnerGet(ByVal millisTimeout As Integer) As TResult
                SyncLock Me
                    Try
                        LogDebug("{0}: InnerGet({1}) SyncLocked", Thread.CurrentThread.Name, millisTimeout)
                        While Not RanOrCancelled(State)
                            Monitor.Wait(Me, millisTimeout)
                        End While

                        If State = CANCELLED Then
                            Throw New CancellationException
                        ElseIf anException IsNot Nothing Then
                            Throw New ExecutionException(anException.Message, anException)
                        End If
                        Return result
                    Finally
                        LogDebug("{0}: InnerGet({1}) SyncLock release", Thread.CurrentThread.Name, millisTimeout)
                    End Try
                End SyncLock
            End Function

            Public Sub InnerSet(ByVal value As TResult)
                SyncLock Me
                    Try
                        LogDebug("{0}: InnerSet({1}) SyncLocked", Thread.CurrentThread.Name, StringUtil.ToString(value))
                        SyncLock stateLock
                            If State = RAN Then
                                Return
                            ElseIf State = CANCELLED Then
                                Return
                            End If

                            LogDebug("{0}: InnerSet({1})", Thread.CurrentThread.Name, StringUtil.ToString(value))

                            State = RAN
                        End SyncLock

                        result = value
                        caller.Done()
                        Monitor.PulseAll(Me)
                    Finally
                        LogDebug("{0}: InnerSet({1}) SyncLock release", Thread.CurrentThread.Name, StringUtil.ToString(value))
                    End Try
                End SyncLock
            End Sub

            Public Sub InnerSetException(ByVal ex As Exception)
                SyncLock Me
                    Try
                        SyncLock stateLock
                            LogDebug("{0}: InnerSetException({1}) SyncLocked", Thread.CurrentThread.Name, ex.GetType.Name)
                            If State = RAN Then
                                Return
                            ElseIf State = CANCELLED Then
                                Return
                            End If

                            LogDebug("{0}: InnerSetException({1})", Thread.CurrentThread.Name, ex.GetType.Name)

                            State = RAN
                        End SyncLock

                        anException = ex
                        result = Nothing
                        caller.Done()
                        Monitor.PulseAll(Me)
                    Finally
                        LogDebug("{0}: InnerSetException({1}) SyncLock release", Thread.CurrentThread.Name, ex.GetType.Name)
                    End Try
                End SyncLock
            End Sub

            Public Sub InnerRun()
                SyncLock Me
                    LogDebug("{0}: InnerRun() SyncLocked", Thread.CurrentThread.Name)
                    Try
                        SyncLock stateLock
                            If State <> 0 Then
                                Return
                            End If
                            State = RUNNING
                        End SyncLock
                        Try
                            Runner = Thread.CurrentThread
                            'Thread.VolatileWrite(DirectCast(runner, Object), Thread.CurrentThread)
                            LogDebug("{0}: InnerRun runner set", Thread.CurrentThread.Name)
                            If State = RUNNING Then
                                InnerSet(aCallable.Call())
                            End If
                        Catch ex As Exception
                            LogDebug("{0}: InnerRun catch exception {1}", Thread.CurrentThread.Name, ex.GetType.Name)
                            InnerSetException(ex)
                        End Try
                    Finally
                        LogDebug("{0}: InnerRun() SyncLock release", Thread.CurrentThread.Name)
                    End Try
                End SyncLock
            End Sub
        End Class

    End Class
End Namespace