Imports System.Threading

Namespace Util.WorkerThread
    ''' <summary>
    ''' ワーカースレッドを保持しタスクの受け渡しを行うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Channel

#Region "Nested classes..."
        Private Class DelegateTask : Implements ITask
            Private ReadOnly aDelegate As [Delegate]
            Private ReadOnly args As Object()
            Public Sub New(ByVal aDelegate As [Delegate], ByVal ParamArray args As Object())
                Me.aDelegate = aDelegate
                Me.args = args
            End Sub
            Public Sub Execute() Implements ITask.Execute
                aDelegate.DynamicInvoke(args)
            End Sub
        End Class

        Private Class AddedTaskCompletion : Implements ITask
            Private ReadOnly task As ITask
            Private ReadOnly completedTaskCallback As Func(Of Object)
            Public Sub New(ByVal task As ITask, ByVal completedTaskCallback As Func(Of Object))
                Me.task = task
                Me.completedTaskCallback = completedTaskCallback
            End Sub
            Public Sub Execute() Implements ITask.Execute
                task.Execute()
                completedTaskCallback.Invoke()
            End Sub
        End Class

        Private Class AfterAllTasksDone
            Private ReadOnly chan As Channel
            Private ReadOnly postTask As DelegateTask

            Public Sub New(ByVal chan As Channel, ByVal postTask As DelegateTask)
                Me.chan = chan
                Me.postTask = postTask
            End Sub

            ''' <summary>
            ''' タスクがすべて終了したらスレッドワーカーを終了する
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub StopWorkersAfterAllTasksDone()
                chan.WaitUntilTaskDone()
                chan.StopWorkers()
                If postTask IsNot Nothing Then
                    postTask.Execute()
                End If
            End Sub
        End Class
#End Region

        ''' <summary>引数なしSubメソッドのdelegate</summary>
        ''' <remarks>System.Windows.Formsを参照していないプロジェクトに必要</remarks>
        Public Delegate Sub MethodInvoker()

        Private Const MAX_QUEUE_COUNT As Integer = 100
        Private ReadOnly taskQueues As ITask()
        Private head As Integer = 0
        Private tail As Integer = 0
        Private queueCount As Integer = 0
        Private executingCount As Integer = 0
        Private requestQueueCount As Integer = 0
        Private ReadOnly aThreadPool As WorkerThread()

        Public Sub New(Optional ByVal threadCount As Integer = 1)
            Me.taskQueues = New ITask(MAX_QUEUE_COUNT - 1) {}
            aThreadPool = New WorkerThread(threadCount - 1) {}
            For i As Integer = 0 To aThreadPool.Length - 1
                aThreadPool(i) = New WorkerThread("Worker-" & i, Me)
            Next
        End Sub

        ''' <summary>
        ''' スレッドワーカーを開始する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub StartWorkers()
            For i As Integer = 0 To aThreadPool.Length - 1
                aThreadPool(i).Start()
            Next
        End Sub

        ''' <summary>
        ''' スレッドワーカーを終了する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub StopWorkers()
            For i As Integer = 0 To aThreadPool.Length - 1
                aThreadPool(i).StopWorker()
            Next
        End Sub

        ''' <summary>
        ''' タスクがすべて終了したらスレッドワーカーを終了する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub StopWorkersAfterAllTasksDone()
            StopWorkersAfterAllTasksDone(Nothing)
        End Sub
        ''' <summary>
        ''' タスクがすべて終了したらスレッドワーカーを終了する
        ''' </summary>
        ''' <param name="postProcess">後処理デリゲート</param>
        ''' <remarks></remarks>
        Public Sub StopWorkersAfterAllTasksDone(ByVal postProcess As MethodInvoker)
            StopWorkersAfterAllTasksDone(postProcess, Nothing)
        End Sub
        ''' <summary>
        ''' タスクがすべて終了したらスレッドワーカーを終了する
        ''' </summary>
        ''' <param name="postProcess">後処理デリゲート</param>
        ''' <param name="args">引数</param>
        ''' <remarks></remarks>
        Public Sub StopWorkersAfterAllTasksDone(ByVal postProcess As [Delegate], ByVal ParamArray args As Object())
            Dim post As New AfterAllTasksDone(Me, If(postProcess Is Nothing, Nothing, New DelegateTask(postProcess, args)))
            Dim t As New Thread(New ThreadStart(AddressOf post.StopWorkersAfterAllTasksDone))
            t.IsBackground = True
            t.Start()
        End Sub

        ''' <summary>
        ''' タスクを登録する
        ''' </summary>
        ''' <param name="aDelegate">デリゲート</param>
        ''' <remarks></remarks>
        Public Sub RegisterTask(ByVal aDelegate As MethodInvoker)
            Me.RegisterTask(aDelegate, Nothing)
        End Sub

        ''' <summary>
        ''' タスクを登録する
        ''' </summary>
        ''' <param name="aDelegate">デリゲート</param>
        ''' <remarks></remarks>
        Public Sub RegisterTask(ByVal aDelegate As [Delegate])
            Me.RegisterTask(aDelegate, Nothing)
        End Sub

        ''' <summary>
        ''' タスクを登録する
        ''' </summary>
        ''' <param name="aDelegate">デリゲート</param>
        ''' <param name="args">引数</param>
        ''' <remarks></remarks>
        Public Sub RegisterTask(ByVal aDelegate As [Delegate], ByVal ParamArray args As Object())
            RegisterTask(New DelegateTask(aDelegate, args))
        End Sub

        ''' <summary>
        ''' タスクを登録する
        ''' </summary>
        ''' <param name="aTask">タスク</param>
        ''' <remarks></remarks>
        Public Sub RegisterTask(ByVal aTask As ITask)
            Interlocked.Increment(requestQueueCount)
            Monitor.Enter(Me)
            Try
                WaitUntilTaskRegisterable()
                taskQueues(tail) = aTask
                AddCount()
                Interlocked.Decrement(requestQueueCount)
                Monitor.PulseAll(Me)
            Finally
                Monitor.Exit(Me)
            End Try
        End Sub

        ''' <summary>
        ''' タスクを取り出す
        ''' </summary>
        ''' <returns>タスク</returns>
        ''' <remarks></remarks>
        Public Function TakeTask() As ITask
            Monitor.Enter(Me)
            Try
                WaitUntilTaskRegistered()
                Dim task As ITask = taskQueues(head)
                ReduceCount()
                Interlocked.Increment(executingCount)
                Monitor.PulseAll(Me)
                Return New AddedTaskCompletion(task, completedTaskCallback:=Function() Interlocked.Decrement(executingCount))
            Finally
                Monitor.Exit(Me)
            End Try
        End Function

        ''' <summary>
        ''' キューカウントを増やす
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub AddCount()
            tail = (tail + 1) Mod taskQueues.Length
            Interlocked.Increment(queueCount)
        End Sub

        ''' <summary>
        ''' キューカウントを減らす
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ReduceCount()
            head = (head + 1) Mod taskQueues.Length
            Interlocked.Decrement(queueCount)
        End Sub

        ''' <summary>
        ''' タスクを実行中かどうか
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsRunning() As Boolean
            Monitor.Enter(Me)
            Try
                Return 0 < queueCount OrElse 0 < executingCount OrElse 0 < requestQueueCount
            Finally
                Monitor.Exit(Me)
            End Try
        End Function

        ''' <summary>
        ''' タスクが登録可能になるまで待機する
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub WaitUntilTaskRegisterable()
            While taskQueues.Length <= queueCount
                Try
                    Monitor.Wait(Me)
                Catch ignore As ThreadInterruptedException
                End Try
            End While
        End Sub

        ''' <summary>
        ''' タスクが登録されるまで待機する
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub WaitUntilTaskRegistered()
            While queueCount <= 0
                Try
                    Monitor.Wait(Me)
                Catch ignore As ThreadInterruptedException
                End Try
            End While
        End Sub

        ''' <summary>
        ''' すべてのタスクが終了するまで待機する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub WaitUntilTaskDone()
            While IsRunning()
                Thread.Sleep(10)
            End While
        End Sub
    End Class
End Namespace