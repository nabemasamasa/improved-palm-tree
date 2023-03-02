Imports NUnit.Framework
Imports System.Threading
Imports Fhi.Fw.Lang.Threading

Namespace Util.Concurrent
    Public MustInherit Class FutureTaskTest

        Private Class TestingCallable : Implements Callable(Of String)

            Private ReadOnly millisSleep As Integer
            Public ReadOnly Log As List(Of String)

            Public Sub New(ByVal millisSleep As Integer)
                Me.millisSleep = millisSleep
                Me.Log = New List(Of String)
            End Sub

            Public Function [Call]() As String Implements Callable(Of String).[Call]
                Log.Add("Caller sleep start")
                Thread.Sleep(millisSleep)
                Log.Add("Caller sleep end")
                Return "Result"
            End Function
        End Class

        Private Class ExceptionCallable(Of E As Exception) : Implements Callable(Of String)

            Public Function [Call]() As String Implements Callable(Of String).[Call]
                Throw VoUtil.NewInstance(Of E)()
            End Function
        End Class

        Private Class GetThread : Inherits AThread

            Private ReadOnly aFuture As FutureTask(Of String)
            Private ReadOnly log As List(Of String)

            Public Sub New(ByVal aFuture As FutureTask(Of String), ByVal log As List(Of String))
                Me.aFuture = aFuture
                Me.log = log
            End Sub

            Public Overrides Sub Run()
                log.Add("Getter start")
                Dim time As New Stopwatch
                time.Start()
                Dim value As String = aFuture.Get()
                time.Stop()
                log.Add(StringUtil.Format("0.000", time.ElapsedMilliseconds / 1000))
                log.Add(String.Format("Got '{0}'", value))
            End Sub
        End Class

        Public Class DefaultTest : Inherits FutureTaskTest

            <Test()> Public Sub Get_Futureの値が設定されるまでGetterが待たされる()
                Dim millisSleep As Integer = 10
                Dim callable As TestingCallable = New TestingCallable(millisSleep)
                Dim aFuture As New FutureTask(Of String)(callable)

                Dim caller As New AThread(aFuture)
                Dim getter As New GetThread(aFuture, callable.Log)

                getter.Start()
                While callable.Log.Count = 0
                    Thread.Sleep(10)
                End While
                caller.Start()

                getter.Join()

                Assert.AreEqual("Getter start", callable.Log(0))
                Assert.AreEqual("Caller sleep start", callable.Log(1))
                Assert.AreEqual("Caller sleep end", callable.Log(2))
                Assert.AreEqual("Got 'Result'", callable.Log(4))
            End Sub

            <Test()> Public Sub Get_Callable処理中の例外はGetメソッド呼び出し時にExecutionExceptionにラップされて返る()
                Dim aFuture As New FutureTask(Of String)(New ExceptionCallable(Of ArgumentException))

                Dim caller As New AThread(aFuture)

                caller.Start()

                Try
                    aFuture.Get()
                    Assert.Fail()
                Catch expected As ExecutionException
                    Assert.IsTrue(TypeOf expected.InnerException Is ArgumentException)
                End Try
            End Sub

            <Test()> Public Sub Get_CancelされればGetメソッド呼出時にCancellationException()
                Dim callable As TestingCallable = New TestingCallable(50)
                Dim aFuture As New FutureTask(Of String)(callable)

                Dim caller As New AThread(aFuture)

                caller.Start()

                Dim actual As Boolean = aFuture.Cancel(False)

                caller.Join()
                Try
                    Assert.IsTrue(actual)

                    aFuture.Get()
                    Assert.Fail()
                Catch expected As CancellationException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub Cancel_True_なら処理中スレッドをInterruptしてCancelする_Getを呼べばCancellationException()
                Dim callable As TestingCallable = New TestingCallable(100)
                Dim aFuture As New FutureTask(Of String)(callable)

                Dim caller As New AThread(aFuture, "caller")

                caller.Start()

                While callable.Log.Count = 0
                    Thread.Sleep(10)
                End While
                Do Until aFuture.Cancel(True)
                    ' callerスレッドが始まっていないから待つ
                    Thread.Sleep(10)
                Loop

                Assert.AreEqual("Caller sleep start", callable.Log(0))

                Try
                    aFuture.Get()
                    Assert.Fail()
                Catch expected As CancellationException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub Cancel_False_なら処理中スレッドはそのままでCancelする_Getを呼べばCancellationException()
                Dim callable As TestingCallable = New TestingCallable(50)
                Dim aFuture As New FutureTask(Of String)(callable)

                Dim caller As New AThread(aFuture, "caller")

                caller.Start()
                While callable.Log.Count = 0
                    Thread.Sleep(10)
                End While
                Do Until aFuture.Cancel(False)
                    ' callerスレッドが始まっていないから待つ
                    Thread.Sleep(10)
                Loop

                caller.Join()

                Assert.AreEqual(2, callable.Log.Count)
                Assert.AreEqual("Caller sleep start", callable.Log(0))
                Assert.AreEqual("Caller sleep end", callable.Log(1))

                Try
                    aFuture.Get()
                    Assert.Fail()
                Catch expected As CancellationException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub 処理中の例外は_ExecutionExceptionでラップする_例外タイトルも引き継ぐ()
                Dim sut As New FutureTask(Of String)(Function()
                                                         Thread.Sleep(10)
                                                         Throw New ArgumentException("aiueo")
                                                     End Function)
                Try
                    sut.RunAsync()
                    Dim dummy As String = sut.Get()
                    Assert.Fail()
                Catch ex As ExecutionException
                    Assert.That(ex.Message, [Is].EqualTo("aiueo"), "例外タイトルも引き継ぐ")
                    Assert.That(ex.InnerException, [Is].TypeOf(Of ArgumentException))
                End Try
            End Sub
        End Class

    End Class
End Namespace