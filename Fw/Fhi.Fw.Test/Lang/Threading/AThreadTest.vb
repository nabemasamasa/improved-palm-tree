Imports NUnit.Framework
Imports System.Threading
Imports Fhi.Fw.Lang.Threading

Namespace Lang.Threading
    Public Class AThreadTest

        Private Class ThreadInterruptedResult
            Private _Result As Boolean?

            Public Property Result() As Boolean
                Get
                    SyncLock Me
                        While _Result Is Nothing
                            Monitor.Wait(Me)
                        End While
                        Return _Result.Value
                    End SyncLock
                End Get
                Friend Set(ByVal value As Boolean)
                    SyncLock Me
                        _Result = value
                        Monitor.PulseAll(Me)
                    End SyncLock
                End Set
            End Property
        End Class

        Private Class TestingThread : Inherits AThread

            Private _shutdown As Boolean
            Private _RequestCallThreadInterrupted As Boolean
            Private _ThreadInterruptedResult As ThreadInterruptedResult

            Public Sub Shutdown()
                _shutdown = True
            End Sub

            Public Overrides Sub Run()
                While Not _shutdown
                    If _RequestCallThreadInterrupted Then
                        _RequestCallThreadInterrupted = False
                        _ThreadInterruptedResult.Result = AThread.Interrupted()
                    End If
                End While
            End Sub

            ''' <summary>
            ''' TestingThreadのスレッドで AThread.Interrupted() の呼び出しを要求する
            ''' </summary>
            ''' <returns>AThread.Interrupted()の結果</returns>
            ''' <remarks></remarks>
            Public Function RequestCallThreadInterrupted() As ThreadInterruptedResult
                _ThreadInterruptedResult = New ThreadInterruptedResult
                _RequestCallThreadInterrupted = True
                Return _ThreadInterruptedResult
            End Function
        End Class

        <Test()> Public Sub Interrupted_スレッドが割り込まれていないならfalse()
            Dim testee As New TestingThread
            testee.Start()

            Dim future As ThreadInterruptedResult = testee.RequestCallThreadInterrupted()

            Assert.IsFalse(future.Result)

            testee.Shutdown()
            testee.Join()
        End Sub

        <Test()> Public Sub Interrupted_スレッドが割り込まれていればtrue()
            Dim testee As New TestingThread
            testee.Start()
            testee.Interrupt()

            Dim future As ThreadInterruptedResult = testee.RequestCallThreadInterrupted()

            Assert.IsTrue(future.Result)

            testee.Shutdown()
            testee.Join()
        End Sub

        <Test()> Public Sub Interrupted_割り込みステータスがクリアされるから二度続けると二度目はfalse()
            Dim testee As New TestingThread
            testee.Start()
            testee.Interrupt()

            Dim future As ThreadInterruptedResult = testee.RequestCallThreadInterrupted()

            Assert.IsTrue(future.Result)

            Dim future2 As ThreadInterruptedResult = testee.RequestCallThreadInterrupted()

            ' 二度目はfalse
            Assert.IsFalse(future2.Result)

            testee.Shutdown()
            testee.Join()
        End Sub

        <Test()> Public Sub Interrupted_別スレッドが割り込まれていても現在のスレッドが割り込まれていなければfalse()
            Dim testee As New TestingThread
            testee.Start()
            testee.Interrupt()

            Assert.IsFalse(AThread.Interrupted)

            testee.Shutdown()
            testee.Join()
        End Sub

        <Test()> Public Sub IsInterrupted_スレッドが割り込まれていないならfalse()
            Dim testee As New TestingThread
            testee.Start()

            Assert.IsFalse(testee.IsInterrupted)

            testee.Shutdown()
            testee.Join()
        End Sub

        <Test()> Public Sub IsInterrupted_スレッドが割り込まれていればtrue()
            Dim testee As New TestingThread
            testee.Start()
            testee.Interrupt()

            Assert.IsTrue(testee.IsInterrupted)

            testee.Shutdown()
            testee.Join()
        End Sub

        <Test()> Public Sub IsInterrupted_割り込まれた後に別スレッドでInterruptedされてもこのスレッドの割り込みステータスはtrue()
            Dim testee As New TestingThread
            testee.Start()
            testee.Interrupt()

            Assert.IsFalse(AThread.Interrupted)

            Assert.IsTrue(testee.IsInterrupted)

            testee.Shutdown()
            testee.Join()
        End Sub

    End Class
End Namespace