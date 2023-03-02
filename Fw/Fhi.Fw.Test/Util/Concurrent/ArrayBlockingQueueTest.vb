Imports NUnit.Framework
Imports System.Threading
Imports Fhi.Fw.Lang.Threading

Namespace Util.Concurrent
    Public Class ArrayBlockingQueueTest

        Private Class PutThread : Inherits AThread

            Private ReadOnly queue As ArrayBlockingQueue(Of String)
            Private ReadOnly putCount As Integer
            Public Ex As ThreadInterruptedException
            Public ReadOnly PutValues As New List(Of String)

            Public Sub New(ByVal queue As ArrayBlockingQueue(Of String), ByVal putCount As Integer)
                Me.queue = queue
                Me.putCount = putCount
            End Sub

            Public Overrides Sub Run()
                Try
                    Dim hoge As New Random(1234)
                    For i As Integer = 0 To putCount - 1
                        Dim value As String = String.Format("{0}", hoge.Next)
                        queue.Put(value)
                        PutValues.Add(value)
                    Next
                Catch ex As ThreadInterruptedException
                    Me.Ex = ex
                End Try
            End Sub
        End Class

        Private Class TakeThread : Inherits AThread

            Private ReadOnly queue As ArrayBlockingQueue(Of String)
            Private ReadOnly takeCount As Integer
            Public Ex As ThreadInterruptedException
            Public ReadOnly TookValues As New List(Of String)

            Public Sub New(ByVal queue As ArrayBlockingQueue(Of String), ByVal takeCount As Integer)
                Me.queue = queue
                Me.takeCount = takeCount
            End Sub

            Public Overrides Sub Run()
                Try
                    For i As Integer = 0 To takeCount - 1
                        TookValues.Add(queue.Take)
                    Next
                Catch ex As ThreadInterruptedException
                    Me.Ex = ex
                End Try
            End Sub
        End Class

        Private Function IsState(ByVal th As AThread, ByVal state As ThreadState) As Boolean
            Return (th.ThreadState And state) = state
        End Function

        <Test()> Public Sub Take_インスタンス生成直後_Put値が無いからTakeはブロックし続ける()

            Dim queue As New ArrayBlockingQueue(Of String)(3)
            Dim taker As New TakeThread(queue, 1)

            taker.Start()
            While Not IsState(taker, ThreadState.WaitSleepJoin)
                ' taker#Runの queue.Take()のなかで、Monitor.Waitされるまで待機
                Thread.Sleep(10)
            End While

            taker.Interrupt()
            While Not IsState(taker, ThreadState.Stopped)
                ' InterruptException発生まで待機
                Thread.Sleep(10)
            End While

            Assert.AreEqual(0, taker.TookValues.Count, "Put値がないから値は取れない")
            Assert.IsNotNull(taker.Ex, "Wait中にInterruptされたのだから例外が発生のハズ")
        End Sub

        <Test()> Public Sub Put_サイズを超えてPutできない_Takeされるまでブロックし続ける()

            Dim queue As New ArrayBlockingQueue(Of String)(capacity:=1)
            Dim putter As New PutThread(queue, Integer.MaxValue)

            putter.Start()
            While Not IsState(putter, ThreadState.WaitSleepJoin)
                ' putter#Runの queue.Put()のなかで、Monitor.Waitされるまで待機
                Thread.Sleep(10)
            End While

            putter.Interrupt()
            While Not IsState(putter, ThreadState.Stopped)
                ' InterruptException発生まで待機
                Thread.Sleep(10)
            End While

            Assert.AreEqual(1, putter.PutValues.Count, "二つ目はTakeされるまでブロックされ続ける")
            Assert.IsNotNull(putter.Ex, "Wait中にInterruptされたのだから例外が発生のハズ")
        End Sub

        <Test()> Public Sub Take起動も値が無いからPutが先に処理される()

            Dim queue As New ArrayBlockingQueue(Of String)(1)
            Dim putter As New PutThread(queue, 1)
            Dim taker As New TakeThread(queue, 1)

            taker.Start()
            Thread.Sleep(10)

            putter.Start()
            Thread.Sleep(10)

            Assert.AreEqual(1, taker.TookValues.Count)
            Assert.AreEqual(1, putter.PutValues.Count)
            Assert.AreEqual(putter.PutValues(0), taker.TookValues(0), "Putした値をTakeしたハズ")
        End Sub
    End Class
End Namespace