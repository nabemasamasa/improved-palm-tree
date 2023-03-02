Imports NUnit.Framework
Imports System.Threading
Imports System.Text.RegularExpressions

Namespace Util.WorkerThread
    Public MustInherit Class ChannelTest

#Region "Task Classes..."
        Private MustInherit Class AbstractTask
            Protected ReadOnly messages As Log
            Public Sub New(ByVal messages As Log)
                Me.messages = messages
            End Sub
        End Class

        Private Class TaskA : Inherits AbstractTask : Implements ITask
            Public Sub New(ByVal messages As Log)
                MyBase.New(messages)
            End Sub
            Public Sub Execute() Implements ITask.Execute
                Threading.Thread.Sleep(70)
                messages.Add(String.Format("{0} {1} is done.", Thread.CurrentThread.Name, GetType(TaskA).Name))
            End Sub
        End Class

        Private Class TaskB : Inherits AbstractTask : Implements ITask
            Public Sub New(ByVal messages As Log)
                MyBase.New(messages)
            End Sub
            Public Sub Execute() Implements ITask.Execute
                Threading.Thread.Sleep(30)
                messages.Add(String.Format("{0} {1} is done.", Thread.CurrentThread.Name, GetType(TaskB).Name))
            End Sub
        End Class

        Private Class TaskC : Inherits AbstractTask : Implements ITask
            Private ReadOnly name As String
            Public Sub New(ByVal messages As Log, ByVal name As String)
                MyBase.New(messages)
                Me.name = name
            End Sub
            Public Sub Execute() Implements ITask.Execute
                Threading.Thread.Sleep(20)
                messages.Add(String.Format("{0} {1} is done.", Thread.CurrentThread.Name, name))
            End Sub
        End Class
#End Region

#Region "Delegates..."
        Private Delegate Sub DelegateA()
        Private Delegate Sub DelegateB()
        Private Delegate Sub DelegateC(ByVal name As String)

        Private Sub DelegateASub()
            Threading.Thread.Sleep(70)
            messages.Add(String.Format("{0} {1} is done.", Thread.CurrentThread.Name, GetType(DelegateA).Name))
        End Sub
        Private Sub DelegateBSub()
            Threading.Thread.Sleep(30)
            messages.Add(String.Format("{0} {1} is done.", Thread.CurrentThread.Name, GetType(DelegateB).Name))
        End Sub
        Private Sub DelegateCSub(ByVal name As String)
            Threading.Thread.Sleep(10)
            messages.Add(String.Format("{0} {1} is done.", Thread.CurrentThread.Name, name))
        End Sub
        Private Class Log
            Private ReadOnly messages As List(Of String)
            Private ReadOnly syncRoot As Object
            Public Sub New()
                messages = New List(Of String)
                syncRoot = DirectCast(messages, ICollection).SyncRoot
            End Sub
            Public Sub Add(ByVal value As String)
                SyncLock syncRoot
                    messages.Add(value)
                End SyncLock
            End Sub

            Public ReadOnly Property Count() As Integer
                Get
                    Return messages.Count
                End Get
            End Property

            Default Public Property Item(ByVal index As Integer) As String
                Get
                    Return messages.Item(index)
                End Get
                Set(ByVal value As String)
                    messages.Item(index) = value
                End Set
            End Property

            Public Function ToArray() As String()
                Return messages.ToArray()
            End Function
        End Class
#End Region

        Private messages As Log

        <SetUp()> Public Overridable Sub SetUp()
            messages = New Log
        End Sub

        Public Class スレッド数1のチャンネルTest : Inherits ChannelTest

            Private aChannel As Channel

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                aChannel = New Channel
                aChannel.StartWorkers()
            End Sub

            <TearDown()> Public Sub TearDown()
                aChannel.StopWorkers()
            End Sub

            <Test()> Public Sub スレッド数が1のチャンネルであれば_タスクは追加された順に実行される()
                aChannel.RegisterTask(New TaskA(messages))
                aChannel.RegisterTask(New TaskB(messages))
                aChannel.RegisterTask(New TaskC(messages, GetType(TaskC).Name))

                aChannel.WaitUntilTaskDone()

                Assert.That(messages.Count, [Is].EqualTo(3))
                Assert.That(messages(0), [Is].EqualTo("Worker-0 TaskA is done."))
                Assert.That(messages(1), [Is].EqualTo("Worker-0 TaskB is done."))
                Assert.That(messages(2), [Is].EqualTo("Worker-0 TaskC is done."))
            End Sub

            <Test()> Public Sub スレッド数が1のチャンネルであれば_デリゲートは追加された順に実行される_正式な書き方の場合()
                aChannel.RegisterTask(New DelegateB(AddressOf DelegateASub))
                aChannel.RegisterTask(New DelegateB(AddressOf DelegateBSub))
                aChannel.RegisterTask(New DelegateC(AddressOf DelegateCSub), GetType(DelegateC).Name)

                aChannel.WaitUntilTaskDone()

                Assert.That(messages.Count, [Is].EqualTo(3))
                Assert.That(messages(0), [Is].EqualTo("Worker-0 DelegateA is done."))
                Assert.That(messages(1), [Is].EqualTo("Worker-0 DelegateB is done."))
                Assert.That(messages(2), [Is].EqualTo("Worker-0 DelegateC is done."))
            End Sub

            <Test()> Public Sub スレッド数が1のチャンネルであれば_デリゲートは追加された順に実行される_MethodInvoker形式の場合()
                aChannel.RegisterTask(AddressOf DelegateASub)
                aChannel.RegisterTask(AddressOf DelegateBSub)

                aChannel.WaitUntilTaskDone()

                Assert.That(messages.Count, [Is].EqualTo(2))
                Assert.That(messages(0), [Is].EqualTo("Worker-0 DelegateA is done."))
                Assert.That(messages(1), [Is].EqualTo("Worker-0 DelegateB is done."))
            End Sub
        End Class

        Public Class スレッド数3のチャンネルTest : Inherits ChannelTest

            Private aChannel As Channel

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                aChannel = New Channel(threadCount:=3)
                aChannel.StartWorkers()
            End Sub

            <TearDown()> Public Sub TearDown()
                aChannel.StopWorkers()
            End Sub

            <Test()> Public Sub タスクは追加した順序で必ずしも実行されないし_終了は同期しない()
                aChannel.RegisterTask(New TaskA(messages))
                aChannel.RegisterTask(New TaskB(messages))
                aChannel.RegisterTask(New TaskC(messages, GetType(TaskC).Name))

                aChannel.WaitUntilTaskDone()

                Assert.That(messages.Count, [Is].EqualTo(3))
                Dim expandMessage As String = Join(messages.ToArray, ",")
                Assert.That(expandMessage.Contains("TaskA is done."), [Is].True, expandMessage)
                Assert.That(expandMessage.Contains("TaskB is done."), [Is].True, expandMessage)
                Assert.That(expandMessage.Contains("TaskC is done."), [Is].True, expandMessage)

                Dim matches As MatchCollection = New Regex("Worker-\d Task").Matches(expandMessage)
                Assert.That(matches.Count, [Is].EqualTo(3))
            End Sub

        End Class

    End Class
End Namespace