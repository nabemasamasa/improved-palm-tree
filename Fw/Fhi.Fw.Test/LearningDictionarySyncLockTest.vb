Imports NUnit.Framework
Imports System.Threading

Public Class LearningDictionarySyncLockTest

    Private Class TestingBaseDictionary
        Private Const KEY As String = "hoge"
        Protected map As Dictionary(Of String, Integer)
        Public Sub New()
            map = New Dictionary(Of String, Integer)
            map.Add(KEY, 2000)
        End Sub

        Public Overridable Sub Read()
            If Not map.ContainsKey(KEY) Then
                Debug.Print("not Contains 'hoge'")
            End If
        End Sub

        Public Overridable Sub Write()
            Dim value As Integer
            If map.ContainsKey(KEY) Then
                value = map(KEY)
                map.Remove(KEY)
            End If
            value += 100
            map.Add(KEY, value)
            Thread.Sleep(99)
            value -= 100
            map.Remove(KEY)
            map.Add(KEY, value)
        End Sub
    End Class

    Private Class TestingDictionaryNecessary : Inherits TestingBaseDictionary
        ''' <summary>引数なしSubメソッドのdelegate</summary>
        ''' <remarks>System.Windows.Formsを参照していないプロジェクトに必要</remarks>
        Private Delegate Sub MethodInvoker()

        Public ExceptionCounter As Integer = 0
        Private ReadOnly allowExceptionMessages As New List(Of String)

        Public Sub New(ByVal ParamArray allowExceptionMessages As String())
            MyBase.New()
            Me.allowExceptionMessages.AddRange(allowExceptionMessages)
        End Sub

        Private Sub FilterException(ByVal invoker As MethodInvoker)
            Try
                invoker.Invoke()
            Catch ex As Exception
                If Not allowExceptionMessages.Contains(ex.Message) Then
                    Throw
                End If
                Interlocked.Add(ExceptionCounter, 1)
            End Try
        End Sub

        Public Overrides Sub Read()
            FilterException(AddressOf MyBase.Read)
        End Sub

        Public Overrides Sub Write()
            FilterException(AddressOf MyBase.Write)
        End Sub
    End Class

    Private Class TestingReaderWriterLock : Inherits TestingBaseDictionary
        Private rwLock As New ReaderWriterLock
        Public Sub New()
            MyBase.New()
        End Sub

        Public Overrides Sub Read()
            'Console.WriteLine("Read() " & Thread.CurrentThread.ManagedThreadId)
            rwLock.AcquireReaderLock(Timeout.Infinite)
            Try
                MyBase.Read()
            Finally
                rwLock.ReleaseReaderLock()
            End Try
        End Sub

        Public Overrides Sub Write()
            'Console.WriteLine("Write() " & Thread.CurrentThread.ManagedThreadId)
            rwLock.AcquireWriterLock(Timeout.Infinite)
            Try
                MyBase.Write()
            Finally
                rwLock.ReleaseWriterLock()
            End Try
        End Sub
    End Class

    Private Class TestingSyncLockAtMapInstance : Inherits TestingBaseDictionary
        Public Sub New()
            MyBase.New()
        End Sub

        Public Overrides Sub Read()
            SyncLock map
                MyBase.Read()
            End SyncLock
        End Sub

        Public Overrides Sub Write()
            'Console.WriteLine("Write() " & Thread.CurrentThread.ManagedThreadId)
            SyncLock map
                MyBase.Write()
            End SyncLock
        End Sub
    End Class

    <Test(), Repeat(30), Ignore("t.Joinでdead lockする時があり、解消するまでignore")> Public Sub ロック制御しない場合_競合するので_キー重複やキー無しエラーになる()
        Dim sss As New TestingDictionaryNecessary("同一のキーを含む項目が既に追加されています。", "インデックスが配列の境界外です。", "指定されたキーはディレクトリ内に存在しませんでした。")
        Dim threads As New List(Of Thread)
        For i As Integer = 0 To 9
            threads.Add(New Thread(New ThreadStart(AddressOf sss.Read)))
            threads.Add(New Thread(New ThreadStart(AddressOf sss.Write)))
        Next
        For Each t As Thread In threads
            t.Start()
        Next
        For Each t As Thread In threads
            t.Join()
        Next
        If 0 < sss.ExceptionCounter Then
            Debug.Print("これは競合するテストだから、例外が " & sss.ExceptionCounter & " 回発生した")
        End If
    End Sub

    <Test()> Public Sub ReaderWriterLockを使用した場合()
        Dim sss As New TestingReaderWriterLock()
        Dim threads As New List(Of Thread)
        For i As Integer = 0 To 9
            threads.Add(New Thread(New ThreadStart(AddressOf sss.Read)))
            threads.Add(New Thread(New ThreadStart(AddressOf sss.Write)))
        Next
        For Each t As Thread In threads
            t.Start()
        Next
        For Each t As Thread In threads
            t.Join()
        Next
    End Sub

    <Test()> Public Sub DictionaryインスタンスをSyncLockした場合()
        Dim sss As New TestingSyncLockAtMapInstance()
        Dim threads As New List(Of Thread)
        For i As Integer = 0 To 9
            threads.Add(New Thread(New ThreadStart(AddressOf sss.Read)))
            threads.Add(New Thread(New ThreadStart(AddressOf sss.Write)))
        Next
        For Each t As Thread In threads
            t.Start()
        Next
        For Each t As Thread In threads
            t.Join()
        Next
    End Sub

End Class
