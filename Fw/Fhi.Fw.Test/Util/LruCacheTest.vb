Imports NUnit.Framework
Imports Fhi.Fw.Util.WorkerThread

Namespace Util
    ''' <summary>
    ''' LRU(Least Recently Used) CacheのFhi.Fw実装のテストクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class LruCacheTest

#Region "Fake classes..."
        Private Class FakeBehavior(Of TKey, TValue) : Implements LruCache(Of TKey, TValue).IBehavior
            Public Property TheWorld As DateTime = DateTime.Now
            Public Function GetNow() As Date Implements LruCache(Of TKey, TValue).IBehavior.GetNow
                Return TheWorld
            End Function
            Public Sub SleepThread(millisecond As Integer)
                TheWorld = TheWorld.Add(New TimeSpan(0, 0, 0, 0, millisecond))
            End Sub
        End Class
#End Region
        Private behavior As FakeBehavior(Of Integer, String)
        <SetUp()> Public Overridable Sub SetUp()
            behavior = New FakeBehavior(Of Integer, String)
        End Sub

        Public Class サイズ指定でTest : Inherits LruCacheTest
            <Test()> Public Sub サイズ指定テスト_サイズ3のLRUキャッシュを利用できる()
                Dim cache As New LruCache(Of Integer, String)(TimeSpan.MaxValue, maxCacheSize:=3)
                cache.Add(1, "ほげ")
                cache.Add(2, "ふが")
                cache.Add(3, "ぴよ")

                cache.Add(4, "むさ")
                Assert.That(cache.Get(1), [Is].Null, "最大サイズ3なので最初のキャッシュが消える")
                Assert.That(cache.Get(2), [Is].EqualTo("ふが"))
                Assert.That(cache.Get(3), [Is].EqualTo("ぴよ"))
                Assert.That(cache.Get(4), [Is].EqualTo("むさ"), "とてもおいしい")
            End Sub

            <Test()> Public Sub サイズ指定テスト_最も最近使われていないものがLRUキャッシュから削除される_その1()
                Dim cache As New LruCache(Of Integer, String)(TimeSpan.MaxValue, maxCacheSize:=3)
                cache.Add(1, "ほげ")
                cache.Add(2, "ふが")
                cache.Add(3, "ぴよ")

                Assert.That(cache.Get(1), [Is].EqualTo("ほげ"), "「ほげ」のアクセス数が1になる")

                cache.Add(4, "むさ")
                Assert.That(cache.Get(1), [Is].EqualTo("ほげ"))
                Assert.That(cache.Get(2), [Is].Null, "アクセス数0の中で一番古い「ふが」が削除された")
                Assert.That(cache.Get(3), [Is].EqualTo("ぴよ"))
                Assert.That(cache.Get(4), [Is].EqualTo("むさ"))
            End Sub

            <Test()> Public Sub サイズ指定テスト_最も最近使われていないものがLRUキャッシュから削除される_その2()
                Dim cache As New LruCache(Of Integer, String)(TimeSpan.MaxValue, maxCacheSize:=3)
                cache.Add(1, "ほげ")
                cache.Add(2, "ふが")
                cache.Add(3, "ぴよ")

                Assert.That(cache.Get(1), [Is].EqualTo("ほげ"), "「ほげ」のアクセス数が1になる")
                Assert.That(cache.Get(1), [Is].EqualTo("ほげ"), "「ほげ」のアクセス数が2になる")
                Assert.That(cache.Get(2), [Is].EqualTo("ふが"), "「ふが」のアクセス数が1になる")

                cache.Add(4, "むさ")
                Assert.That(cache.Get(1), [Is].EqualTo("ほげ"))
                Assert.That(cache.Get(2), [Is].EqualTo("ふが"))
                Assert.That(cache.Get(3), [Is].Null, "最もアクセス数の少ない「ぴよ」が削除された")
                Assert.That(cache.Get(4), [Is].EqualTo("むさ"))
            End Sub
        End Class

        Public Class intervalでTest : Inherits LruCacheTest
            <Test()> Public Sub intervalでクリアされる前でもexpiredされるので_Getできない()
                Dim cache As New LruCache(Of Integer, String)(expiryTimeoutMillis:=20, memoryRefreshInterval:=500, behavior:=behavior)
                cache.Add(1, "むさむさ")

                behavior.SleepThread(10)
                Assert.That(cache.Get(1), [Is].EqualTo("むさむさ"), "expiry前なので取得できる")

                behavior.SleepThread(20)
                Assert.That(cache.Get(1), [Is].Null, "interval前だけどexpiredだから取得できない")
            End Sub

            <Test()> Public Sub intervalでクリアされる前に_同一keyでAddされても_expiryによって取得はできる()
                Dim cache As New LruCache(Of Integer, String)(expiryTimeoutMillis:=20, memoryRefreshInterval:=500, behavior:=behavior)
                cache.Add(1, "むさむさ")

                behavior.SleepThread(30)
                Assert.That(cache.Get(1), [Is].Null, "interval前だけどexpiredだから取得できない")

                cache.Add(1, "ほげ")
                behavior.SleepThread(10)
                Assert.That(cache.Get(1), [Is].EqualTo("ほげ"), "expiry前なので取得できる")

                behavior.SleepThread(20)
                Assert.That(cache.Get(1), [Is].Null, "interval前だけどexpiredだから取得できない")
            End Sub
        End Class

        Public Class MultiThreadingTest : Inherits LruCacheTest

            Private r As Random
            Private sut As LruCache(Of Integer, String)

            <SetUp()> Public Overrides Sub SetUp()
                MyBase.SetUp()
                r = New Random
                sut = New LruCache(Of Integer, String)(New TimeSpan(0, 0, 0, 10))
            End Sub

            <Test()> Public Sub マルチスレッドでも利用できる()
                Const TRY_COUNT As Integer = 1000
                Dim aChannel As New Channel(threadCount:=23)
                aChannel.StartWorkers()
                Try
                    For i As Integer = 0 To TRY_COUNT - 1
                        aChannel.RegisterTask(AddressOf PerformAdd)
                    Next
                Finally
                    aChannel.StopWorkersAfterAllTasksDone()
                End Try
            End Sub

            Private Sub PerformAdd()
                Dim i As Integer = r.Next(1000)
                sut.Add(i, i.ToString())
                Threading.Thread.Sleep(5)
                Dim actual As String = sut.Get(i)
                Assert.That(actual, [Is].EqualTo(i.ToString))
            End Sub

        End Class

    End Class
End Namespace