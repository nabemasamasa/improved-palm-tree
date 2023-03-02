Imports NUnit.Framework

Namespace Util
    Public MustInherit Class TimeoutDictionaryTest

        Private sut As TimeoutDictionary(Of String, Integer)

        <SetUp()> Public Overridable Sub SetUp()
            sut = New TimeoutDictionary(Of String, Integer)
        End Sub

        Public Class TimeoutTest : Inherits TimeoutDictionaryTest

            <Test()> Public Sub _1000msなら_50ms待機しても_Timeout前だから_値がある()
                Const KEY As String = "A"
                sut.Add(KEY, 1)
                sut.SetTimeout(1000)
                Threading.Thread.Sleep(50)
                Assert.That(sut.ContainsKey(KEY), [Is].True)
            End Sub

            <Test()> Public Sub _1msなら_50ms待機したら_Timeout後だから_値をclearしてる()
                Const KEY As String = "B"
                sut.Add(KEY, 1)
                sut.SetTimeout(1)
                Threading.Thread.Sleep(50)
                Assert.That(sut.ContainsKey(KEY), [Is].False)
            End Sub

            <Test()> Public Sub _50msなら_期間100msの間に破棄される()
                Const KEY As String = "C"
                sut.Add(KEY, 1)
                Dim timer As New Stopwatch
                timer.Start()
                sut.SetTimeout(50)
                Dim isAlive As Boolean = True
                Dim clearedMillis As Long = 0
                While timer.ElapsedMilliseconds < 100
                    If isAlive Then
                        If Not sut.ContainsKey(KEY) Then
                            isAlive = False
                            clearedMillis = timer.ElapsedMilliseconds
                        End If
                    Else
                        Assert.That(sut.ContainsKey(KEY), [Is].False, "clearされた後、復活する訳ない")
                    End If
                    Threading.Thread.Sleep(1)
                End While
                Debug.Print(clearedMillis & "ms")
                Assert.That(clearedMillis, [Is].GreaterThanOrEqualTo(50), "だいたい58～85msになる")
                Assert.That(isAlive, [Is].False)
            End Sub

        End Class

    End Class
End Namespace