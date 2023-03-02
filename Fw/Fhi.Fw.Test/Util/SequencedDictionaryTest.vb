Imports NUnit.Framework

Namespace Util
    Public MustInherit Class SequencedDictionaryTest

        Private Function ConvEnumeratorToList(dictionary As IDictionary(Of String, Integer)) As ICollection(Of KeyValuePair(Of String, Integer))
            ' dictionary.ToList にするべからず
            Dim keyValuePairs As New List(Of KeyValuePair(Of String, Integer))
            For Each keyValuePair As KeyValuePair(Of String, Integer) In dictionary
                keyValuePairs.Add(keyValuePair)
            Next
            Return keyValuePairs
        End Function

        Public Class 順序を保持できるTest : Inherits SequencedDictionaryTest

            Private sut As IDictionary(Of String, Integer)

            <SetUp()> Public Sub SetUp()
                sut = New SequencedDictionary(Of String, Integer)
                'sut = New Dictionary(Of String, Integer)   ' 標準Dictionaryで動作確認するならコレ
                sut.Add("foo", 16)
                sut.Add("bar", 72)
                sut.Add("baz", 42)
            End Sub

            <Test()> Public Sub Add_しただけ()
                Assert.That(sut.Keys.ToArray, [Is].EqualTo({"foo", "bar", "baz"}))
                Assert.That(sut.Values.ToArray, [Is].EqualTo({16, 72, 42}))
            End Sub

            <Test()> Public Sub Add_さらに追加()
                sut.Add("qux", 23)
                Assert.That(sut.Keys.ToArray, [Is].EqualTo({"foo", "bar", "baz", "qux"}))
                Assert.That(sut.Values.ToArray, [Is].EqualTo({16, 72, 42, 23}))
            End Sub

            <Test()> Public Sub Add_KeyValuePairのAddを混ぜても保持()
                sut.Clear()
                sut.Add("qux", 23)
                sut.Add(New KeyValuePair(Of String, Integer)("foo", 12))
                sut.Add(New KeyValuePair(Of String, Integer)("bar", 103))
                Assert.That(sut.Keys.ToArray, [Is].EqualTo({"qux", "foo", "bar"}))
                Assert.That(sut.Values.ToArray, [Is].EqualTo({23, 12, 103}))
            End Sub

            <Test()> Public Sub Remove_しただけ()
                sut.Remove("bar")
                Assert.That(sut.Keys.ToArray, [Is].EqualTo({"foo", "baz"}))
                Assert.That(sut.Values.ToArray, [Is].EqualTo({16, 42}))
            End Sub

            <TestCase(36)>
            <TestCase(102)>
            Public Sub SetValue_で値変更しても保持(value As Integer)
                sut("bar") = value
                Assert.That(sut("bar"), [Is].EqualTo(value))
                Assert.That(sut.Keys.ToArray, [Is].EqualTo({"foo", "bar", "baz"}))
                Assert.That(sut.Values.ToArray, [Is].EqualTo({16, value, 42}))
            End Sub

            <Test()> Public Sub Remove_Add_したら_新しい順序()
                sut.Remove("bar")
                sut.Add("bar", 24)
                Assert.That(sut.Keys.ToArray, [Is].EqualTo({"foo", "baz", "bar"}))
                Assert.That(sut.Values.ToArray, [Is].EqualTo({16, 42, 24}))
            End Sub

            <Test()> Public Sub Remove_Add_した時の_GetEnumeratorも順序を保持できる()
                sut.Remove("bar")
                sut.Add("bar", 37)
                Dim keyValuePairs As ICollection(Of KeyValuePair(Of String, Integer)) = ConvEnumeratorToList(sut)
                Assert.That(keyValuePairs(0), [Is].EqualTo(New KeyValuePair(Of String, Integer)("foo", 16)))
                Assert.That(keyValuePairs(1), [Is].EqualTo(New KeyValuePair(Of String, Integer)("baz", 42)))
                Assert.That(keyValuePairs(2), [Is].EqualTo(New KeyValuePair(Of String, Integer)("bar", 37)))
            End Sub

            <Test()> Public Sub Remove_Add_した時の_順序を保持したCopyToをできる()
                sut.Remove("bar")
                sut.Add("bar", 37)
                Dim array As KeyValuePair(Of String, Integer)() = New KeyValuePair(Of String, Integer)() {Nothing, Nothing, Nothing}
                sut.CopyTo(array, 0)
                Assert.That(array(0), [Is].EqualTo(New KeyValuePair(Of String, Integer)("foo", 16)))
                Assert.That(array(1), [Is].EqualTo(New KeyValuePair(Of String, Integer)("baz", 42)))
                Assert.That(array(2), [Is].EqualTo(New KeyValuePair(Of String, Integer)("bar", 37)))
            End Sub

            <Test()> Public Sub Count()
                Assert.That(sut.Count, [Is].EqualTo(3))
            End Sub

            <Test()> Public Sub Clear()
                sut.Clear()
                Assert.That(sut, [Is].Empty)
                Assert.That(sut.Keys, [Is].Empty)
                Assert.That(sut.Values, [Is].Empty)
            End Sub

        End Class

    End Class
End Namespace