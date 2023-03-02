Imports NUnit.Framework

Namespace Util
    Public MustInherit Class AbstractImmutablePairTest

        Public Class TestingAbstractImmutableStringPair : Inherits AbstractImmutablePair(Of String, String)

            Public ReadOnly Property Value1() As String
                Get
                    Return MyBase.PairA
                End Get
            End Property

            Public ReadOnly Property Value2() As String
                Get
                    Return MyBase.PairB
                End Get
            End Property

            Public Sub New(ByVal pair1 As String, ByVal pair2 As String)
                MyBase.New(pair1, pair2)
            End Sub
        End Class

        Public Class Equalsのテスト : Inherits AbstractImmutablePairTest

            <Test()> Public Sub 中身が同じなら_Equalsがtrue()
                Dim pair1 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair("a", "b")
                Dim pair2 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair("a", "b")
                Assert.IsTrue(pair1.Equals(pair2))
            End Sub

            <Test()> Public Sub Nullが含まれても_中身が同じなら_Equalsがtrue_両方Null()
                Dim pair1 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair(Nothing, Nothing)
                Dim pair2 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair(Nothing, Nothing)
                Assert.IsTrue(pair1.Equals(pair2))
            End Sub

            <Test(), Sequential()> Public Sub Nullが含まれても_中身が同じなら_Equalsがtrue_片側Null( _
                    <Values("a", Nothing)> ByVal value1 As String, <Values(Nothing, "b")> ByVal value2 As String, _
                    <Values("a", Nothing)> ByVal value3 As String, <Values(Nothing, "b")> ByVal value4 As String)
                Dim pair1 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair(value1, value2)
                Dim pair2 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair(value3, value4)
                Assert.IsTrue(pair1.Equals(pair2))
            End Sub

            <Test(), Sequential()> Public Sub 中身が違うなら_Equalsがfalse( _
                    <Values("a", "a2", "a", Nothing)> ByVal value1 As String, <Values("b", "b", "b", "b")> ByVal value2 As String, _
                    <Values("a", "a", "a", "a")> ByVal value3 As String, <Values("b2", "b", Nothing, "b")> ByVal value4 As String)
                Dim pair1 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair(value1, value2)
                Dim pair2 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair(value3, value4)
                Assert.IsFalse(pair1.Equals(pair2))
            End Sub

        End Class

        Public Class Equalsのテスト_相手がobject型 : Inherits AbstractImmutablePairTest

            <Test()> Public Sub 中身が同じならEqualsがtrue()
                Dim pair1 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair("a", "b")
                Dim pair2 As Object = New TestingAbstractImmutableStringPair("a", "b")
                Assert.IsTrue(pair1.Equals(pair2))
            End Sub

            <Test()> Public Sub 中身が同じ_nothing_でもEqualsがtrue()
                Dim pair1 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair(Nothing, Nothing)
                Dim pair2 As Object = New TestingAbstractImmutableStringPair(Nothing, Nothing)
                Assert.IsTrue(pair1.Equals(pair2))
            End Sub

            <Test()> Public Sub 中身が違うならEqualsがfalse()
                Dim pair1 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair("a", "b")
                Dim pair2 As Object = New TestingAbstractImmutableStringPair("a", "b2")
                Assert.IsFalse(pair1.Equals(pair2))
            End Sub

            <Test()> Public Sub 中身が違うならEqualsがfalse_その2()
                Dim pair1 As TestingAbstractImmutableStringPair = New TestingAbstractImmutableStringPair("a2", "b")
                Dim pair2 As Object = New TestingAbstractImmutableStringPair("a", "b")
                Assert.IsFalse(pair1.Equals(pair2))
            End Sub

        End Class

        Public Class Dictionaryでテスト : Inherits AbstractImmutablePairTest

            <Test()> Public Sub キーの中身が同じならContainsKeyがtrue()
                Dim hoge As New Dictionary(Of TestingAbstractImmutableStringPair, String)
                hoge.Add(New TestingAbstractImmutableStringPair("a", "b"), "c")
                Assert.IsTrue(hoge.ContainsKey(New TestingAbstractImmutableStringPair("a", "b")))
            End Sub

            <Test()> Public Sub キーの中身が同じ_nothing_でもContainsKeyがtrue()
                Dim hoge As New Dictionary(Of TestingAbstractImmutableStringPair, String)
                hoge.Add(New TestingAbstractImmutableStringPair(Nothing, Nothing), "c")
                Assert.IsTrue(hoge.ContainsKey(New TestingAbstractImmutableStringPair(Nothing, Nothing)))
            End Sub

            <Test()> Public Sub キーの中身が違うからContainsKeyがfalse()
                Dim hoge As New Dictionary(Of TestingAbstractImmutableStringPair, String)
                hoge.Add(New TestingAbstractImmutableStringPair("a", "b"), "c")
                Assert.IsFalse(hoge.ContainsKey(New TestingAbstractImmutableStringPair("a", "b2")))
            End Sub

            <Test()> Public Sub キーの中身が同じなら値を取得できる()
                Dim hoge As New Dictionary(Of TestingAbstractImmutableStringPair, String)
                hoge.Add(New TestingAbstractImmutableStringPair("a", "b"), "c")
                Assert.AreEqual("c", hoge(New TestingAbstractImmutableStringPair("a", "b")))
            End Sub

        End Class
    End Class
End Namespace