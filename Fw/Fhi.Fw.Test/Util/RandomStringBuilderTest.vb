Imports NUnit.Framework

Namespace Util
    ''' <summary>
    ''' ランダムな文字列を生成するクラスのテストクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class RandomStringBuilderTest

        Private builder As RandomStringBuilder

        <SetUp()> Public Overridable Sub SetUp()
            builder = New RandomStringBuilder
        End Sub

        Public Class DefaultTest : Inherits RandomStringBuilderTest

            <Test()> Public Sub ランダムな文字列を生成できる()
                Dim actual As String = (New RandomStringBuilder).CreateRandomString(length:=4)
                Assert.That(actual.Length, [Is].EqualTo(4))
            End Sub

            <Test()> Public Sub 桁数を指定してランダムな文字列を生成できる(<Values(1, 4, 8)> ByVal length As Integer)
                Dim random1 As String = builder.CreateRandomString(length)
                Dim random2 As String = builder.CreateRandomString(length)
                Dim random3 As String = builder.CreateRandomString(length)
                Assert.That(random1, [Is].Not.EqualTo(random2))
                Assert.That(random2, [Is].Not.EqualTo(random3))
                Assert.That(random3, [Is].Not.EqualTo(random1))
            End Sub

            <Test()> _
            Public Sub 桁数が1桁なら62文字目まで重複なく文字列を生成できる()
                Dim actual As String() = Enumerable.Range(1, 62).Select(Function(i) builder.CreateRandomString(1)).ToArray
                Assert.That(actual.Count, [Is].EqualTo(62))
                Assert.That(actual.Count, [Is].EqualTo(actual.Distinct().Count))
            End Sub

            <Test()> _
            Public Sub 桁数が1桁なら62文字なので63文字目で例外になる()
                Try
                    For i As Integer = 1 To 63
                        builder.CreateRandomString(1)
                    Next
                    Assert.Fail("62文字と1桁なら63文字目で重複が発生しなければおかしい")
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("これ以上1桁で重複のない文字列を抽出できません"))
                    Assert.Pass()
                End Try
            End Sub

            <Test()> _
            Public Sub インスタンス毎に出力される文字に重複はない()
                Dim actual As String() = Enumerable.Range(1, 62).Select(Function(i) builder.CreateRandomString(1)).ToArray
                builder = New RandomStringBuilder
                Dim actual2 As String() = Enumerable.Range(1, 62).Select(Function(i) builder.CreateRandomString(1)).ToArray
                Assert.That(Join(actual, ""), [Is].Not.EqualTo(Join(actual2, "")))
            End Sub
        End Class

    End Class
End Namespace
