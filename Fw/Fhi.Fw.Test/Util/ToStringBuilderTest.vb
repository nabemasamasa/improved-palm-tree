Imports NUnit.Framework
Imports Fhi.Fw.Util

Namespace Util
    Public Class ToStringBuilderTest
        Private Class NumNameVo
            Private _num As Integer?
            Private _name As String
        End Class

        <Test()> Public Sub ToString_文字列はシングルクォートで囲む()
            Dim builder As New ToStringBuilder
            builder.Add("hoge", "fuga")
            Assert.AreEqual("hoge='fuga'", builder.ToString)
        End Sub

        <Test()> Public Sub ToString_文字列はシングルクォートで囲む_二つ()
            Dim builder As New ToStringBuilder
            builder.Add("hoge", "fuga")
            builder.Add("piyo", "poo")
            Assert.AreEqual("hoge='fuga', piyo='poo'", builder.ToString)
        End Sub

        <Test()> Public Sub ToString_数値はシングルクォートで囲まない()
            Dim builder As New ToStringBuilder
            builder.Add("hoge", 12)
            Assert.AreEqual("hoge=12", builder.ToString)
        End Sub

        <Test()> Public Sub ToString_Nullable数値もシングルクォートで囲まない()
            Dim num As Integer? = 13
            Dim builder As New ToStringBuilder
            builder.Add("hoge", num)
            Assert.AreEqual("hoge=13", builder.ToString)
        End Sub

        <Test()> Public Sub ToString_配列は中括弧で囲まれて_要素indexも付与される()
            Dim builder As New ToStringBuilder
            builder.Add("hoge", New String() {"fuga", "piyo"})
            Assert.AreEqual("hoge={(0)='fuga', (1)='piyo'}", builder.ToString)
        End Sub

        <Test()> Public Sub ToString_Listも中括弧で囲まれて_要素indexも付与される()
            Dim builder As New ToStringBuilder
            builder.Add("hoge", EzUtil.NewList("fuga", "piyo"))
            Assert.AreEqual("hoge={(0)='fuga', (1)='piyo'}", builder.ToString)
        End Sub

        <Test()> Public Sub ToString_配列_文字列()
            Dim builder As New ToStringBuilder
            builder.Add("hoge", New String() {"fuga", Nothing})
            Assert.AreEqual("hoge={(0)='fuga', (1)=<nothing>}", builder.ToString)
        End Sub

        <Test()> Public Sub ToString_配列_数値()
            Dim builder As New ToStringBuilder
            builder.Add("hoge", New Integer() {12, 34})
            Assert.AreEqual("hoge={(0)=12, (1)=34}", builder.ToString)
        End Sub

        <Test()> Public Sub ToString_配列_Nullable数値()
            Dim builder As New ToStringBuilder
            builder.Add("hoge", New Integer?() {Nothing, 123})
            Assert.AreEqual("hoge={(0)=<nothing>, (1)=123}", builder.ToString)
        End Sub

        <Test()> Public Sub ToString_配列のネスト()
            Dim builder As New ToStringBuilder
            builder.Add("hoge", New Object() {"fuga", New Object() {"piyo", "poo"}})
            Assert.AreEqual("hoge={(0)='fuga', (1)={(0)='piyo', (1)='poo'}}", builder.ToString)
        End Sub


    End Class
End Namespace