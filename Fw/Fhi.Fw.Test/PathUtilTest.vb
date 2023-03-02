Imports NUnit.Framework

''' <summary>
''' パスに関するユーティリティクラスのテストクラス
''' </summary>
''' <remarks></remarks>
Public MustInherit Class PathUtilTest

    Public Class CombineTest : Inherits PathUtilTest

        <Test()> Public Sub 引数がnullの場合_例外になる()
            Try
                Dim ignore As String = PathUtil.Combine(Nothing)
                Assert.Fail()
            Catch expected As ArgumentNullException
                Assert.That(expected.Message, [Is].StringContaining("値を Null にすることはできません").Or.StringContaining("Value cannot be null."))
            End Try
        End Sub

        <Test()> Public Sub 引数が長さ0のコレクションの場合_例外になる()
            Try
                Dim ignore As String = PathUtil.Combine(New String() {})
                Assert.Fail()
            Catch expected As ArgumentNullException
                Assert.That(expected.Message, [Is].StringContaining("値を Null にすることはできません").Or.StringContaining("Value cannot be null."))
            End Try
        End Sub

        <Test()> Public Sub 値が空文字の場合_空文字を返す()
            Assert.That(PathUtil.Combine(String.Empty), [Is].Empty)
        End Sub

        <Test()> Public Sub 任意数のパスを結合できる()
            Assert.That(PathUtil.Combine("hoge"), [Is].EqualTo("hoge"))
            Assert.That(PathUtil.Combine("hoge", "fuga"), [Is].EqualTo("hoge\fuga"))
            Assert.That(PathUtil.Combine("hoge", "fuga", "piyo\musa"), [Is].EqualTo("hoge\fuga\piyo\musa"))
        End Sub

        <Test()> Public Sub 途中にnullがある場合_例外になる()
            Try
                Dim ignore As String = PathUtil.Combine(PathUtil.Combine("hoge", "fuga", Nothing, "musa"))
                Assert.Fail()
            Catch expected As ArgumentNullException
                Assert.That(expected.Message, [Is].StringContaining("値を Null にすることはできません").Or.StringContaining("Value cannot be null."))
            End Try
        End Sub

        <Test()> Public Sub 途中に空文字がある場合_無視されて連結される()
            Assert.That(PathUtil.Combine(PathUtil.Combine("hoge", "fuga", "", "musa")), [Is].EqualTo("hoge\fuga\musa"))
            Assert.That(PathUtil.Combine(PathUtil.Combine("hoge", "", "fuga", "", "musa")), [Is].EqualTo("hoge\fuga\musa"))
        End Sub

    End Class

    Public Class OmitPathTest : Inherits PathUtilTest

        <TestCase("c:\hoge\fuga\piyo.txt", 100, "c:\hoge\fuga\piyo.txt")>
        <TestCase("c:\hoge\fuga\piyo.txt", 21, "c:\hoge\fuga\piyo.txt")>
        <TestCase("c:\hoge\fuga\piyo.txt", 20, "c:\...\fuga\piyo.txt")>
        <TestCase("c:\hoge\fuga\piyo.txt", 19, "c:\...\piyo.txt")>
        <TestCase("c:\hoge\fuga\piyo.txt", 15, "c:\...\piyo.txt")>
        <TestCase("c:\hoge\fuga\piyo.txt", 14, "...\piyo.txt")>
        <TestCase("c:\hoge\fuga\piyo.txt", 1, "...\piyo.txt")>
        <TestCase("\\foo\bar\hoge\fuga\piyo.txt", 100, "\\foo\bar\hoge\fuga\piyo.txt")>
        <TestCase("\\foo\bar\hoge\fuga\piyo.txt", 28, "\\foo\bar\hoge\fuga\piyo.txt")>
        <TestCase("\\foo\bar\hoge\fuga\piyo.txt", 27, "\\foo\bar\...\fuga\piyo.txt")>
        <TestCase("\\foo\bar\hoge\fuga\piyo.txt", 26, "\\foo\bar\...\piyo.txt")>
        <TestCase("\\foo\bar\hoge\fuga\piyo.txt", 22, "\\foo\bar\...\piyo.txt")>
        <TestCase("\\foo\bar\hoge\fuga\piyo.txt", 21, "...\piyo.txt")>
        <TestCase("\\foo\bar\hoge\fuga\piyo.txt", 1, "...\piyo.txt")>
        Public Sub ファイルパスを中略できる_ファイル名は必ず返す_パスルートは優先度高だけど絶対じゃない(filePath As String, maxLength As Integer, expected As String)
            Assert.That(PathUtil.OmitPath(filePath, maxLength), [Is].EqualTo(expected))
        End Sub

    End Class

End Class
