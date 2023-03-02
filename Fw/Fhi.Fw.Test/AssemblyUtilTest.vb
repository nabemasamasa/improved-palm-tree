Imports NUnit.Framework

''' <summary>
''' アセンブリに関するユーティリティのテストクラス
''' </summary>
''' <remarks></remarks>
Public MustInherit Class AssemblyUtilTest

    Public Class DefaultTest : Inherits AssemblyUtilTest

        <Test()> Public Sub GetAssemblyNameOfProject_クラスが属するプロジェクトのアセンブリ名を取得できる()
            Assert.That(AssemblyUtil.GetAssemblyNameOfProject(GetType(AssemblyUtilTest)), [Is].EqualTo("Fhi.Fw.Test"))
        End Sub

        <Test()> Public Sub GetAssemblyNameOfSolution_スタートアッププロジェクトのアセンブリ名を取得できる()
            Assert.That(AssemblyUtil.GetAssemblyNameOfSolution, [Is].EqualTo("Fhi.Fw"))
        End Sub

        <Test()> Public Sub GetTypesInNamespace_指定した名前空間内の型一覧を取得できる()
            Dim types As Type() = AssemblyUtil.GetTypesInNamespace("Fhi.Fw.TestUtil.TestDir", GetType(AssemblyUtilTest))
            Assert.That(types.Count, [Is].EqualTo(3))
            Assert.That(types(0).FullName, [Is].EqualTo("Fhi.Fw.TestUtil.TestDir.A"))
            Assert.That(types(1).FullName, [Is].EqualTo("Fhi.Fw.TestUtil.TestDir.B"))
            Assert.That(types(2).FullName, [Is].EqualTo("Fhi.Fw.TestUtil.TestDir.B+C"))
            Assert.That(types(0).Name, [Is].EqualTo("A"))
            Assert.That(types(1).Name, [Is].EqualTo("B"))
            Assert.That(types(2).Name, [Is].EqualTo("C"))
        End Sub

    End Class

    Public Class LearningUriBuilderTest : Inherits AssemblyUtilTest

        <Test()> Public Sub HostプロパティはUNCパスのホストになる()
            Dim builder As New UriBuilder("\\hoge\path\to")
            Assert.That(builder.Host, [Is].EqualTo("hoge"))
        End Sub

        <Test()> Public Sub PathプロパティはUNCパスのHOSTが消える()
            Dim builder As New UriBuilder("\\hoge\path\to")
            Assert.That(builder.Path, [Is].EqualTo("/path/to"))
        End Sub

        <Test()> Public Sub LocalPathプロパティはUNCパスのホストは消えない()
            Dim builder As New UriBuilder("\\hoge\path\to")
            Assert.That(builder.Uri.LocalPath, [Is].EqualTo("\\hoge\path\to"))
        End Sub

        <Test()> Public Sub LocalPathプロパティはローカルパスならそのまま()
            Dim builder As New UriBuilder("d:\path\to")
            Assert.That(builder.Uri.LocalPath, [Is].EqualTo("d:\path\to"))
        End Sub

        <Test()> Public Sub AbsolutePathプロパティはUNCパスのホストは消える()
            Dim builder As New UriBuilder("\\hoge\path\to")
            Assert.That(builder.Uri.AbsolutePath, [Is].EqualTo("/path/to"))
        End Sub

        <Test()> Public Sub UriプロパティのToStringだと_fileプロトコルが付く_UNCパス()
            Dim builder As New UriBuilder("\\hoge\path\to")
            Assert.That(builder.Uri.ToString, [Is].EqualTo("file://hoge/path/to"))
        End Sub

        <Test()> Public Sub UriプロパティのToStringだと_fileプロトコルが付く_ローカルパスだとスラッシュ3つ()
            Dim builder As New UriBuilder("d:\path\to")
            Assert.That(builder.Uri.ToString, [Is].EqualTo("file:///d:/path/to"))
        End Sub

    End Class
End Class
