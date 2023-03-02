Imports System.IO
Imports NUnit.Framework

Public MustInherit Class FileUtilTest
    Public Class DefaultTest : Inherits FileUtilTest
        Private Const PATH_FILE_NAME As String = "hoge.txt"

        <Test()> Public Sub CreateHardLink_リンク元ファイルと中身が同じである_相対パス()
            Dim destPathFileName As String = Path.Combine(AssemblyUtil.GetPath, "fuga.txt")
            Dim srcPathFileName As String = Path.Combine(AssemblyUtil.GetPath, "hoge.txt")

            FileUtil.DeleteFileIfExist(destPathFileName)

            FileUtil.WriteFile(srcPathFileName, "aiueo" & vbCrLf & "1234", True)
            Try
                FileUtil.CreateHardLink(srcPathFileName, destPathFileName)
                Try
                    Assert.IsTrue(File.Exists(destPathFileName))

                    Dim rows As String() = FileUtil.ReadFileAsArray(destPathFileName)
                    Assert.AreEqual(2, rows.Length)
                    Assert.AreEqual("aiueo", rows(0))
                    Assert.AreEqual("1234", rows(1))

                Finally
                    FileUtil.DeleteFileIfExist(destPathFileName)
                End Try
            Finally
                FileUtil.DeleteFileIfExist(srcPathFileName)
            End Try
        End Sub

        <Test()> Public Sub CreateHardLink_リンク元ファイルと中身が同じである_絶対パス()

            Dim destPathFileName As String = Path.Combine(AssemblyUtil.GetPath, "piyo.txt")
            Dim srcPathFileName As String = Path.Combine(AssemblyUtil.GetPath, "hoge.txt")

            FileUtil.DeleteFileIfExist(destPathFileName)

            FileUtil.WriteFile(srcPathFileName, "123" & vbCrLf & "qwerty", True)
            Try
                FileUtil.CreateHardLink(srcPathFileName, destPathFileName)
                Try
                    Assert.IsTrue(File.Exists(destPathFileName))

                    Dim rows As String() = FileUtil.ReadFileAsArray(destPathFileName)
                    Assert.AreEqual(2, rows.Length)
                    Assert.AreEqual("123", rows(0))
                    Assert.AreEqual("qwerty", rows(1))

                Finally
                    FileUtil.DeleteFileIfExist(destPathFileName)
                End Try
            Finally
                FileUtil.DeleteFileIfExist(srcPathFileName)
            End Try
        End Sub

        <Test()> Public Sub MakeFolderIfNotExists_通常()
            Dim entryFolderPath As String = Path.Combine(AssemblyUtil.GetPath, "a")
            Dim testingFolderPath As String = Path.Combine(AssemblyUtil.GetPath, "a\b\c")
            FileUtil.DeleteFolderIfExist(entryFolderPath, True)
            FileUtil.MakeFolderIfNotExists(testingFolderPath)

            Try
                Assert.IsTrue(Directory.Exists(testingFolderPath))
            Finally
                FileUtil.DeleteFolderIfExist(entryFolderPath, True)
            End Try

        End Sub

        <Test(), Sequential()> Public Sub EscapeCsvValue_CSV値は必要なければ_エスケープしない( _
            <Values("x", "あ", vbTab, ".")> ByVal value As String)
            Assert.That(FileUtil.EscapeCsvValue("a" & value & "b"), [Is].EqualTo("a" & value & "b"))
            Assert.That(FileUtil.EscapeCsvValue(value), [Is].EqualTo(value))
        End Sub

        <Test(), Sequential()> Public Sub EscapeCsvValue_CSV値は必要ならエスケープされる( _
            <Values(",", vbCr, vbLf, vbCrLf, """")> ByVal value As String, _
            <Values(",", vbCr, vbLf, vbCrLf, """""")> ByVal actual As String)
            Const DQ As String = """"
            Assert.That(FileUtil.EscapeCsvValue("a" & value & "b"), [Is].EqualTo(DQ & "a" & actual & "b" & DQ))
            Assert.That(FileUtil.EscapeCsvValue(value), [Is].EqualTo(DQ & actual & DQ))
        End Sub

        <Test(), Sequential()> Public Sub EscapeTsvValue_CSV値は必要なければ_エスケープしない( _
            <Values("x", "あ", ",")> ByVal value As String)
            Assert.That(FileUtil.EscapeTsvValue("a" & value & "b"), [Is].EqualTo("a" & value & "b"))
            Assert.That(FileUtil.EscapeTsvValue(value), [Is].EqualTo(value))
        End Sub

        <Test(), Sequential()> Public Sub EscapeTsvValue_CSV値は必要ならエスケープされる( _
            <Values(vbTab, vbCr, vbLf, vbCrLf, """")> ByVal value As String, _
            <Values(vbTab, vbCr, vbLf, vbCrLf, """""")> ByVal actual As String)
            Const DQ As String = """"
            Assert.That(FileUtil.EscapeTsvValue("a" & value & "b"), [Is].EqualTo(DQ & "a" & actual & "b" & DQ))
            Assert.That(FileUtil.EscapeTsvValue(value), [Is].EqualTo(DQ & actual & DQ))
        End Sub

        <Test()> Public Sub EscapeTsvValue_TSV値は必要ならエスケープされる()
            Const DQ As String = """"
            Assert.AreEqual("a", FileUtil.EscapeTsvValue("a"))
            Assert.AreEqual(DQ & "a" & DQ & DQ & "b" & DQ, FileUtil.EscapeTsvValue("a""b"))
            Assert.AreEqual(DQ & "a" & vbCr & "b" & DQ, FileUtil.EscapeTsvValue("a" & vbCr & "b"))
            Assert.AreEqual(DQ & "a" & vbLf & "b" & DQ, FileUtil.EscapeTsvValue("a" & vbLf & "b"))
            Assert.AreEqual(DQ & "a" & vbCrLf & "b" & DQ, FileUtil.EscapeTsvValue("a" & vbCrLf & "b"))
            Assert.AreEqual(DQ & "a" & vbTab & "b" & DQ, FileUtil.EscapeTsvValue("a" & vbTab & "b"))
        End Sub

        <Test()> Public Sub EscapeCsvValue_CSV値でタブはエスケープされない()
            Assert.AreEqual("a" & vbTab & "b", FileUtil.EscapeCsvValue("a" & vbTab & "b"))
        End Sub

        <Test()> Public Sub EscapeTsvValue_TSV値でカンマはエスケープされない()
            Assert.AreEqual("a,b", FileUtil.EscapeTsvValue("a,b"))
        End Sub

        <Test()> Public Sub RenameFileIfExist_ファイル名が変更される()
            Dim destPathFileName As String = Path.Combine(AssemblyUtil.GetPath, "fuga.txt")
            Dim srcPathFileName As String = Path.Combine(AssemblyUtil.GetPath, "hoge.txt")

            FileUtil.DeleteFileIfExist(destPathFileName)
            FileUtil.WriteFile(srcPathFileName, "aiueo" & vbCrLf & "1234", True)

            Try
                FileUtil.RenameFileIfExist(srcPathFileName, destPathFileName)
                Try
                    Assert.IsFalse(File.Exists(srcPathFileName))
                    Assert.IsTrue(File.Exists(destPathFileName))
                Finally
                    FileUtil.DeleteFileIfExist(destPathFileName)
                End Try
            Finally
                FileUtil.DeleteFileIfExist(srcPathFileName)
            End Try
        End Sub
    End Class

    Public Class GetFilePathsUnderDirectoryTest : Inherits FileUtilTest
        Private Shared ReadOnly parent As String = PathUtil.Combine(AssemblyUtil.GetPath, "Test")

        <Test()>
        Public Sub フォルダ配下のファイルパスを取得できる()
            Dim targets As New List(Of String) From {Path.Combine(parent, "piyo.txt"),
                                                       Path.Combine(parent, "hoge.txt")}
            FileUtil.MakeFolderIfNotExists(parent)
            targets.ForEach(Sub(target) FileUtil.WriteFile(target, "hoge"))
            Try
                Assert.That(FileUtil.GetFilePathsUnderDirectory(parent), [Is].EquivalentTo(targets.ToArray))
            Finally
                targets.ForEach(Sub(target) FileUtil.DeleteFileIfExist(target))
                FileUtil.DeleteFolderIfExist(parent, True)
            End Try
        End Sub

        <Test()>
        Public Sub 第二引数指定がなければ_第一引数のフォルダ直下のファイルパスだけ取得する()
            Dim targets As New List(Of String) From {
                    Path.Combine(parent, "piyo.txt"),
                    Path.Combine(parent, "hoge.txt")}
            FileUtil.MakeFolderIfNotExists(Path.Combine(parent, "Fuga"))
            targets.Concat({PathUtil.Combine(parent, "Fuga", "musa.txt")}).ToList.ForEach(Sub(target) FileUtil.WriteFile(target, "hoge"))
            Try
                Assert.That(FileUtil.GetFilePathsUnderDirectory(parent), [Is].EquivalentTo(targets.ToArray))
            Finally
                targets.Concat({PathUtil.Combine(parent, "Fuga", "musa.txt")}).ToList.ForEach(Sub(target) FileUtil.DeleteFileIfExist(target))
                FileUtil.DeleteFolderIfExist(parent, True)
            End Try
        End Sub

        <Test()>
        Public Sub 第二引数指定がTrueなら_再帰的にファイルパスを取得できる()
            Dim targets As New List(Of String) From {
                    PathUtil.Combine(parent, "Fuga", "musa.txt"),
                    Path.Combine(parent, "piyo.txt"),
                    Path.Combine(parent, "hoge.txt")}

            FileUtil.MakeFolderIfNotExists(Path.Combine(parent, "Fuga"))
            targets.ForEach(Sub(target) FileUtil.WriteFile(target, "hoge"))

            Try
                Assert.That(FileUtil.GetFilePathsUnderDirectory(parent, True), [Is].EquivalentTo(targets.ToArray))
            Finally
                targets.Concat({PathUtil.Combine(parent, "Fuga", "musa.txt")}).ToList.ForEach(Sub(target) FileUtil.DeleteFileIfExist(target))
                FileUtil.DeleteFolderIfExist(parent, True)
            End Try
        End Sub

        <Test()>
        Public Sub 再帰的に取得する場合_深い階層が上にくるように並び変えられる()
            Dim targets As New List(Of String) From {
                    PathUtil.Combine(parent, "A", "C", "D", "F", "f.txt"),
                    PathUtil.Combine(parent, "A", "C", "D", "d.txt"),
                    PathUtil.Combine(parent, "A", "C", "E", "e.txt"),
                    PathUtil.Combine(parent, "A", "B", "b.txt"),
                    PathUtil.Combine(parent, "A", "C", "c.txt"),
                    PathUtil.Combine(parent, "A", "a.txt")}
            FileUtil.MakeFolderIfNotExists(PathUtil.Combine(parent, "A", "C", "D", "F"))
            FileUtil.MakeFolderIfNotExists(PathUtil.Combine(parent, "A", "C", "E"))
            FileUtil.MakeFolderIfNotExists(PathUtil.Combine(parent, "A", "B"))
            targets.ForEach(Sub(target) FileUtil.WriteFile(target, "hoge"))
            Try
                Assert.That(FileUtil.GetFilePathsUnderDirectory(parent, True), [Is].EqualTo(targets.ToArray))
            Finally
                targets.ForEach(Sub(target) FileUtil.DeleteFileIfExist(target))
                FileUtil.DeleteFolderIfExist(parent, True)
            End Try
        End Sub

    End Class

End Class
