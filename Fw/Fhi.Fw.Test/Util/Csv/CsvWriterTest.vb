Imports System.IO
Imports NUnit.Framework

Namespace Util.Csv
    ''' <summary>
    ''' Csvの書き込みを担うクラスのテストクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CsvWriterTest

        Private Const FILE_NAME As String = "fugafuga.csv"
        Private writer As CsvWriter

        <SetUp()> Public Sub Setup()
            writer = New CsvWriter(FILE_NAME)
        End Sub

        <TearDown()> Public Sub TearDown()
            File.Delete(FILE_NAME)
        End Sub

        <Test()> Public Sub Append_1文字ずつ追加できる_カンマ区切りの場合()
            writer.Separator = ","
            writer.Append("a")
            writer.Append("b")
            writer.Append("c")
            writer.Next() ' Appendの場合、Nextで改行が必要
            writer.Write()
            Assert.That(FileUtil.ReadFile(FILE_NAME), [Is].EqualTo("""a"",""b"",""c"""))
        End Sub

        <Test()> Public Sub Append_1文字ずつ追加できる_タブ区切りの場合()
            writer.Append("a")
            writer.Append("b")
            writer.Append("c")
            writer.Next() ' Appendの場合、Nextで改行が必要
            writer.Write()
            Assert.That(FileUtil.ReadFile(FILE_NAME), [Is].EqualTo("""a""" & vbTab & """b""" & vbTab & """c"""))
        End Sub

        <Test()> Public Sub AppendLine_1行ずつ追加できる_カンマ区切りの場合()
            writer.Separator = ","
            writer.AppendLine("a", "b", "c")
            writer.AppendLine("", "d", "e")
            writer.Write()
            Assert.That(FileUtil.ReadFile(FILE_NAME), [Is].EqualTo("""a"",""b"",""c""" & vbCrLf & """"",""d"",""e"""))
        End Sub

        <Test()> Public Sub AppendLine_1行ずつ追加できる_タブ区切りの場合()
            writer.AppendLine("a", "b", "c")
            writer.AppendLine("", "d", "e")
            writer.Write()
            Assert.That(FileUtil.ReadFile(FILE_NAME), [Is].EqualTo("""a""" & vbTab & """b""" & vbTab & """c""" & vbCrLf & """""" & vbTab & """d""" & vbTab & """e"""))
        End Sub
    End Class
End Namespace