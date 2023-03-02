Imports System.IO
Imports NUnit.Framework

Namespace Util.Csv
    Public Class CsvReaderTest

        Private Const FILE_NAME As String = "hogehoge.csv"

        <SetUp()> Public Sub Setup()
            File.Delete(FILE_NAME)
        End Sub

        <TearDown()> Public Sub TearDown()
            File.Delete(FILE_NAME)
        End Sub

        <Test()> Public Sub タイトル行無し_の指定なら列名はFx()
            File.WriteAllLines(FILE_NAME, New String() {"a,b,10"})

            Dim testee As New CsvReader
            testee.Read(FILE_NAME, False)   'タイトル行無しの指定
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual(3, actual.Columns.Count)

            Assert.AreEqual("F1", actual.Columns(0).ColumnName)
            Assert.AreEqual("F2", actual.Columns(1).ColumnName)
            Assert.AreEqual("F3", actual.Columns(2).ColumnName)
        End Sub

        <Test()> Public Sub タイトル行有り_の指定なら一行目が列名()
            File.WriteAllLines(FILE_NAME, New String() {"a,b,10", "x,y,99"})

            Dim testee As New CsvReader
            testee.Read(FILE_NAME, True)    'タイトル行有りの指定
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual(3, actual.Columns.Count)

            Assert.AreEqual("a", actual.Columns(0).ColumnName)
            Assert.AreEqual("b", actual.Columns(1).ColumnName)
            Assert.AreEqual("10", actual.Columns(2).ColumnName)
        End Sub

        <Test()> Public Sub タイトル行無し_の指定なら全行データ()
            File.WriteAllLines(FILE_NAME, New String() {"a,b,10", "x,y,99"})

            Dim testee As New CsvReader
            testee.Read(FILE_NAME, False)  'タイトル行無しの指定
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual(2, actual.Rows.Count)

            Assert.AreEqual("b", actual.Rows(0).Item(1))
            Assert.AreEqual("10", actual.Rows(0).Item("F3"))
            Assert.AreEqual("99", actual.Rows(1).Item(2))
        End Sub

        <Test()> Public Sub タイトル行有り_の指定なら2行目以降がデータ()
            File.WriteAllLines(FILE_NAME, New String() {"a,b,10", "x,y,99"})

            Dim testee As New CsvReader
            testee.Read(FILE_NAME, True)  'タイトル行有りの指定
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual(1, actual.Rows.Count)

            Assert.AreEqual("x", actual.Rows(0).Item(0))
            Assert.AreEqual("y", actual.Rows(0).Item("b"))
            Assert.AreEqual("99", actual.Rows(0).Item(2))
        End Sub

        <Test()> Public Sub ダブルクォートで囲まれた文字列はそのまま取得される_カンマ等の記号は無視される()
            File.WriteAllLines(FILE_NAME, New String() {"1,""hoge,TOP"",b"})

            Dim testee As New CsvReader
            testee.Read(FILE_NAME, False)
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual(1, actual.Rows.Count)

            Assert.AreEqual("1", actual.Rows(0).Item(0))
            Assert.AreEqual("hoge,TOP", actual.Rows(0).Item(1))
            Assert.AreEqual("b", actual.Rows(0).Item(2))
        End Sub

        <Test()> Public Sub ダブルクォートで囲まれた文字列はそのまま取得される_改行は行区切りとして扱わない()
            File.WriteAllLines(FILE_NAME, New String() {"1,""hoge" & vbCrLf & "TOP"",b"})

            Dim testee As New CsvReader
            testee.Read(FILE_NAME, False)
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual(1, actual.Rows.Count)

            Assert.AreEqual("1", actual.Rows(0).Item(0))
            Assert.AreEqual("hoge" & vbCrLf & "TOP", actual.Rows(0).Item(1))
            Assert.AreEqual("b", actual.Rows(0).Item(2))
        End Sub

        <Test()> Public Sub ダブルクォートで囲まれた文字列はそのまま取得される_ダブルクォートは二重にしてあれば読み込まれる()
            File.WriteAllLines(FILE_NAME, New String() {"1,""hoge""""TOP"",b"})

            Dim testee As New CsvReader
            testee.Read(FILE_NAME, False)
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual(1, actual.Rows.Count)

            Assert.AreEqual("1", actual.Rows(0).Item(0))
            Assert.AreEqual("hoge""TOP", actual.Rows(0).Item(1))
            Assert.AreEqual("b", actual.Rows(0).Item(2))
        End Sub

        <Test()> Public Sub ダブルクォートで囲まれた文字列はそのまま取得される_ダブルクォートが一重でも読めるように_FHIのFTPインポートが一重で返すから()
            File.WriteAllLines(FILE_NAME, New String() {"1,""hoge""TOP"",b"})

            Dim testee As New CsvReader
            testee.Read(FILE_NAME, False)
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual(1, actual.Rows.Count)

            Assert.AreEqual("1", actual.Rows(0).Item(0))
            Assert.AreEqual("hoge""TOP", actual.Rows(0).Item(1))
            Assert.AreEqual("b", actual.Rows(0).Item(2))
        End Sub

        <Test()> Public Sub TrimWhitespace_初期値は左右のtrim()
            File.WriteAllLines(FILE_NAME, New String() {"1,""  hoge  """})

            Dim testee As New CsvReader()
            testee.Read(FILE_NAME, False)
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual(1, actual.Rows.Count)
            Assert.AreEqual("hoge", actual.Rows(0).Item(1))
        End Sub

        <Test()> Public Sub TrimWhitespace_false設定でtrimしない()
            File.WriteAllLines(FILE_NAME, New String() {"1,""  hoge  """})

            Dim testee As New CsvReader()
            testee.TrimWhitespace = False
            testee.Read(FILE_NAME, False)
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual(1, actual.Rows.Count)
            Assert.AreEqual("  hoge  ", actual.Rows(0).Item(1))
        End Sub

        <Test()> Public Sub コメント行は読まれない()
            File.WriteAllLines(FILE_NAME, New String() {"#a,b,10", "x,y,99"})

            Dim testee As New CsvReader()
            testee.Read(FILE_NAME, False)
            Dim actual As DataTable = testee.GetTable

            Assert.AreEqual("x", actual.Rows(0).Item(0))
            Assert.AreEqual("y", actual.Rows(0).Item(1))
            Assert.AreEqual("99", actual.Rows(0).Item(2))
        End Sub

        <Test()> Public Sub CorrectIllegalDblQuote_補正の必要なし()
            Assert.AreEqual(",,""a,b"",", CsvReader.CorrectIllegalDblQuote(",,""a,b"","))
        End Sub

        <Test()> Public Sub CorrectIllegalDblQuote_補正する()
            Assert.AreEqual(",,""a,""""b"",", CsvReader.CorrectIllegalDblQuote(",,""a,""b"","))
        End Sub

        <Test()> Public Sub CorrectIllegalDblQuote_補正する2()
            Assert.AreEqual(",,""a,b"""""",", CsvReader.CorrectIllegalDblQuote(",,""a,b"""","))
        End Sub

        <Test()> Public Sub CorrectIllegalDblQuote_補正の必要なし2()
            Assert.AreEqual(",,""a,b"""""",", CsvReader.CorrectIllegalDblQuote(",,""a,b"""""","))
        End Sub

        <Test()> Public Sub CorrectIllegalDblQuote_FHIの実際のデータ_補正する()
            Assert.AreEqual("aa,""clip """"P"""" "",cc", CsvReader.CorrectIllegalDblQuote("aa,""clip ""P"" "",cc"))
        End Sub

        <Test()> Public Sub 別のプロセスに使用されていても_ファイル読み込みができる()
            File.WriteAllLines(FILE_NAME, New String() {"a,b,10"})

            Using New FileStream(FILE_NAME, FileMode.Open)
                Dim testee As New CsvReader
                testee.Read(FILE_NAME, False)
                Dim actual As DataTable = testee.GetTable()

                Assert.AreEqual(3, actual.Columns.Count)

                Assert.AreEqual("F1", actual.Columns(0).ColumnName)
                Assert.AreEqual("F2", actual.Columns(1).ColumnName)
                Assert.AreEqual("F3", actual.Columns(2).ColumnName)
            End Using
        End Sub

    End Class
End Namespace