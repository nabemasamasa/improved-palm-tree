Imports System.IO
Imports NUnit.Framework
Imports System.Text

Namespace Sys

    Public Class PrivateProfileTest

        Private Shared ReadOnly FILE_NAME As String = Path.Combine(AssemblyUtil.GetPath, "hoge.ini")

        Private Sub DeleteIniFileIfExists()

            If File.Exists(FILE_NAME) Then
                File.Delete(FILE_NAME)
            End If
        End Sub

        <SetUp()> Public Sub Setup()
            DeleteIniFileIfExists()
        End Sub

        <TearDown()> Public Sub TearDown()
            DeleteIniFileIfExists()
        End Sub

        <Test()> Public Sub ファイルが無ければファイルが作成される()

            Assert.IsFalse(File.Exists(FILE_NAME), "ファイルが存在しないときに")

            Dim a As New PrivateProfile(FILE_NAME)
            a.Write("DIV", "KEY", "hogehoge")

            Assert.IsTrue(File.Exists(FILE_NAME), "Write()したからファイルが出来る")
        End Sub

        <Test()> Public Sub Writeしたファイル内容()
            Dim a As New PrivateProfile(FILE_NAME)

            a.Write("DIV", "KEY", "hogehoge")

            Dim lines As String() = File.ReadAllLines(FILE_NAME)
            Assert.AreEqual("[DIV]" & vbCrLf _
                            & "KEY=hogehoge", Join(lines, vbCrLf))
        End Sub

        <Category("IgnoreGHA")>
        <Test()> Public Sub Writeしたファイル内容_内容はShiftJisで出力される()
            Dim a As New PrivateProfile(FILE_NAME)

            a.Write("DIV", "あああ", "いいいい")

            Dim lines As String() = File.ReadAllLines(FILE_NAME, Encoding.GetEncoding("SHIFT_JIS"))
            Assert.AreEqual("[DIV]" & vbCrLf _
                            & "あああ=いいいい", Join(lines, vbCrLf))
        End Sub

        <Category("IgnoreGHA")>
        <Test()> Public Sub Writeしたファイル内容_キーに改行が含まれてもおかまいなしに出力される()
            Dim a As New PrivateProfile(FILE_NAME)

            a.Write("DIV", "国内" & vbCrLf _
                        & "SDN" & vbCrLf _
                        & "925" & vbCrLf _
                        & "DOHC" & vbCrLf _
                        & "AWD" & vbCrLf _
                        & "CVT" & vbCrLf _
                        & "Φ152" & vbCrLf _
                        & "2.5i-L I" & vbCrLf _
                        & "RHD", "hogehoge")

            Dim lines As String() = File.ReadAllLines(FILE_NAME, Encoding.GetEncoding(932))
            Assert.AreEqual("[DIV]" & vbCrLf _
                            & "国内" & vbCrLf _
                            & "SDN" & vbCrLf _
                            & "925" & vbCrLf _
                            & "DOHC" & vbCrLf _
                            & "AWD" & vbCrLf _
                            & "CVT" & vbCrLf _
                            & "Φ152" & vbCrLf _
                            & "2.5i-L I" & vbCrLf _
                            & "RHD=hogehoge", Join(lines, vbCrLf))
        End Sub

        <Test()> Public Sub Write直後にReadも出来る()
            Dim a As New PrivateProfile(FILE_NAME)

            a.Write("DIV", "KEY", "hogehoge")

            Assert.AreEqual("hogehoge", a.Read("DIV", "KEY"))
        End Sub

        <Test()> Public Sub 書式に沿ったiniファイルならReadできる()
            File.WriteAllLines(FILE_NAME, New String() {"[SEC]", "hoge=333"})

            Dim a As New PrivateProfile(FILE_NAME)

            Assert.AreEqual("333", a.Read("SEC", "hoge"))
        End Sub

        <Test()> Public Sub MaxValueLengthで読込サイズを指定できる()
            File.WriteAllLines(FILE_NAME, New String() {"[SEC]", "hoge=1234567890"})

            Dim a As New PrivateProfile(FILE_NAME, 8)

            Assert.That(a.Read("SEC", "hoge"), [Is].EqualTo("1234567"), "読込サイズ8だと7文字読み込める")
        End Sub

        <Test()> Public Sub 書式に沿ったiniファイルならReadできる_初めに読み込んだ値が後から読み込んだ値に上書きされていないことを確認()
            File.WriteAllLines(FILE_NAME, New String() {"[SEC]", "hoge=333", "fuga=abc"})

            Dim a As New PrivateProfile(FILE_NAME)

            Dim actualHoge As String = a.Read("SEC", "hoge")
            Dim actualFuga As String = a.Read("SEC", "fuga")
            Assert.AreEqual("333", actualHoge)
            Assert.AreEqual("abc", actualFuga)
        End Sub

        <Test()> Public Sub Remove_キー毎消える()
            File.WriteAllLines(FILE_NAME, New String() {"[SEC]", "hoge=333", "fuga=abc"})

            Dim a As New PrivateProfile(FILE_NAME)

            a.Remove("SEC", "hoge")

            Dim actuals As String() = File.ReadAllLines(FILE_NAME)
            Assert.AreEqual(2, actuals.Length, "hogeが消えるので行数は2")
            Assert.AreEqual("[SEC]", actuals(0))
            Assert.AreEqual("fuga=abc", actuals(1))
        End Sub

        <Test()> Public Sub GetKeys_キー名の一覧を返す()

            File.WriteAllLines(FILE_NAME, New String() {"[SEC]", "hoge=333", "fuga=abc"})

            Dim a As New PrivateProfile(FILE_NAME)

            Dim actuals As String() = a.GetKeys("SEC")

            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("hoge", actuals(0))
            Assert.AreEqual("fuga", actuals(1))
        End Sub

        <Test()> Public Sub GetKeys_セクションが無ければ配列長さ0()

            File.WriteAllLines(FILE_NAME, New String() {"[SEC]", "hoge=333", "fuga=abc"})

            Dim a As New PrivateProfile(FILE_NAME)

            Dim actuals As String() = a.GetKeys("fuga")

            Assert.AreEqual(0, actuals.Length)
        End Sub

        <Test()> Public Sub GetSections_セクション名の一覧を返す()

            File.WriteAllLines(FILE_NAME, New String() {"[SEC]", "hoge=333", "[TION]", "fuga=abc"})

            Dim a As New PrivateProfile(FILE_NAME)

            Dim actuals As String() = a.GetSections

            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("SEC", actuals(0))
            Assert.AreEqual("TION", actuals(1))
        End Sub

    End Class
End Namespace