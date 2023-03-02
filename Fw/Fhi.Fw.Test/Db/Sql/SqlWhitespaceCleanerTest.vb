Imports NUnit.Framework
Imports Fhi.Fw.Db.Sql

Namespace Db.Sql
    Public Class SqlWhitespaceCleanerTest
        <Test()> Public Sub n個の連続スペースを一文字に()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b", sql.Clean("a   b"))
        End Sub

        <Test()> Public Sub 先頭スペースを除去()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("ab", sql.Clean(" ab"))
        End Sub

        <Test()> Public Sub 末尾スペースを除去()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("ab", sql.Clean("ab "))
        End Sub

        <Test()> Public Sub n個のタブをスペース一文字に()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b", sql.Clean("a" & vbTab & vbTab & vbTab & "b"))
        End Sub

        <Test()> Public Sub 一個のタブもスペース一文字に()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b", sql.Clean("a" & vbTab & "b"))
        End Sub

        <Test()> Public Sub 改行をスペース一文字に()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b", sql.Clean("a" & vbCrLf & vbCrLf & vbCrLf & "b"))
        End Sub

        <Test()> Public Sub スペース_タブ_改行をスペース一文字に()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b", sql.Clean("a " & vbTab & " " & vbCrLf & " " & vbTab & " " & vbCrLf & "b"))
        End Sub

        <Test()> Public Sub シングルクォート中のWhitespaceはそのまま()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b=' " & vbTab & vbCrLf & "'", sql.Clean("a   b=' " & vbTab & vbCrLf & "'"))
        End Sub

        <Test()> Public Sub ダブルクォート中のWhitespaceもそのまま()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b="" " & vbTab & vbCrLf & """", sql.Clean("a   b="" " & vbTab & vbCrLf & """"))
        End Sub

        <Test()> Public Sub シングルクォート中のダブルクォートで誤動作しないこと()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b=' " & vbTab & """" & vbCrLf & "'", sql.Clean("a   b=' " & vbTab & """" & vbCrLf & "'"))
        End Sub

        <Test()> Public Sub シングルクォート中のダブルクォートで誤動作しないこと2()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b=' """ & vbTab & """" & vbCrLf & "'", sql.Clean("a   b=' """ & vbTab & """" & vbCrLf & "'"))
        End Sub

        <Test()> Public Sub ダブルクォート中のシングルクォートで誤動作しないこと()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b="" '" & vbTab & vbCrLf & """", sql.Clean("a   b="" '" & vbTab & vbCrLf & """"))
        End Sub

        <Test()> Public Sub ダブルクォート中のシングルクォートで誤動作しないこと2()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("a b="" '" & vbTab & "'" & vbCrLf & """", sql.Clean("a   b="" '" & vbTab & "'" & vbCrLf & """"))
        End Sub

        <Test()> Public Sub hogeダブルクォート中のシングルクォートで誤動作しないこと2()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("@SceneNum )", sql.Clean("@SceneNum  )"))
        End Sub

        <Test()> Public Sub シングルクォート前後のスペースで誤動作しないこと()
            Dim sql As New SqlWhitespaceCleaner
            Assert.AreEqual("A = '#CMN' AND B = 3", sql.Clean("A = '#CMN'  AND B = 3"))
        End Sub
        '
    End Class
End Namespace