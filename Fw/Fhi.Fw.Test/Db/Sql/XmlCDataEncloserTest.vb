Imports NUnit.Framework
Imports Fhi.Fw.Db.Sql

Namespace Db.Sql

    <TestFixture()> Public Class XmlCDataEncloserTest

        <Test()> Public Sub Enclose_単純なLTはCDataで囲む()
            Dim sql = "id < 10"
            Dim encloser As New XmlCDataEncloser("if", "where", "set")
            Assert.AreEqual("id <![CDATA[<]]> 10", encloser.Enclose(sql))
        End Sub

        <Test()> Public Sub Enclose_単純なLEはCDataで囲む()
            Dim sql = "id <= 10"
            Dim encloser As New XmlCDataEncloser("if", "where", "set")
            Assert.AreEqual("id <![CDATA[<=]]> 10", encloser.Enclose(sql))
        End Sub

        <Test()> Public Sub Enclose_単純なGTはCDataで囲む()
            Dim sql = "id > 10"
            Dim encloser As New XmlCDataEncloser("if", "where", "set")
            Assert.AreEqual("id <![CDATA[>]]> 10", encloser.Enclose(sql))
        End Sub

        <Test()> Public Sub Enclose_単純なGEはCDataで囲む()
            Dim sql = "id >= 10"
            Dim encloser As New XmlCDataEncloser("if", "where", "set")
            Assert.AreEqual("id <![CDATA[>=]]> 10", encloser.Enclose(sql))
        End Sub

        <Test()> Public Sub Enclose_単純なNEはCDataで囲む()
            Dim sql = "id <> 10"
            Dim encloser As New XmlCDataEncloser("if", "where", "set")
            Assert.AreEqual("id <![CDATA[<>]]> 10", encloser.Enclose(sql))
        End Sub

        <Test()> Public Sub Enclose_XMLタグと同居したLTはCDataで囲む()
            Dim sql = "<where>id < 10</where>"
            Dim encloser As New XmlCDataEncloser("if", "where", "set")
            Assert.AreEqual("<where>id <![CDATA[<]]> 10</where>", encloser.Enclose(sql))
        End Sub
        <Test()> Public Sub Enclose_XMLタグと同居したLEはCDataで囲む()
            Dim sql = "<where>id <= 10</where>"
            Dim encloser As New XmlCDataEncloser("if", "where", "set")
            Assert.AreEqual("<where>id <![CDATA[<=]]> 10</where>", encloser.Enclose(sql))
        End Sub
        <Test()> Public Sub Enclose_XMLタグと同居したGTはCDataで囲む()
            Dim sql = "<if test='hoge'>id > 10</if>"
            Dim encloser As New XmlCDataEncloser("if", "where", "set")
            Assert.AreEqual("<if test='hoge'>id <![CDATA[>]]> 10</if>", encloser.Enclose(sql))
        End Sub
        <Test()> Public Sub Enclose_XMLタグと同居したGEはCDataで囲む()
            Dim sql = "<if test='hoge'>id >= 10</if>"
            Dim encloser As New XmlCDataEncloser("if", "where", "set")
            Assert.AreEqual("<if test='hoge'>id <![CDATA[>=]]> 10</if>", encloser.Enclose(sql))
        End Sub
        <Test()> Public Sub Enclose_SQL文中のリテラルに含まれる不等号は囲まない()
            Dim sql = "id >= 10 and name like '%><%'"
            Dim encloser As New XmlCDataEncloser("if", "where", "set")
            Assert.AreEqual("id <![CDATA[>=]]> 10 and name like '%><%'", encloser.Enclose(sql))
        End Sub

    End Class
End Namespace
