Imports NUnit.Framework
Imports Fhi.Fw.Db.Sql
Imports System.Xml

Namespace Db.Sql
    Public Class XmlInequalityTest

        <Test()> Public Sub Learning_XmlDocument_タグに囲まれた内容の引用符に囲まれた不等号はエラーになる()
            Dim doc As New XmlDocument
            Try
                doc.LoadXml("<aa>a < 'cc<>dd'</aa>")
                Assert.Fail()
            Catch expected As XmlException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub Learning_XmlDocument_属性の引用符に囲まれた不等号はエラーになる()
            Dim doc As New XmlDocument
            Try
                doc.LoadXml("<aa hoge='0 < a.length'>bbb</aa>")
                Assert.Fail()
            Catch expected As XmlException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub ConvInequality_タグに囲まれた内容_不等号はそのまま_CDATAの対象だから()
            Assert.AreEqual("<aa>a < b</aa>", XmlInequality.ConvInequality("<aa>a < b</aa>"))
        End Sub

        <Test()> Public Sub ConvInequality_タグに囲まれた内容_引用符に囲まれた不等号は変換対象()
            Assert.AreEqual("<aa>a < 'cc&lt;&gt;dd'</aa>", XmlInequality.ConvInequality("<aa>a < 'cc<>dd'</aa>"))
        End Sub

        <Test()> Public Sub ConvInequality_タグに囲まれた内容_引用符に囲まれた不等号は変換対象_ダブルクォート版()
            Assert.AreEqual("<aa>a < ""cc&lt;&gt;dd""</aa>", XmlInequality.ConvInequality("<aa>a < ""cc<>dd""</aa>"))
        End Sub

        <Test()> Public Sub ConvInequality_属性_引用符に囲まれた不等号は変換対象()
            Assert.AreEqual("<aa hoge='0 &gt; a.length'>bbb</aa>", XmlInequality.ConvInequality("<aa hoge='0 > a.length'>bbb</aa>"))
            Assert.AreEqual("<aa hoge='0 &lt; a.length'>bbb</aa>", XmlInequality.ConvInequality("<aa hoge='0 < a.length'>bbb</aa>"))
        End Sub

    End Class
End Namespace