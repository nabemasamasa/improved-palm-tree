Imports NUnit.Framework
Imports System.Xml

Namespace Db.Sql

    <TestFixture()> Public MustInherit Class SqlAnalyzerTest

        Private Const POPULAR_SELECT_SQL As String = _
            "select * from hoge " _
            & "<where>" _
            & "<if test='@Id != null'>id = @Id</if>" _
            & "<if test='@Name != null'>and name = @Name</if>" _
            & "<if test='@Last != null'>and last = @Last</if>" _
            & "</where>"
        Private Const POPULAR_UPDATE_SQL As String = _
            "update hoge" _
            & "<set>" _
            & "<if test='@Id != null'>id = @Id,</if>" _
            & "<if test='@Name != null'>name = @Name,</if>" _
            & "<if test='@Last != null'>last = @Last</if>" _
            & "</set>"
        Private Const POPULAR_UPDATE_SQL_WHERE As String = _
            "update hoge" _
            & "<set>" _
            & "<if test='@Name != null'>name = @Name,</if>" _
            & "<if test='@Last != null'>last = @Last</if>" _
            & "</set>" _
            & "<where>" _
            & "<if test='@Id != null'>id = @Id</if>" _
            & "</where>"

        Public Class DefaultTest : Inherits SqlAnalyzerTest

            <Test()> Public Sub LearningXML()
                Dim doc As New XmlDocument
                doc.LoadXml("<sql>" & POPULAR_SELECT_SQL & "</sql>")
                Assert.AreEqual("sql", doc.ChildNodes(0).Name)
                Assert.AreEqual("#text", doc.ChildNodes(0).ChildNodes(0).Name)
                Dim whereNode As XmlNode = doc.ChildNodes(0).ChildNodes(1)
                Assert.AreEqual("where", whereNode.Name)
                Dim ifNode As XmlNode = whereNode.ChildNodes(0)
                Assert.AreEqual("if", ifNode.Name)
                Assert.AreEqual("id = @Id", ifNode.ChildNodes(0).Value)
                Dim ifTestNode As XmlNode = ifNode.Attributes(0)
                Assert.AreEqual("test", ifTestNode.Name)
                Assert.AreEqual("@Id != null", ifTestNode.Value)
            End Sub
            <Test()> Public Sub Analyze_Where_Nothingの場合はWhere句なし()
                Dim analyzer As New SqlAnalyzer(POPULAR_SELECT_SQL, Nothing)
                analyzer.Analyze()
                Assert.AreEqual("select * from hoge", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub Analyze_Where_EmptyParamの場合もWhere句なし()
                Dim param As New HogeVo
                Dim analyzer As New SqlAnalyzer(POPULAR_SELECT_SQL, param)
                analyzer.Analyze()
                Assert.AreEqual("select * from hoge", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub Analyze_SetWhere_ParamがidとnameでSetとWhere()
                Dim param As New HogeVo
                param.Id = 123
                param.Name = "asd"
                Dim analyzer As New SqlAnalyzer(POPULAR_UPDATE_SQL_WHERE, param)
                analyzer.Analyze()
                Assert.AreEqual("update hoge SET name = @Name WHERE id = @Id", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub Analyze_SQLにLTがあってもOK()
                Dim sql = "id < 10"
                Dim analyzer As New SqlAnalyzer(sql, Nothing)
                analyzer.Analyze()
                Assert.AreEqual("id < 10", analyzer.AnalyzedSql)
            End Sub
            <Test()> Public Sub Analyze_SQLにLEがあってもOK()
                Dim sql = "id <= 10"
                Dim analyzer As New SqlAnalyzer(sql, Nothing)
                analyzer.Analyze()
                Assert.AreEqual("id <= 10", analyzer.AnalyzedSql)
            End Sub
            <Test()> Public Sub Analyze_SQLにGTがあってもOK()
                Dim sql = "id > 10"
                Dim analyzer As New SqlAnalyzer(sql, Nothing)
                analyzer.Analyze()
                Assert.AreEqual("id > 10", analyzer.AnalyzedSql)
            End Sub
            <Test()> Public Sub Analyze_SQLにGEがあってもOK()
                Dim sql = "id >= 10"
                Dim analyzer As New SqlAnalyzer(sql, Nothing)
                analyzer.Analyze()
                Assert.AreEqual("id >= 10", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub パラメータ値がnullでもwhereタグに囲まれてたら丸々除去()
                Dim sql As String = "aa <where>" _
                & "<if test='@Id != null'>id = @Id</if>" _
                & "<if test='@Name != null'>and name = @Name</if>" _
                & "<if test='@Last != null'>and last = @Last</if>" _
                & "</where>"
                Dim analyzer As New SqlAnalyzer(sql, Nothing)
                analyzer.Analyze()
                Assert.AreEqual("aa", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub パラメータ値がnullでもwhereタグに囲まれてたらifタグ条件trueだけ()
                Dim sql As String = "aa <where>" _
                & "<if test='false'>id = 'Id'</if>" _
                & "<if test='true'>and name = 'Name'</if>" _
                & "<if test='@Last != null'>and last = @Last</if>" _
                & "</where>"
                Dim analyzer As New SqlAnalyzer(sql, Nothing)
                analyzer.Analyze()
                Assert.AreEqual("aa WHERE name = 'Name'", analyzer.AnalyzedSql)
            End Sub

            Private Class ArrayVo
                Private _strArray As String()

                Public Property StrArray() As String()
                    Get
                        Return _strArray
                    End Get
                    Set(ByVal value As String())
                        _strArray = value
                    End Set
                End Property
            End Class

            <Test()> Public Sub join_配列を()
                Dim sql As String = "aa <where>" _
                & "<if test='@StrArray != null and 0 < @StrArray.Length'>id in (<join property='@StrArray' separator=',' />)</if>" _
                & "and a = '><'" _
                & "</where>"
                Dim vo As New ArrayVo
                vo.StrArray = New String() {"a", "b"}
                Dim analyzer As New SqlAnalyzer(sql, vo)
                analyzer.Analyze()
                Assert.AreEqual("aa WHERE id in ( @StrArray#0,@StrArray#1 ) and a = '><'", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub andタグ_解析できる()
                Dim sql As String = "aa <and>" _
                & "and a = '><'" _
                & "</and>"
                Dim analyzer As New SqlAnalyzer(sql, Nothing)
                analyzer.Analyze()
                Assert.AreEqual("aa AND (a = '><')", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub orタグ_解析できる()
                Dim sql As String = "aa <or>" _
                & "and a = '><'" _
                & "</or>"
                Dim analyzer As New SqlAnalyzer(sql, Nothing)
                analyzer.Analyze()
                Assert.AreEqual("aa OR (a = '><')", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub andタグ_解析できる_where句の先頭ならANDは除く()
                Dim sql As String = "aa <where> " _
                & "<if test='false'>false</if> " _
                & "<and>" _
                & "and a = '><'" _
                & "</and>" _
                & "</where>"
                Dim analyzer As New SqlAnalyzer(sql, Nothing)
                analyzer.Analyze()
                Assert.AreEqual("aa WHERE (a = '><')", analyzer.AnalyzedSql)
            End Sub

        End Class

        Public Class アンド文字が含まれても動くよTest : Inherits SqlAnalyzerTest

            <Test()> Public Sub シングルクォート一つで囲んでても動く()
                Dim analyzer As New SqlAnalyzer("set hoge='A&B'", Nothing)
                analyzer.Analyze()
                Assert.AreEqual("set hoge='A&B'", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub シングルクォート二つで囲んでても動く()
                Dim analyzer As New SqlAnalyzer("set hoge=''A&B''", Nothing)
                analyzer.Analyze()
                Assert.AreEqual("set hoge=''A&B''", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub 不等号とセットでも動く_外側()
                Dim analyzer As New SqlAnalyzer("set hoge<'A&B'", Nothing)
                analyzer.Analyze()
                Assert.AreEqual("set hoge < 'A&B'", analyzer.AnalyzedSql, "不等号は内部でCDATAに囲まれるから、結果左右にスペースができるけど、実害ないのでこれで良い")
            End Sub

            <Test()> Public Sub 不等号とセットでも動く_内側()
                Dim analyzer As New SqlAnalyzer("set hoge='A&B<C'", Nothing)
                analyzer.Analyze()
                Assert.AreEqual("set hoge='A&B<C'", analyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub クォートの外側と内側とにアンド文字があっても動く()
                Dim analyzer As New SqlAnalyzer("set hoge&'A&B'", Nothing)
                analyzer.Analyze()
                Assert.AreEqual("set hoge&'A&B'", analyzer.AnalyzedSql)
            End Sub

        End Class

    End Class
End Namespace
