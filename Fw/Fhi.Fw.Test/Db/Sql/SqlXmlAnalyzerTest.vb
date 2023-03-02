Imports NUnit.Framework
Imports Fhi.Fw.Db.Sql

Namespace Db.Sql
    <TestFixture()> Public MustInherit Class SqlXmlAnalyzerTest

        Protected Class FakeEvaluator : Implements ISqlExpressionEvaluator
            Public Function Evaluate(ByVal expression As String, ByVal param As Object) As Boolean Implements ISqlExpressionEvaluator.Evaluate
                Return Convert.ToBoolean(expression)
            End Function
        End Class

        Public Class whereタグ : Inherits SqlXmlAnalyzerTest

            <Test()> Public Sub whereタグを使わないでifタグだけだと先頭のANDは除外出来ない()
                Dim sql As String = "<sql>...where " _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='true'>AND NAME = @Name</if>" _
                & "<if test='true'>AND ADDRESS = @Address</if>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("...where  AND NAME = @Name AND ADDRESS = @Address", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub whereタグとifタグを使うと先頭のANDを除外する()
                Dim sql As String = "<sql>...<where>" _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='true'>AND NAME = @Name</if>" _
                & "<if test='true'>AND ADDRESS = @Address</if>" _
                & "</where>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... WHERE NAME = @Name AND ADDRESS = @Address", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub whereタグとifタグを使うと先頭のORを除外する()
                Dim sql As String = "<sql>...<where>" _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='true'>OR NAME = @Name</if>" _
                & "<if test='true'>OR ADDRESS = @Address</if>" _
                & "</where>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... WHERE NAME = @Name OR ADDRESS = @Address", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub whereタグを使わないでifタグだけで全falseだとWhere句除去できない()
                Dim sql As String = "<sql>...where " _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='false'>AND NAME = @Name</if>" _
                & "<if test='false'>AND ADDRESS = @Address</if>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("...where ", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub whereタグとifタグを使い全falseだとWhere句除去()
                Dim sql As String = "<sql>...<where> " _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='false'>AND NAME = @Name</if>" _
                & "<if test='false'>AND ADDRESS = @Address</if>" _
                & "</where>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("...", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub whereタグ内にifタグとifに囲まれていない条件分があり_ifのtrueの結果で先頭のANDを除外する_ifが先()
                Dim sql As String = "<sql>...<where> " _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='true'>AND NAME = @Name</if>" _
                & "AND ADDRESS = @Address" _
                & "</where>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... WHERE NAME = @Name AND ADDRESS = @Address", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub whereタグ内にifタグとifに囲まれていない条件分があり_ifのtrueの結果で先頭のANDを除外する_囲まれてない条件が先()
                Dim sql As String = "<sql>...<where> " _
                & "<if test='false'>ID = @Id</if>" _
                & "AND NAME = @Name" _
                & "<if test='true'>AND ADDRESS = @Address</if>" _
                & "</where>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... WHERE NAME = @Name AND ADDRESS = @Address", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub whereタグ内にifタグとifに囲まれていない条件分があり_ifは全falseで_ifに囲まれていない条件文のANDを除外する()
                Dim sql As String = "<sql>...<where> " _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='false'>AND NAME = @Name</if>" _
                & "AND ADDRESS = @Address" _
                & "</where>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... WHERE ADDRESS = @Address", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub whereタグ内にifタグとifに囲まれていない条件分があり_ifは全falseで_ifに囲まれていない条件文のANDを除外する_先頭にwhitespace()
                Dim sql As String = "<sql>...<where> " _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='false'>AND NAME = @Name</if>" _
                & "    AND ADDRESS = @Address" _
                & "</where>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... WHERE ADDRESS = @Address", xmlAnalyzer.AnalyzedSql)
            End Sub

        End Class

        Public Class setタグ : Inherits SqlXmlAnalyzerTest

            <Test()> Public Sub setタグを使わないでifタグだけだと末尾のカンマは除外出来ない()
                Dim sql As String = "<sql>...set " _
                & "<if test='true'>ID = @Id,</if>" _
                & "<if test='true'>NAME = @Name,</if>" _
                & "<if test='false'>ADDRESS = @Address</if>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("...set  ID = @Id, NAME = @Name,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub setタグとifタグを使うと末尾のカンマを除外する()
                Dim sql As String = "<sql>...<set> " _
                & "<if test='true'>ID = @Id,</if>" _
                & "<if test='true'>NAME = @Name,</if>" _
                & "<if test='false'>ADDRESS = @Address</if>" _
                & "</set>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... SET ID = @Id, NAME = @Name", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub setタグとifタグを使うと先頭のカンマを除外する()
                Dim sql As String = "<sql>...<set> " _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='true'>, NAME = @Name</if>" _
                & "<if test='true'>, ADDRESS = @Address</if>" _
                & "</set>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... SET NAME = @Name , ADDRESS = @Address", xmlAnalyzer.AnalyzedSql)
            End Sub

        End Class

        Public Class joinタグ : Inherits SqlXmlAnalyzerTest
            <Test()> Public Sub joinタグ_separator属性なしならseparatorが空文字として動作する()
                Dim sql As String = "<sql>...<join property='@StrArray' />,,,</sql>"
                Dim vo As New FugaVo
                vo.StrArray = New String() {"a", "b", "c"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @StrArray#0@StrArray#1@StrArray#2 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_配列_String()
                Dim sql As String = "<sql>...<join property='@StrArray' separator=',' />,,,</sql>"
                Dim vo As New FugaVo
                vo.StrArray = New String() {"a", "b", "c"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @StrArray#0,@StrArray#1,@StrArray#2 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_List_String()
                Dim sql As String = "<sql>...<join property='@StrList' separator=',' />,,,</sql>"
                Dim vo As New FugaVo
                vo.StrList = New List(Of String)(New String() {"a", "b", "c"})
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @StrList#0,@StrList#1,@StrList#2 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_タグ間にプロパティ名括弧を指定すればより詳細()
                Dim sql As String = "<sql>...<join property='@StrArray' separator='/' >pre @StrArray() suf</join>,,,</sql>"
                Dim vo As New FugaVo
                vo.StrArray = New String() {"a", "b", "c"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... pre @StrArray#0 suf/pre @StrArray#1 suf/pre @StrArray#2 suf ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_配列直接_String()
                Dim sql As String = "<sql>...<join property='@Value' separator=',' />,,,</sql>"
                Dim param As String() = New String() {"a", "b", "c"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, param, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @Value#0,@Value#1,@Value#2 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_List直接_String()
                Dim sql As String = "<sql>...<join property='@Value' separator=',' />,,,</sql>"
                Dim param As New List(Of String)(New String() {"a", "b", "c"})
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, param, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @Value#0,@Value#1,@Value#2 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_タグ間にプロパティ名括弧を指定すればより詳細_直接配列()
                Dim sql As String = "<sql>...<join property='@Value' separator='/' >pre @Value() suf</join>,,,</sql>"
                Dim param As String() = New String() {"a", "b", "c"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, param, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... pre @Value#0 suf/pre @Value#1 suf/pre @Value#2 suf ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_タグ間に指定したスペースも繰り返される()
                Dim sql As String = "<sql>...<join property='@StrArray'>@StrArray() </join>,,,</sql>"
                Dim vo As New FugaVo
                vo.StrArray = New String() {"a", "b"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @StrArray#0 @StrArray#1  ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_配列_HogeVo()
                Dim sql As String = "<sql>...<join property='@HogeArray' separator=','>@HogeArray().Name</join>,,,</sql>"
                Dim vo As New FugaVo
                vo.HogeArray = New HogeVo() {New HogeVo, New HogeVo}
                vo.HogeArray(0).Name = "one"
                vo.HogeArray(1).Name = "two"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @HogeArray#0$Name,@HogeArray#1$Name ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_List_HogeVo()
                Dim sql As String = "<sql>...<join property='@HogeList' separator=','>@HogeList().Name</join>,,,</sql>"
                Dim vo As New FugaVo
                vo.HogeList = New List(Of HogeVo)(New HogeVo() {New HogeVo, New HogeVo})
                vo.HogeList(0).Name = "one"
                vo.HogeList(1).Name = "two"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @HogeList#0$Name,@HogeList#1$Name ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_配列直接_HogeVo()
                Dim sql As String = "<sql>...<join property='@Value' separator=','>@Value().Name</join>,,,</sql>"
                Dim vos As HogeVo() = New HogeVo() {New HogeVo, New HogeVo}
                vos(0).Name = "one"
                vos(1).Name = "two"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vos, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @Value#0$Name,@Value#1$Name ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_List直接_HogeVo()
                Dim sql As String = "<sql>...<join property='@Value' separator=','>@Value().Name</join>,,,</sql>"
                Dim vos As New List(Of HogeVo)(New HogeVo() {New HogeVo, New HogeVo})
                vos(0).Name = "one"
                vos(1).Name = "two"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vos, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @Value#0$Name,@Value#1$Name ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_配列_Integer()
                Dim sql As String = "<sql>...<join property='@IntArray' separator=',' />,,,</sql>"
                Dim vo As New FugaVo
                vo.IntArray = New Integer() {5, 6, 7}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @IntArray#0,@IntArray#1,@IntArray#2 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_配列VOの中の配列_String()
                Dim sql As String = "<sql>...<join property='@FugaArray(0).HogeArray' separator=','>@FugaArray(0).HogeArray().Name</join>,,,</sql>"
                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo}
                vo.FugaArray(0).HogeArray = New HogeVo() {New HogeVo, New HogeVo}
                vo.FugaArray(0).HogeArray(0).Name = "one"
                vo.FugaArray(0).HogeArray(1).Name = "two"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @FugaArray#0$HogeArray#0$Name,@FugaArray#0$HogeArray#1$Name ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub Joinタグのネスト_2階層_配列Voの中のStr配列()
                Dim sql As String = "<sql>...<join property='@FugaArray' separator=';'><join property='@FugaArray().StrArray' separator=',' /></join>,,,</sql>"
                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo, New FugaVo}
                vo.FugaArray(0).StrArray = New String() {"a", "b"}
                vo.FugaArray(1).StrArray = New String() {"c", "d"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @FugaArray#0$StrArray#0,@FugaArray#0$StrArray#1;@FugaArray#1$StrArray#0,@FugaArray#1$StrArray#1 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub Joinタグのネスト_2階層_配列Voの中のTextとStr配列()
                Dim sql As String = "<sql>...<join property='@FugaArray' separator=';'>@FugaArray().Name <join property='@FugaArray().StrArray' separator=',' /></join>,,,</sql>"
                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo, New FugaVo}
                vo.FugaArray(0).StrArray = New String() {"a", "b"}
                vo.FugaArray(1).StrArray = New String() {"c", "d"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @FugaArray#0$Name @FugaArray#0$StrArray#0,@FugaArray#0$StrArray#1;@FugaArray#1$Name @FugaArray#1$StrArray#0,@FugaArray#1$StrArray#1 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub Joinタグのネスト_2階層目にif_配列Voの中のStr配列()
                Dim sql As String = "<sql>...<join property='@FugaArray' separator=';'>" _
                                    & "<if test='@FugaArray().StrArray != null and 1 < @FugaArray().StrArray.Length'>@FugaArray().StrArray(1)</if>" _
                                    & "<if test='@FugaArray().StrArray != null and 0 < @FugaArray().StrArray.Length'>@FugaArray().StrArray(0)</if>" _
                                    & "<if test='@FugaArray().Name != null'>@FugaArray().Name</if>" _
                                    & "</join>,,,</sql>"
                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo, New FugaVo}
                vo.FugaArray(0).Name = "a"
                vo.FugaArray(1).Name = "b"
                vo.FugaArray(1).StrArray = New String() {"c", "d"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(XmlInequality.ConvInequality(sql), vo, New SqlExpressionEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @FugaArray#0$Name;@FugaArray#1$StrArray#1@FugaArray#1$StrArray#0@FugaArray#1$Name ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub Joinタグのネスト_3階層_配列Voの中の配列Voの中のStr配列_シンプル_3階層目だけn件()
                Dim sql As String = "<sql>...<join property='@FugaArray' separator='['>" _
                                    & "<join property='@FugaArray().FugaArray' separator='{' >" _
                                        & "<join property='@FugaArray().FugaArray().StrArray' separator='^' /></join></join>,,,</sql>"
                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo}
                vo.FugaArray(0).FugaArray = New FugaVo() {New FugaVo}
                vo.FugaArray(0).FugaArray(0).StrArray = New String() {"a", "b"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @FugaArray#0$FugaArray#0$StrArray#0^@FugaArray#0$FugaArray#0$StrArray#1 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub Joinタグのネスト_3階層_配列Voの中の配列Voの中のStr配列_シンプル_2階層目3階層目をn件()
                Dim sql As String = "<sql>...<join property='@FugaArray' separator='['>" _
                                    & "<join property='@FugaArray().FugaArray' separator='{' >" _
                                        & "<join property='@FugaArray().FugaArray().StrArray' separator='^' /></join></join>,,,</sql>"
                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo}
                vo.FugaArray(0).FugaArray = New FugaVo() {New FugaVo, New FugaVo}
                vo.FugaArray(0).FugaArray(0).StrArray = New String() {"a", "b"}
                vo.FugaArray(0).FugaArray(1).StrArray = New String() {"c", "d"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @FugaArray#0$FugaArray#0$StrArray#0^@FugaArray#0$FugaArray#0$StrArray#1{@FugaArray#0$FugaArray#1$StrArray#0^@FugaArray#0$FugaArray#1$StrArray#1 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub Joinタグのネスト_3階層_配列Voの中の配列Voの中のStr配列_複雑()
                Dim sql As String = "<sql>...<join property='@FugaArray' separator='['>" _
                                    & "<join property='@FugaArray().FugaArray' separator='{' >" _
                                        & "<join property='@FugaArray().FugaArray().StrArray' separator='^' /></join></join>,,,</sql>"
                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo, New FugaVo}
                vo.FugaArray(0).FugaArray = New FugaVo() {New FugaVo}
                vo.FugaArray(1).FugaArray = New FugaVo() {New FugaVo, New FugaVo}
                vo.FugaArray(0).FugaArray(0).StrArray = New String() {"a", "b"}
                vo.FugaArray(1).FugaArray(1).StrArray = New String() {"c"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @FugaArray#0$FugaArray#0$StrArray#0^@FugaArray#0$FugaArray#0$StrArray#1[{@FugaArray#1$FugaArray#1$StrArray#0 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub Joinタグのネスト_3階層目にif_配列Voの中のStr配列()
                Dim sql As String = "<sql>...<join property='@FugaArray' separator=';'>" _
                                    & "<join property='@FugaArray().FugaArray' separator=':' >" _
                                        & "<if test='@FugaArray().FugaArray().StrArray != null and 1 < @FugaArray().FugaArray().StrArray.Length'>@FugaArray().FugaArray().StrArray(1)</if>" _
                                        & "<if test='@FugaArray().FugaArray().StrArray != null and 0 < @FugaArray().FugaArray().StrArray.Length'>@FugaArray().FugaArray().StrArray(0)</if>" _
                                        & "<if test='@FugaArray().FugaArray().Name != null'>@FugaArray().FugaArray().Name</if>" _
                                    & "</join></join>,,,</sql>"
                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo, New FugaVo}
                vo.FugaArray(0).FugaArray = New FugaVo() {New FugaVo}
                vo.FugaArray(0).FugaArray(0).StrArray = New String() {"a", "b"}

                vo.FugaArray(1).FugaArray = New FugaVo() {New FugaVo, New FugaVo}
                vo.FugaArray(1).FugaArray(0).Name = "e"
                vo.FugaArray(1).FugaArray(1).StrArray = New String() {"d"}
                vo.FugaArray(1).FugaArray(1).Name = "f"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(XmlInequality.ConvInequality(sql), vo, New SqlExpressionEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @FugaArray#0$FugaArray#0$StrArray#1@FugaArray#0$FugaArray#0$StrArray#0" _
                                & ";@FugaArray#1$FugaArray#0$Name" _
                                    & ":@FugaArray#1$FugaArray#1$StrArray#0@FugaArray#1$FugaArray#1$Name ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_括弧付きの要素を使っていないとエラー()
                Dim sql As String = "<sql>...<join property='@FugaArray' separator=';'>@FugaArray<if test='@Value'>aiueo</if></join>,,,</sql>"
                Dim vo As New FugaVo
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                Try
                    xmlAnalyzer.Analyze()
                    Assert.Fail()

                Catch expected As InvalidOperationException
                    Assert.AreEqual("<join> : '@FugaArray' の一要素を表す '@FugaArray()' を使用すべき. @FugaArray<if test=""@Value"">aiueo</if>", expected.Message)
                End Try
            End Sub

            <Test()> Public Sub joinタグ_property属性にnullをもつプロパティ名だったら_処理しない_無視する()
                Dim sql As String = "<sql>...<join property='@StrArray' separator=',' />,,,</sql>"
                Dim vo As New FugaVo
                vo.StrArray = Nothing
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("...  ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub joinタグ_CollectionObject_Stringを解決できる()
                Const sql As String = "<sql>...<join property='@StrCollection' separator=',' />,,,</sql>"
                Dim vo As New FugaVo
                vo.StrCollection = New FugaVo.StrCollectionObject({"a", "b", "c"})
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @StrCollection#0,@StrCollection#1,@StrCollection#2 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <TestCase({"10", "9", "8", "7"}, "... @Name0#0,@Name0#1,@Name0#2,@Name0#3 ,,,")>
            <TestCase({"ヒャア がまんできねえ 0だ！"}, "... @Name0#0 ,,,")>
            Public Sub joinタグ_Criteria_Anyで設定した値の数に応じて展開できる(searchValues As String(), expected As String)
                Const sql As String = "<sql>...<join property='@Name0' separator=',' />,,,</sql>"
                Dim vo As New FugaVo
                Dim criteria As New Criteria(Of FugaVo)(vo)
                criteria.Any(vo.Name, searchValues)
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, criteria, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.That(xmlAnalyzer.AnalyzedSql, [Is].EqualTo(expected))
            End Sub

        End Class

        Public Class その他 : Inherits SqlXmlAnalyzerTest
            <Test()> Public Sub タグは大文字でも動く_念の為の機能()
                Dim sql As String = "<SQL>...<SET> " _
                & "<IF test='true'>ID = @Id,</IF>" _
                & "<IF test='true'>NAME = @Name,</IF>" _
                & "<IF test='false'>ADDRESS = @Address</IF>" _
                & "</SET>" _
                & "<WHERE>" _
                & "<IF test='false'>ID = @Id</IF>" _
                & "<IF test='true'>AND NAME = @Name</IF>" _
                & "<IF test='true'>AND ADDRESS = @Address</IF>" _
                & "</WHERE>" _
                & "</SQL>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... SET ID = @Id, NAME = @Name WHERE NAME = @Name AND ADDRESS = @Address", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub 配列の要素番号でアクセス_String型()
                Dim sql As String = "<sql>... @StrArray(0) ,,,</sql>"
                Dim vo As New FugaVo
                vo.StrArray = New String() {"a", "b"}
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @StrArray#0 ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub 配列の要素番号でアクセス_Vo型()
                Dim sql As String = "<sql>... @FugaArray(0).Name ,,,</sql>"
                Dim vo As New FugaVo
                vo.FugaArray = New FugaVo() {New FugaVo}
                vo.FugaArray(0).Name = "aaa"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, vo, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... @FugaArray#0$Name ,,,", xmlAnalyzer.AnalyzedSql)
            End Sub

        End Class

        Public Class andタグ : Inherits SqlXmlAnalyzerTest
            <Test()> Public Sub andタグを使わないでifタグだけだと先頭のANDは除外出来ない()
                Dim sql As String = "<sql>...and (" _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='true'>AND NAME = @Name</if>" _
                & "<if test='true'>AND ADDRESS = @Address</if>" _
                & ")</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("...and ( AND NAME = @Name AND ADDRESS = @Address )", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub andタグとifタグを使うと先頭のANDを除外する()
                Dim sql As String = "<sql>...<and>" _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='true'>AND NAME = @Name</if>" _
                & "<if test='true'>AND ADDRESS = @Address</if>" _
                & "</and>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... AND (NAME = @Name AND ADDRESS = @Address)", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub andタグとifタグを使うと先頭のORを除外する()
                Dim sql As String = "<sql>...<and>" _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='true'>OR NAME = @Name</if>" _
                & "<if test='true'>OR ADDRESS = @Address</if>" _
                & "</and>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... AND (NAME = @Name OR ADDRESS = @Address)", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub andタグを使わないでifタグだけで全falseだとand句除去できない()
                Dim sql As String = "<sql>...and (" _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='false'>AND NAME = @Name</if>" _
                & "<if test='false'>AND ADDRESS = @Address</if>" _
                & ")</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("...and ( )", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub andタグとifタグを使い全falseだとand句除去()
                Dim sql As String = "<sql>...<and> " _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='false'>AND NAME = @Name</if>" _
                & "<if test='false'>AND ADDRESS = @Address</if>" _
                & "</and>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("...", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub andタグ内にifタグとifに囲まれていない条件分があり_ifのtrueの結果で先頭のANDを除外する_ifが先()
                Dim sql As String = "<sql>...<and> " _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='true'>AND NAME = @Name</if>" _
                & "AND ADDRESS = @Address" _
                & "</and>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... AND (NAME = @Name AND ADDRESS = @Address)", xmlAnalyzer.AnalyzedSql)
            End Sub

            <Test()> Public Sub andタグ内にifタグとifに囲まれていない条件分があり_ifは全falseで_ifに囲まれていない条件文のANDを除外する_先頭にwhitespace()
                Dim sql As String = "<sql>...<and> " _
                & "<if test='false'>ID = @Id</if>" _
                & "<if test='false'>AND NAME = @Name</if>" _
                & "    AND ADDRESS = @Address" _
                & "</and>" _
                & "</sql>"
                Dim xmlAnalyzer As New SqlXmlAnalyzer(sql, Nothing, New FakeEvaluator)
                xmlAnalyzer.Analyze()
                Assert.AreEqual("... AND (ADDRESS = @Address)", xmlAnalyzer.AnalyzedSql)
            End Sub

        End Class
    End Class
End Namespace
