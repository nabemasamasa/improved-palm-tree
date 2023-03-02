Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports Fhi.Fw.Db.Impl
Imports NUnit.Framework

Namespace Db

    <TestFixture()> Public MustInherit Class SqlUtilTest

        Public Class [Default] : Inherits SqlUtilTest

            <Test()> Public Sub MakeWhereClause_VoがEmptyだから空文字()
                Dim vo As New SampleVo
                Dim actual As String = SqlUtil.MakeWhereClause(vo)
                Assert.AreEqual(String.Empty, actual)
            End Sub

            <Test()> Public Sub MakeWhereClause_VoのInt値で検索条件()
                Dim vo As New SampleVo
                vo.HogeId = 123
                Dim actual As String = SqlUtil.MakeWhereClause(vo)
                Assert.AreEqual(" WHERE HOGE_ID = @HogeId", actual)
            End Sub

            <Test()> Public Sub MakeWhereClause_VoのString値とDateTimea値で検索条件()
                Dim vo As New SampleVo
                vo.HogeDate = DateTime.Parse("2010/07/07")
                vo.HogeName = "hogeName"
                Dim actual As String = SqlUtil.MakeWhereClause(vo)
                Assert.AreEqual(" WHERE HOGE_NAME = @HogeName AND HOGE_DATE = @HogeDate", actual)
            End Sub

            <Test()> Public Sub MakeWhereClause_VoのDecimal値で検索条件()
                Dim vo As New SampleVo
                vo.HogeDecimal = CDec("123.45678901")
                Dim actual As String = SqlUtil.MakeWhereClause(vo)
                Assert.AreEqual(" WHERE HOGE_DECIMAL = @HogeDecimal", actual)
            End Sub

            <Test()> Public Sub MakeWhereClause_VoのEnum値で検索条件()
                Dim vo As New SampleVo
                vo.HogeEnum = SampleEnum.B
                Dim actual As String = SqlUtil.MakeWhereClause(vo)
                Assert.AreEqual(" WHERE HOGE_ENUM = @HogeEnum", actual)
            End Sub

            <Test()> Public Sub MakeWhereClause_個別のDB項目名でwhere句を作成する()
                Dim vo As New SampleVo
                vo.HogeId = 123
                Dim propertyAndNames As New Dictionary(Of String, String)
                propertyAndNames.Add("HogeId", "HOGE_FUGA_ID")
                Dim actual As String = SqlUtil.MakeWhereClause(vo, propertyAndNames)
                Assert.AreEqual(" WHERE HOGE_FUGA_ID = @HogeId", actual)
            End Sub

            <Test()> Public Sub MakeWhereClauseOnly_Nullでも検索条件に含める()
                Dim vo As New SampleVo
                vo.HogeDecimal = CDec("123.45678901")
                Dim onlyFields As New List(Of PropertyInfo)
                onlyFields.Add(vo.GetType.GetProperty("HogeId"))
                onlyFields.Add(vo.GetType.GetProperty("HogeName"))

                Dim actual As String = SqlUtil.MakeWhereClauseOnly(vo, onlyFields)
                Assert.AreEqual(" WHERE HOGE_ID = @HogeId AND HOGE_NAME = @HogeName", actual)
            End Sub

            <Test()> Public Sub MakeWhereClauseOnly_個別のDB項目名でwhere句を作成する()
                Dim vo As New SampleVo
                vo.HogeDecimal = CDec("123.45678901")
                Dim onlyFields As New List(Of PropertyInfo)
                onlyFields.Add(vo.GetType.GetProperty("HogeId"))
                onlyFields.Add(vo.GetType.GetProperty("HogeName"))

                Dim propertyAndNames As New Dictionary(Of String, String)
                propertyAndNames.Add("HogeId", "HOGE_FUGA_ID")

                Dim actual As String = SqlUtil.MakeWhereClauseOnly(vo, onlyFields, propertyAndNames)
                Assert.AreEqual(" WHERE HOGE_FUGA_ID = @HogeId AND HOGE_NAME = @HogeName", actual)
            End Sub

            <Test()> Public Sub MakeInsertInto_VOが空でも作れる()
                Dim vo As New SampleVo
                Dim actual As String = SqlUtil.MakeInsertInto(vo)
                Assert.AreEqual("(HOGE_ID, HOGE_NAME, HOGE_DATE, HOGE_DECIMAL, HOGE_ENUM) VALUES (@HogeId, @HogeName, @HogeDate, @HogeDecimal, @HogeEnum)", actual)
            End Sub

            <Test()> Public Sub MakeInsertInto_VOが空でも作れる_n件()
                Dim actual As String = SqlUtil.MakeInsertInto({New SampleVo, New SampleVo, New SampleVo})
                Assert.AreEqual("(HOGE_ID, HOGE_NAME, HOGE_DATE, HOGE_DECIMAL, HOGE_ENUM) VALUES (@Value#0$HogeId, @Value#0$HogeName, @Value#0$HogeDate, @Value#0$HogeDecimal, @Value#0$HogeEnum), (@Value#1$HogeId, @Value#1$HogeName, @Value#1$HogeDate, @Value#1$HogeDecimal, @Value#1$HogeEnum), (@Value#2$HogeId, @Value#2$HogeName, @Value#2$HogeDate, @Value#2$HogeDecimal, @Value#2$HogeEnum)", actual)
            End Sub

            <Test()> Public Sub MakeInsertInto_VOが空でも作れる_デキャメライズ以外の項目名も個別に設定できる()
                Dim vo As New SampleVo
                Dim specifiedFieldNames As New Dictionary(Of String, String)
                specifiedFieldNames.Add("HogeId", "!HOGE-ID!") ' ←例外的な項目名を指定
                Dim actual As String = SqlUtil.MakeInsertInto(vo, specifiedFieldNames)
                Assert.AreEqual("(!HOGE-ID!, HOGE_NAME, HOGE_DATE, HOGE_DECIMAL, HOGE_ENUM) VALUES (@HogeId, @HogeName, @HogeDate, @HogeDecimal, @HogeEnum)", actual)
            End Sub

            <Test()> Public Sub MakeInsertIntoWithoutNull_VOが空だと例外()
                Dim vo As New SampleVo
                Try
                    SqlUtil.MakeInsertIntoWithoutNull(vo)
                Catch expected As ArgumentException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub MakeInsertIntoWithoutNull_プロパティ値がnull以外の項目で作成する()
                Dim vo As New SampleVo
                vo.HogeName = "aiueo"
                vo.HogeDate = DateTime.Now
                Dim actual As String = SqlUtil.MakeInsertIntoWithoutNull(vo)
                Assert.AreEqual("(HOGE_NAME, HOGE_DATE) VALUES (@HogeName, @HogeDate)", actual)
            End Sub

            <Test()> Public Sub MakeInsertIntoWithoutNull_プロパティ値がnull以外の項目で作成する_デキャメライズ以外の項目名も個別に設定できる()
                Dim vo As New SampleVo
                vo.HogeId = 123
                vo.HogeName = "aiueo"
                Dim specifiedFieldNames As New Dictionary(Of String, String)
                specifiedFieldNames.Add("HogeName", "!HOGE-NAME!") ' ←例外的な項目名を指定
                Dim actual As String = SqlUtil.MakeInsertIntoWithoutNull(vo, specifiedFieldNames)
                Assert.AreEqual("(HOGE_ID, !HOGE-NAME!) VALUES (@HogeId, @HogeName)", actual)
            End Sub

            <Test()> Public Sub MakeUpdateSetWhere_VOが空でも作れる()
                Dim vo As New SampleVo
                Dim pkFields As List(Of PropertyInfo) = EzUtil.NewList(Of PropertyInfo)(vo.GetType.GetProperty("HogeId"))
                Dim actual As String = SqlUtil.MakeUpdateSetWhere(vo, pkFields)
                Assert.AreEqual("HOGE_NAME = @HogeName, HOGE_DATE = @HogeDate, HOGE_DECIMAL = @HogeDecimal, HOGE_ENUM = @HogeEnum WHERE HOGE_ID = @HogeId", actual)
            End Sub

            <Test()> Public Sub MakeUpdateSetWhere_VOが空でも作れる_デキャメライズ以外の項目名も個別に設定できる()
                Dim vo As New SampleVo
                Dim specifiedFieldNames As New Dictionary(Of String, String)
                specifiedFieldNames.Add("HogeName", "!HOGE-NAME")
                specifiedFieldNames.Add("HogeId", "!HOGE-ID")
                Dim pkFields As List(Of PropertyInfo) = EzUtil.NewList(Of PropertyInfo)(vo.GetType.GetProperty("HogeId"))
                Dim actual As String = SqlUtil.MakeUpdateSetWhere(vo, pkFields, specifiedFieldNames)
                Assert.AreEqual("!HOGE-NAME = @HogeName, HOGE_DATE = @HogeDate, HOGE_DECIMAL = @HogeDecimal, HOGE_ENUM = @HogeEnum WHERE !HOGE-ID = @HogeId", actual)
            End Sub

            <Test()> Public Sub MakeUpdateSetWithoutNullAndWhere_VOが空だと例外()
                Dim vo As New SampleVo
                Dim pkFields As List(Of PropertyInfo) = EzUtil.NewList(Of PropertyInfo)(vo.GetType.GetProperty("HogeId"))
                Try
                    SqlUtil.MakeUpdateSetWithoutNullAndWhere(vo, pkFields)
                Catch expected As ArgumentException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub MakeUpdateSetWithoutNullAndWhere_Set句はプロパティ値がnull以外の項目で作成する()
                Dim vo As New SampleVo
                vo.HogeName = "aiueo"
                vo.HogeDate = DateTime.Now
                Dim pkFields As List(Of PropertyInfo) = EzUtil.NewList(Of PropertyInfo)(vo.GetType.GetProperty("HogeId"))
                Dim actual As String = SqlUtil.MakeUpdateSetWithoutNullAndWhere(vo, pkFields)
                Assert.AreEqual("HOGE_NAME = @HogeName, HOGE_DATE = @HogeDate WHERE HOGE_ID = @HogeId", actual)
            End Sub

            <Test()> Public Sub MakeUpdateSetWithoutNullAndWhere_Set句はプロパティ値がnull以外の項目で作成する_デキャメライズ以外の項目名も個別に設定できる()
                Dim vo As New SampleVo
                vo.HogeName = "aiueo"
                vo.HogeDate = DateTime.Now
                Dim specifiedFieldNames As New Dictionary(Of String, String)
                specifiedFieldNames.Add("HogeName", "!HOGE-NAME")
                specifiedFieldNames.Add("HogeId", "!HOGE-ID")
                Dim pkFields As List(Of PropertyInfo) = EzUtil.NewList(Of PropertyInfo)(vo.GetType.GetProperty("HogeId"))
                Dim actual As String = SqlUtil.MakeUpdateSetWithoutNullAndWhere(vo, pkFields, specifiedFieldNames)
                Assert.AreEqual("!HOGE-NAME = @HogeName, HOGE_DATE = @HogeDate WHERE !HOGE-ID = @HogeId", actual)
            End Sub

            <Test()> Public Sub MakeCreateTableForSQLite_PrimaryKey無し()
                Dim actual As String = SqlUtil.MakeCreateTableForSQLite(Of SampleVo)()
                Assert.AreEqual("CREATE TABLE SAMPLE (HOGE_ID INTEGER, HOGE_NAME TEXT, HOGE_DATE TEXT, HOGE_DECIMAL REAL, HOGE_ENUM INTEGER)", actual)
            End Sub

            <Test()> Public Sub MakeCreateTableForSQLite_PrimaryKey有り()
                Dim vo As New SampleVo
                Dim pkFields As List(Of PropertyInfo) = EzUtil.NewList(Of PropertyInfo)(vo.GetType.GetProperty("HogeId"))
                Dim actual As String = SqlUtil.MakeCreateTableForSQLite(Of SampleVo)(pkFields)
                Assert.AreEqual("CREATE TABLE SAMPLE (HOGE_ID INTEGER, HOGE_NAME TEXT, HOGE_DATE TEXT, HOGE_DECIMAL REAL, HOGE_ENUM INTEGER) " _
                                & "PRIMARY KEY (HOGE_ID)", actual)
            End Sub

        End Class

        Public Class MakeWhereClauseByCriteriaTest : Inherits SqlUtilTest

            Private vo As SampleVo
            Private criteria As CriteriaBinder

            <SetUp()> Public Sub SetUp()
                vo = New SampleVo
                criteria = New Criteria(Of SampleVo)(vo)
            End Sub

            <Test()> Public Sub 検索値と等しいかで検索するWHERE句が作成できる()
                criteria.Equal(vo.HogeName, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME = @HogeName0"))
            End Sub

            <Test()> Public Sub 検索値と等しいかで検索するWHERE句が作成できる_NULLならIS_NULLになる()
                criteria.Equal(vo.HogeName, Nothing)
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME IS NULL"))
            End Sub

            <Test()> Public Sub 検索値と等しくないかで検索するWHERE句が作成できる()
                criteria.Not.Equal(vo.HogeName, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE NOT HOGE_NAME = @HogeName0"))
            End Sub

            <Test()> Public Sub 検索値と等しくないかで検索するWHERE句が作成できる_NULLならIS_NOT_NULLになる()
                criteria.Not.Equal(vo.HogeName, Nothing)
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE NOT HOGE_NAME IS NULL"))
            End Sub

            <Test()> Public Sub 検索値のいずれかに等しいかで検索するWHERE句が作成できる()
                criteria.Any(vo.HogeName, {""})
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME IN (<join property='@HogeName0' separator=',' />)"))
            End Sub

            <Test()> Public Sub 検索値のいずれとも等しくないかで検索するWHERE句が作成できる()
                criteria.Not.Any(vo.HogeName, {""})
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE NOT HOGE_NAME IN (<join property='@HogeName0' separator=',' />)"))
            End Sub

            <Test()> Public Sub ワイルドカード検索するWHERE句が作成できる()
                criteria.[Like](vo.HogeName, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME LIKE @HogeName0"))
            End Sub

            <Test()> Public Sub 検索値より大きいかで検索するWHERE句が作成できる()
                criteria.GreaterThan(vo.HogeName, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE @HogeName0 < HOGE_NAME"))
            End Sub

            <Test()> Public Sub 検索値より小さいかで検索するWHERE句が作成できる()
                criteria.LessThan(vo.HogeName, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME < @HogeName0"))
            End Sub

            <Test()> Public Sub 検索値以上かで検索するWHERE句が作成できる()
                criteria.GreaterEqual(vo.HogeName, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE @HogeName0 <= HOGE_NAME"))
            End Sub

            <Test()> Public Sub 検索値以下かで検索するWHERE句が作成できる()
                criteria.LessEqual(vo.HogeName, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME <= @HogeName0"))
            End Sub

            <Test()> Public Sub 複数の条件を指定したら_ANDで繋がったWHERE句が作成できる()
                criteria.Equal(vo.HogeName, "").Not.Equal(vo.HogeId, 0).GreaterThan(vo.HogeDate, "").LessThan(vo.HogeDate, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME = @HogeName0 AND NOT HOGE_ID = @HogeId1 AND @HogeDate2 < HOGE_DATE AND HOGE_DATE < @HogeDate3"))
            End Sub

            <Test()> Public Sub AND検索するWHERE句が作成できる_括弧で囲める()
                criteria.Equal(vo.HogeName, "").And(Sub() criteria.Equal(vo.HogeName, "").Or.Equal(vo.HogeId, ""))
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME = @HogeName0 AND ( HOGE_NAME = @HogeName1 OR HOGE_ID = @HogeId2 )"))
            End Sub

            <Test()> Public Sub 括弧のあとに続けて条件を設定できる()
                criteria.Nest(Sub() criteria.Equal(vo.HogeId, "").Or.Equal(vo.HogeId, "")).Not.Equal(vo.HogeName, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE ( HOGE_ID = @HogeId0 OR HOGE_ID = @HogeId1 ) AND NOT HOGE_NAME = @HogeName2"))
            End Sub

            <Test()> Public Sub OR検索するWHERE句が作成できる()
                criteria.Equal(vo.HogeName, "").Or.Equal(vo.HogeName, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME = @HogeName0 OR HOGE_NAME = @HogeName1"))
            End Sub

            <Test()> Public Sub OR検索するWHERE句が作成できる_括弧で囲める()
                criteria.Equal(vo.HogeName, "").Or(Sub() criteria.Equal(vo.HogeName, "").Equal(vo.HogeId, ""))
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME = @HogeName0 OR ( HOGE_NAME = @HogeName1 AND HOGE_ID = @HogeId2 )"))
            End Sub

            <Test()> Public Sub NOT検索するWHERE句が作成できる()
                criteria.Not.Equal(vo.HogeName, "").Not.Equal(vo.HogeName, "").Or.Not.Equal(vo.HogeName, "")
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE NOT HOGE_NAME = @HogeName0 AND NOT HOGE_NAME = @HogeName1 OR NOT HOGE_NAME = @HogeName2"))
            End Sub

            <Test()> Public Sub NOT検索するWHERE句が作成できる_括弧で囲める()
                criteria.Equal(vo.HogeName, "").Not(Sub() criteria.Equal(vo.HogeName, "")).Or.Not(Sub() criteria.Equal(vo.HogeName, ""))
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME = @HogeName0 AND NOT ( HOGE_NAME = @HogeName1 ) OR NOT ( HOGE_NAME = @HogeName2 )"))
            End Sub

            <Test()> Public Sub 入れ子のWHERE句が作成できる()
                criteria.Equal(vo.HogeName, "") _
                        .And(Sub() criteria.Nest(Sub() criteria.Equal(vo.HogeId, "").Equal(vo.HogeId, "")) _
                                           .Or(Sub() criteria.Equal(vo.HogeId, "").Equal(vo.HogeId, "")))
                Dim actual As String = SqlUtil.MakeWhereClauseByCriteria(criteria, New Dictionary(Of String, String)())
                Assert.That(actual, [Is].EqualTo(" WHERE HOGE_NAME = @HogeName0 AND ( ( HOGE_ID = @HogeId1 AND HOGE_ID = @HogeId2 ) OR ( HOGE_ID = @HogeId3 AND HOGE_ID = @HogeId4 ) )"))
            End Sub

        End Class

        Public Class Escape処理 : Inherits SqlUtilTest

            <TestCase("%", "[%]")>
            <TestCase("_", "[_]")>
            <TestCase("[", "[[]")>
            Public Sub エスケープ処理する(ByVal inputValue As String, ByVal expectedValue As String)
                Assert.That(SqlUtil.EscapeForBindParameterOfLike(inputValue), [Is].EqualTo(expectedValue))
            End Sub

            <TestCase(",", ",")>
            <TestCase("'", "'")>
            <TestCase("""", """")>
            <TestCase("*", "*")>
            Public Sub エスケープ処理されない(ByVal inputValue As String, ByVal expectedValue As String)
                Assert.That(SqlUtil.EscapeForBindParameterOfLike(inputValue), [Is].EqualTo(expectedValue))
            End Sub

            <Test(), Sequential()> Public Sub 部分一致のワイルドカードを付与する()
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("A"), [Is].EqualTo("%A%"), "通常、部分一致")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("B*"), [Is].EqualTo("B%"), "前方一致にもなる")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("*C"), [Is].EqualTo("%C"), "後方一致にもなる")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("*D*"), [Is].EqualTo("%D%"), "部分一致にもなる")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("*"), [Is].EqualTo("%"), "何かしら入力した値を検索")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("\*"), [Is].EqualTo("%*%"), "\*で * を表す")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("\**"), [Is].EqualTo("*%"), "")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("*\*"), [Is].EqualTo("%*"), "")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("*\**"), [Is].EqualTo("%*%"), "")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("%"), [Is].EqualTo("%[%]%"), "%を検索")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("%*"), [Is].EqualTo("[%]%"), "%の前方一致")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("*%"), [Is].EqualTo("%[%]"), "%の後方一致")
            End Sub

            <Test(), Sequential()> Public Sub 前方一致のワイルドカードを付与する()
                Assert.That(SqlUtil.ConvToForwardMatchWildcard("A"), [Is].EqualTo("A%"), "通常、前方一致")
                Assert.That(SqlUtil.ConvToForwardMatchWildcard("B*"), [Is].EqualTo("B%"), "前方一致にもなる")
                Assert.That(SqlUtil.ConvToForwardMatchWildcard("*C"), [Is].EqualTo("%C"), "後方一致にもなる")
                Assert.That(SqlUtil.ConvToForwardMatchWildcard("*D*"), [Is].EqualTo("%D%"), "部分一致にもなる")
                Assert.That(SqlUtil.ConvToForwardMatchWildcard("*"), [Is].EqualTo("%"), "何かしら入力した値を検索")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("\*"), [Is].EqualTo("%*%"), "\*で * を表す")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("\**"), [Is].EqualTo("*%"), "")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("*\*"), [Is].EqualTo("%*"), "")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("*\**"), [Is].EqualTo("%*%"), "")
                Assert.That(SqlUtil.ConvToForwardMatchWildcard("%"), [Is].EqualTo("[%]%"), "%を検索")
                Assert.That(SqlUtil.ConvToForwardMatchWildcard("%*"), [Is].EqualTo("[%]%"), "%の前方一致")
                Assert.That(SqlUtil.ConvToForwardMatchWildcard("*%"), [Is].EqualTo("%[%]"), "%の後方一致")
            End Sub

            <Test(), Sequential()> Public Sub 後方一致のワイルドカードを付与する()
                Assert.That(SqlUtil.ConvToBackwardMatchWildcard("A"), [Is].EqualTo("%A"), "通常、後方一致")
                Assert.That(SqlUtil.ConvToBackwardMatchWildcard("B*"), [Is].EqualTo("B%"), "前方一致にもなる")
                Assert.That(SqlUtil.ConvToBackwardMatchWildcard("*C"), [Is].EqualTo("%C"), "後方一致にもなる")
                Assert.That(SqlUtil.ConvToBackwardMatchWildcard("*D*"), [Is].EqualTo("%D%"), "部分一致にもなる")
                Assert.That(SqlUtil.ConvToBackwardMatchWildcard("*"), [Is].EqualTo("%"), "何かしら入力した値を検索")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("\*"), [Is].EqualTo("%*%"), "\*で * を表す")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("\**"), [Is].EqualTo("*%"), "")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("*\*"), [Is].EqualTo("%*"), "")
                Assert.That(SqlUtil.ConvToPartialMatchWildcard("*\**"), [Is].EqualTo("%*%"), "")
                Assert.That(SqlUtil.ConvToBackwardMatchWildcard("%"), [Is].EqualTo("%[%]"), "%を検索")
                Assert.That(SqlUtil.ConvToBackwardMatchWildcard("%*"), [Is].EqualTo("[%]%"), "%の前方一致")
                Assert.That(SqlUtil.ConvToBackwardMatchWildcard("*%"), [Is].EqualTo("%[%]"), "%の後方一致")
            End Sub

        End Class

        Public Class MakeSelectClauseTest : Inherits SqlUtilTest

            Private Class Boolean3
                Public Property B1() As Boolean?
                Public Property B2() As Boolean?
                Public Property B3() As Boolean?
            End Class

            <Test()> Public Sub callbackでプロパティを指定して_DB項目名を取得できる()
                Dim actual As String() = SqlUtil.ConvertPropertyNameToDbFieldName(Of THogeFugaVo)(
                    Function(selection, vo) selection.Is(vo.HogeId, vo.HogeName, vo.HogeDate, vo.HogeDecimal), Nothing)

                Assert.That(actual, [Is].EqualTo({"HOGE_ID", "HOGE_NAME", "HOGE_DATE", "HOGE_DECIMAL"}))
            End Sub

            <Test()> Public Sub callbackでvoを指定して_DB項目名全てを取得できる()
                Dim actual As String() = SqlUtil.ConvertPropertyNameToDbFieldName(Of THogeFugaVo)(
                    Function(selection, vo) selection.Is(vo), Nothing)

                Assert.That(actual, [Is].EquivalentTo({"HOGE_ID", "HOGE_SUB", "HOGE_NAME", "HOGE_DATE", "HOGE_DECIMAL", "IS_HOGE", "HOGE_ENUM", "UPDATED_USER_ID", "UPDATED_DATE", "UPDATED_TIME"}))
            End Sub

            <Test()> Public Sub callbackがnullなら_NullOrEmptyになる()
                Dim actual As String() = SqlUtil.ConvertPropertyNameToDbFieldName(Of THogeFugaVo)(selectionCallback:=Nothing, specifiedNameByNormal:=Nothing)
                Assert.That(actual, [Is].Null.Or.Empty)
            End Sub

            <Test()> Public Sub プロパティ名とDB項目名が異なる例外パターンを適用して取得できる()
                Dim specifiedNameByNormal As New Dictionary(Of String, String) From {{"HogeId", "HOGE_123"}, {"HogeName", "HOGE_あいう"}}
                Dim actual As String() = SqlUtil.ConvertPropertyNameToDbFieldName(Of THogeFugaVo)(
                    Function(selection, vo) selection.Is(vo.HogeId, vo.HogeName, vo.HogeDate, vo.HogeDecimal), specifiedNameByNormal)

                Assert.That(actual, [Is].EqualTo({"HOGE_123", "HOGE_あいう", "HOGE_DATE", "HOGE_DECIMAL"}))
            End Sub

            <Test()> Public Sub Boolean項目3つ以上用に_delegateで指定できる()
                Dim actual As String() = SqlUtil.ConvertPropertyNameToDbFieldName(Of Boolean3)(
                    Function(selection, vo) selection.Is(Function() vo.B3).Is(Function() vo.B1).Is(Function() vo.B2), Nothing)

                Assert.That(actual, [Is].EqualTo({"B_3", "B_1", "B_2"}))
            End Sub

        End Class

    End Class
End Namespace
