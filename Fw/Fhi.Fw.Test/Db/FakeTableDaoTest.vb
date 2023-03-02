Imports Fhi.Fw.Domain
Imports Fhi.Fw.TestUtil.DebugString
Imports NUnit.Framework

Namespace Db
    Public MustInherit Class FakeTableDaoTest

#Region "Nested classes..."
        Protected Class StrPvo : Inherits PrimitiveValueObject(Of String)
            Public Sub New(value As String)
                MyBase.New(value)
            End Sub
        End Class
        Protected Class IntPvo : Inherits PrimitiveValueObject(Of Integer)
            Public Sub New(value As Integer)
                MyBase.New(value)
            End Sub
        End Class
        Protected Class SampleAdvVo : Inherits SampleVo
            Public Property HogePvo As StrPvo
            Public Property FugaPvo As IntPvo
        End Class
        Protected Class TestingDao : Inherits FakeTableDao(Of SampleAdvVo)
        End Class
        Private Class FooVo
            Public Property Id As Integer?
            Public Property Bar As BarVo
        End Class
        Private Class BarVo
            Public Property Name As String
        End Class
#End Region

        Private Overloads Function ToString(records As SampleAdvVo()) As String
            Dim maker As New DebugStringMaker(Of SampleAdvVo)(Function(defineBy As IDebugStringRuleBinder, vo As SampleAdvVo) _
                                                              defineBy.Bind(vo.HogeId, vo.HogeName, vo.HogeDecimal, vo.HogePvo))
            Return maker.MakeString(records)
        End Function

        Protected dao As TestingDao

        <SetUp()> Public Overridable Sub SetUp()
            dao = New TestingDao
            dao.MakePrimaryUniqueKey = Function(vo As SampleAdvVo) StringUtil.ToString(vo.HogeId)
        End Sub

        Public Class 初期化直後は : Inherits FakeTableDaoTest

            <Test()> Public Sub レコードゼロ件()
                Assert.AreEqual(0, dao.Records.Length)
            End Sub

            <Test()> Public Sub FindForList_全件検索しても結果ゼロ件()
                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(Function(vo As SampleAdvVo) True)
                Assert.IsNotNull(actuals)
                Assert.AreEqual(0, actuals.Count)
            End Sub

            <Test()> Public Sub FindForObject_どれでも一件検索しても見つからないので結果Null()
                Dim actual As SampleAdvVo = dao.FindForObject(Function(vo As SampleAdvVo) True)
                Assert.IsNull(actual)
            End Sub

            <Test()> Public Sub DeleteBy_全件削除_一致するわけないからゼロ件()
                Assert.AreEqual(0, dao.DeleteBy(Function(vo As SampleAdvVo) True))
            End Sub

            <Test()> Public Sub UpdateBy_全件更新_一致するわけないからゼロ件()
                Assert.AreEqual(0, dao.UpdateBy(New SampleAdvVo, Function(vo As SampleAdvVo) ""))
            End Sub

            <Test()> Public Sub UpdateIgnoreNullBy_全件更新_一致するわけないからゼロ件()
                Assert.AreEqual(0, dao.UpdateIgnoreNullBy(New SampleAdvVo, Function(vo As SampleAdvVo) ""))
            End Sub

            <Test()> Public Sub InsertBy_PKは重複しないから一件()
                Assert.AreEqual(1, dao.InsertBy(New SampleAdvVo With {.HogeId = 1, .HogeName = "a"}))
                Assert.AreEqual(1, dao.Records.Length, "登録できたから擬似テーブル上は1件")
                With dao.Records(0)
                    Assert.AreEqual(1, .HogeId)
                    Assert.AreEqual("a", .HogeName)
                End With
            End Sub

            <Test()> Public Sub InsertBy_Insert後に元のVoを変更してもRecordsには反映されない()
                Dim actual As New SampleAdvVo With {.HogeId = 1, .HogeName = "hoge"}

                dao.InsertBy(actual)

                actual.HogeId = 2
                actual.HogeName = "piyo"

                Assert.That(dao.Records.First.HogeId, [Is].EqualTo(1))
                Assert.That(dao.Records.First.HogeName, [Is].EqualTo("hoge"))
            End Sub

        End Class

        Public Class 三件データがある時 : Inherits FakeTableDaoTest

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                dao.InitializeTable(New SampleAdvVo With {.HogeId = 1, .HogeName = "one", .HogeDate = CDate("2012/03/04"), .HogeDecimal = 11.11D},
                                    New SampleAdvVo With {.HogeId = 2, .HogeName = "two", .HogeDate = CDate("05:06:07"), .HogeDecimal = 222D},
                                    New SampleAdvVo With {.HogeId = 3, .HogeName = "three", .HogeDate = CDate("2012/08/09 10:11:12")})
            End Sub

            <Test()> Public Sub FindForList_一致したデータを返す()
                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(Function(vo As SampleAdvVo) EzUtil.IsTrue(vo.HogeId = 2) OrElse EzUtil.IsTrue(vo.HogeId = 3) OrElse EzUtil.IsTrue(vo.HogeId = 4))
                Assert.IsNotNull(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("two", actuals(0).HogeName)
                Assert.AreEqual("three", actuals(1).HogeName)
            End Sub

            <Test()> Public Sub FindForObject_どれでも一件検索しても見つからないので結果Null()
                Dim actual As SampleAdvVo = dao.FindForObject(Function(vo As SampleAdvVo) EzUtil.IsTrue(vo.HogeId = 2))
                Assert.IsNotNull(actual)
                Assert.AreEqual(2, actual.HogeId)
                Assert.AreEqual("two", actual.HogeName)
            End Sub

            <Test()> Public Sub DeleteBy_全件削除は結果三件()
                Assert.AreEqual(3, dao.DeleteBy(Function(vo As SampleAdvVo) True))
                Assert.AreEqual(0, dao.Records.Length, "削除したんだから0件")
            End Sub

            <Test()> Public Sub DeleteBy_条件のVOと一致したデータだけ削除する_NULLの項目は条件から除外()
                dao.DeleteBy(New SampleAdvVo() With {.HogeId = 2, .HogeName = "two", .HogeDecimal = Nothing})

                Assert.That(ToString(dao.Records), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 null   " & vbCrLf &
                            "     3 'three'  null        null   "))
            End Sub

            <Test()> Public Sub UpdateBy_全件更新_Nullでも上書き()
                Assert.AreEqual(3, dao.UpdateBy(New SampleAdvVo With {.HogeName = "aa", .HogeDecimal = 3.33D}, Function(vo As SampleAdvVo) ""))
                Assert.AreEqual(3, dao.Records.Length)
                With dao.Records(0)
                    Assert.IsNull(.HogeId, "#UpdateByは更新値で丸々更新(Nullでも更新)")
                    Assert.AreEqual("aa", .HogeName)
                    Assert.IsNull(.HogeDate)
                    Assert.AreEqual(3.33D, .HogeDecimal)
                End With
                With dao.Records(2)
                    Assert.IsNull(.HogeId, "#UpdateByは更新値で丸々更新(Nullでも更新)")
                    Assert.AreEqual("aa", .HogeName)
                    Assert.IsNull(.HogeDate)
                    Assert.AreEqual(3.33D, .HogeDecimal)
                End With
            End Sub

            <Test()> Public Sub UpdateIgnoreNullBy_全件更新_Nullでは上書きしない()
                Assert.AreEqual(3, dao.UpdateIgnoreNullBy(New SampleAdvVo With {.HogeName = "aa", .HogeDecimal = 3.33D}, Function(vo As SampleAdvVo) ""))
                Assert.AreEqual(3, dao.Records.Length)
                With dao.Records(0)
                    Assert.AreEqual(1, .HogeId, "更新値がnullだから上書きされていない")
                    Assert.AreEqual("aa", .HogeName)
                    Assert.AreEqual(CDate("2012/03/04"), .HogeDate, "更新値がnullだから上書きされていない")
                    Assert.AreEqual(3.33D, .HogeDecimal)
                End With
                With dao.Records(1)
                    Assert.AreEqual(2, .HogeId)
                    Assert.AreEqual("aa", .HogeName)
                    Assert.AreEqual(CDate("05:06:07"), .HogeDate)
                    Assert.AreEqual(3.33D, .HogeDecimal)
                End With
                With dao.Records(2)
                    Assert.AreEqual(3, .HogeId)
                    Assert.AreEqual("aa", .HogeName)
                    Assert.AreEqual(CDate("2012/08/09 10:11:12"), .HogeDate)
                    Assert.AreEqual(3.33D, .HogeDecimal)
                End With
            End Sub

            <Test()> Public Sub InsertBy_登録済みのPKと重複するならエラー()
                Try
                    dao.InsertBy(New SampleAdvVo With {.HogeId = 2, .HogeName = "Fuga"})
                    Assert.Fail()
                Catch expected As InvalidProgramException
                    Assert.AreEqual("PK重複エラー. 2 は登録済み", expected.Message)
                End Try
            End Sub

            <Test()> Public Sub FindForObjectでデータを取得した後_FakeTable内の同一レコードを操作しても取得したデータに影響はない()
                Dim sut As SampleAdvVo = dao.FindForObject(Function(vo) 1.Equals(vo.HogeId))
                dao.UpdateByPk(New SampleAdvVo With {.HogeId = 1, .HogeName = "first"})

                Assert.That(sut.HogeId, [Is].EqualTo(1))
                Assert.That(sut.HogeName, [Is].EqualTo("one"))
            End Sub

            <Test()> Public Sub FindForListでデータを取得した後_FakeTable内の同一レコードを操作しても取得したデータに影響はない()
                Dim sut As New List(Of SampleAdvVo)
                sut.AddRange(dao.FindForList(Function(vo) vo.HogeName.StartsWith("t")))
                dao.UpdateByPk(New SampleAdvVo With {.HogeId = 2, .HogeName = "second"})
                dao.UpdateByPk(New SampleAdvVo With {.HogeId = 3, .HogeName = "third"})

                Assert.That(sut(0).HogeId, [Is].EqualTo(2))
                Assert.That(sut(0).HogeName, [Is].EqualTo("two"))
                Assert.That(sut(1).HogeId, [Is].EqualTo(3))
                Assert.That(sut(1).HogeName, [Is].EqualTo("three"))
            End Sub

        End Class

        Public Class FindForListTest_Criteriaで取得 : Inherits FakeTableDaoTest
            Private vo As SampleAdvVo
            Private criteria As Criteria(Of SampleAdvVo)

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                dao.InitializeTable(New SampleAdvVo With {.HogeId = 1, .HogeName = "one", .HogeDate = CDate("2012/03/04"), .HogeDecimal = 11.11D, .HogePvo = New StrPvo("ICHI")},
                                    New SampleAdvVo With {.HogeId = 2, .HogeName = "two", .HogeDate = CDate("05:06:07"), .HogeDecimal = 222D, .HogePvo = New StrPvo("NI")},
                                    New SampleAdvVo With {.HogeId = 3, .HogeName = "three", .HogeDate = CDate("2012/08/09 10:11:12"), .HogePvo = New StrPvo("SAN")})

                vo = New SampleAdvVo()
                criteria = New Criteria(Of SampleAdvVo)(vo)
            End Sub

            <Test()> Public Sub 一致で取得できる()
                criteria.Equal(vo.HogeName, "one")

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' "))
            End Sub

            <Test()> Public Sub 不一致で取得できる()
                criteria.Not.Equal(vo.HogeDecimal, 222D)

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' " & vbCrLf &
                            "     3 'three'  null        'SAN'  "))
            End Sub

            <Test()> Public Sub 含まれるかで取得できる()
                criteria.Any(vo.HogeId, {1, 2})

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' " & vbCrLf &
                            "     2 'two'         222    'NI'   "))
            End Sub

            <Test()> Public Sub 含まないかで取得できる()
                criteria.Not.Any(vo.HogeId, {1, 2})

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                    "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                    "     3 'three'  null        'SAN'  "))
            End Sub

            <Test()> Public Sub ワイルドカード検索で取得できる()
                criteria.Like(vo.HogeName, "t%")

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     2 'two'            222 'NI'   " & vbCrLf &
                            "     3 'three'  null        'SAN'  "))
            End Sub

            <Test()> Public Sub ワイルドカード検索で取得できる_PVO項目を検索()
                criteria.Like(vo.HogePvo, "%I")

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                    "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                    "     1 'one'          11.11 'ICHI' " & vbCrLf &
                    "     2 'two'         222    'NI'   "))
            End Sub

            <Test()> Public Sub NOTワイルドカード検索で取得できる()
                criteria.Not.Like(vo.HogeName, "t%")

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                    "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                    "     1 'one'          11.11 'ICHI' "))
            End Sub

            <Test()> Public Sub 比較で取得できる_検索値より上()
                criteria.GreaterThan(vo.HogeDate, CDate("2012/03/05"))

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     3 'three'  null        'SAN'  "))
            End Sub

            <Test()> Public Sub 比較で取得できる_検索値より下()
                criteria.LessThan(vo.HogeDate, CDate("2012/08/09"))

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' " & vbCrLf &
                            "     2 'two'         222    'NI'   "))
            End Sub

            <Test()> Public Sub 比較で取得できる_検索値以上()
                criteria.GreaterEqual(vo.HogeDate, CDate("2012/03/04"))

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' " & vbCrLf &
                            "     3 'three'  null        'SAN'  "))
            End Sub

            <Test()> Public Sub 比較で取得できる_検索値以下()
                criteria.LessEqual(vo.HogeDate, CDate("2012/03/04"))

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' " & vbCrLf &
                            "     2 'two'         222    'NI'   "))
            End Sub

            <Test()> Public Sub Null除外して比較で取得できる_検索値より上()
                criteria.GreaterThan(vo.HogeDate, CDate("2012/03/05"))
                dao.InsertBy(New SampleAdvVo With {.HogeId = 4, .HogeDate = Nothing})

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     3 'three'  null        'SAN'  "))
            End Sub

            <Test()> Public Sub Null除外して比較で取得できる_検索値より下()
                criteria.LessThan(vo.HogeDate, CDate("2012/08/09"))
                dao.InsertBy(New SampleAdvVo With {.HogeId = 4, .HogeDate = Nothing})

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' " & vbCrLf &
                            "     2 'two'         222    'NI'   "))
            End Sub

            <Test()> Public Sub Null除外して比較で取得できる_検索値以上()
                criteria.GreaterEqual(vo.HogeDate, CDate("2012/03/04"))
                dao.InsertBy(New SampleAdvVo With {.HogeId = 4, .HogeDate = Nothing})

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' " & vbCrLf &
                            "     3 'three'  null        'SAN'  "))
            End Sub

            <Test()> Public Sub Null除外して比較で取得できる_検索値以下()
                criteria.LessEqual(vo.HogeDate, CDate("2012/03/04"))
                dao.InsertBy(New SampleAdvVo With {.HogeId = 4, .HogeDate = Nothing})

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' " & vbCrLf &
                            "     2 'two'         222    'NI'   "))
            End Sub

            <Test()> Public Sub ANDで取得できる()
                criteria.Equal(vo.HogeId, 1).Equal(vo.HogeName, "one")

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' "))
            End Sub

            <Test()> Public Sub ORで取得できる()
                criteria.Equal(vo.HogeId, 2).Or.Equal(vo.HogeId, 3).Or.Equal(vo.HogeId, 4)

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     2 'two'            222 'NI'   " & vbCrLf &
                            "     3 'three'  null        'SAN'  "))
            End Sub

            <Test()> Public Sub ANDとORを繋げて書いたら_ANDが先に処理される()
                criteria.Equal(vo.HogeId, 1).Or.Equal(vo.HogeName, "two").Equal(vo.HogeId, 3)

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' "), "「HogeId = 1 OR (HogeName = 'two' AND HogeId = 3)」と同様になるはず")
            End Sub

            <Test()> Public Sub ネストした条件で取得できる_AND_OR_AND()
                '(条件1 AND 条件2) OR (条件3 AND 条件4)
                criteria.Nest(Sub() criteria.Equal(vo.HogeId, 1).Equal(vo.HogeName, "one")).Or(Sub() criteria.Equal(vo.HogeId, 2).Equal(vo.HogeName, "two"))

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' " & vbCrLf &
                            "     2 'two'         222    'NI'   "))
            End Sub

            <Test()> Public Sub ネストした条件で取得できる_OR_AND_OR()
                '(条件1 OR 条件2) AND (条件3 OR 条件4)
                criteria.Nest(Sub() criteria.Equal(vo.HogeId, 1).Or.Equal(vo.HogeId, 2)).And(Sub() criteria.Equal(vo.HogeName, "two").Or.Equal(vo.HogeName, "three"))

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     2 'two'            222 'NI'   "))
            End Sub

            <Test()> Public Sub すごく長い条件で動作テスト1()
                criteria.Nest(Sub() criteria.Equal(vo.HogeId, 1).Equal(vo.HogeName, "one")) _
                        .Or.Not.Equal(vo.HogeId, 4).Or.Equal(vo.HogeId, 2).Equal(vo.HogeName, "two") _
                        .Or(Sub() criteria.Equal(vo.HogeId, 5).Equal(vo.HogeName, "five").Or.Like(vo.HogeName, "thr%")) _
                        .Not.Any(vo.HogeName, {"four", "five"})

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     1 'one'          11.11 'ICHI' " & vbCrLf &
                            "     2 'two'         222    'NI'   " & vbCrLf &
                            "     3 'three'  null        'SAN'  "))
            End Sub

            <Test()> Public Sub すごく長い条件で動作テスト2()
                criteria.Nest(Sub() criteria.Equal(vo.HogeId, 11).Equal(vo.HogeName, "eleven")) _
                        .Or(Sub() criteria.Nest(Sub() criteria.Equal(vo.HogeId, 1).Or.Equal(vo.HogeId, 2)) _
                                          .And(Sub() criteria.Equal(vo.HogeName, "two").Or.Equal(vo.HogeName, "three")) _
                                          .Or(Sub() criteria.Equal(vo.HogeId, 6).Or.Equal(vo.HogeId, 7))) _
                        .Equal(vo.HogeEnum, Nothing).Not(Sub() criteria.Equal(vo.HogeId, 1).Or.Equal(vo.HogeId, 3))

                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(criteria)

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                            "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                            "     2 'two'            222 'NI'   "))
            End Sub

            <Test()> Public Sub FindForListなら_ラムダ式で条件指定して_一致するレコード全部取得できる()
                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(Function(criteria, vo) criteria.Not.Equal(vo.HogeName, "one"))

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                    "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                    "     2 'two'            222 'NI'   " & vbCrLf &
                    "     3 'three'  null        'SAN'  "))
            End Sub

            <Test()> Public Sub FindForObjectなら_ラムダ式で条件指定して_一致するレコードの先頭一件を取得できる()
                Dim actual As SampleAdvVo = dao.FindForObject(Function(criteria, vo) criteria.Not.Equal(vo.HogeName, "one"))

                Assert.That(ToString({actual}), [Is].EqualTo(
                    "HogeId HogeName HogeDecimal HogePvo" & vbCrLf &
                    "     2 'two'            222 'NI'   "))
            End Sub
        End Class

        Public Class FindTest_Criteriaで取得_PVOとPrimitive型 : Inherits FakeTableDaoTest

            <SetUp> Public Overrides Sub SetUp()
                MyBase.SetUp()
                dao.InitializeTable(New SampleAdvVo With {.HogeId = 123, .HogeName = "abc", .HogePvo = New StrPvo("xyz"), .FugaPvo = New IntPvo(789)})
            End Sub

            <Test> Public Sub Primitive型の項目をPVOで取得できる_Equal()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogeName, New StrPvo("abc"))))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogeName, New StrPvo("xyz"))))
            End Sub

            <Test> Public Sub PVOの項目をPrimitive型で取得できる_Equal()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogePvo, "xyz")))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogePvo, "abc")))
            End Sub

            <Test> Public Sub Primitive型の項目をPVOで取得できる_Any()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogeName, New Object() {New StrPvo("abc"), New StrPvo("def"), New StrPvo("ghi")})))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogeName, New Object() {New StrPvo("def"), New StrPvo("ghi"), New StrPvo("jkl")})))
            End Sub

            <Test> Public Sub PVOの項目をPrimitive型で取得できる_Any()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogePvo, New Object() {"rst", "uvw", "xyz"})))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogePvo, New Object() {"opq", "rst", "uvw"})))
            End Sub

            <Test> Public Sub Primitive型の項目をPVOで取得できる_Like()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogeName, New StrPvo("a%"))))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogeName, New StrPvo("x%"))))
            End Sub

            <Test> Public Sub PVOの項目をPrimitive型で取得できる_Like()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogePvo, "x%")))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogePvo, "a%")))
            End Sub

            <Test> Public Sub Primitive型の項目をPVOで取得できる_GreaterThan()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.GreaterThan(vo.HogeId, New IntPvo(122))))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.GreaterThan(vo.HogeId, New IntPvo(123))))
            End Sub

            <Test> Public Sub PVOの項目をPrimitive型で取得できる_GreaterThan()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.GreaterThan(vo.FugaPvo, 788)))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.GreaterThan(vo.FugaPvo, 789)))
            End Sub

            <Test> Public Sub Primitive型の項目をPVOで取得できる_LessThan()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.LessThan(vo.HogeId, New IntPvo(124))))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.LessThan(vo.HogeId, New IntPvo(123))))
            End Sub

            <Test> Public Sub PVOの項目をPrimitive型で取得できる_LessThan()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.LessThan(vo.FugaPvo, 790)))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.LessThan(vo.FugaPvo, 789)))
            End Sub

            <Test> Public Sub Primitive型の項目をPVOで取得できる_GreaterEqual()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.GreaterEqual(vo.HogeId, New IntPvo(123))))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.GreaterEqual(vo.HogeId, New IntPvo(124))))
            End Sub

            <Test> Public Sub PVOの項目をPrimitive型で取得できる_GreaterEqual()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.GreaterEqual(vo.FugaPvo, 789)))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.GreaterEqual(vo.FugaPvo, 790)))
            End Sub

            <Test> Public Sub Primitive型の項目をPVOで取得できる_LessEqual()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.LessEqual(vo.HogeId, New IntPvo(123))))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.LessEqual(vo.HogeId, New IntPvo(122))))
            End Sub

            <Test> Public Sub PVOの項目をPrimitive型で取得できる_LessEqual()
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.LessEqual(vo.FugaPvo, 789)))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.LessEqual(vo.FugaPvo, 788)))
            End Sub
        End Class

        Public Class FindTest_Criteriaで取得_IgnoreCase : Inherits FakeTableDaoTest

            <SetUp> Public Overrides Sub SetUp()
                MyBase.SetUp()
                dao.InitializeTable(New SampleAdvVo With {.HogeId = 123, .HogeName = "abc", .HogePvo = New StrPvo("xyz"), .FugaPvo = New IntPvo(789)},
                                    New SampleAdvVo With {.HogeName = "", .HogePvo = New StrPvo("")},
                                    New SampleAdvVo With {.HogeName = Nothing, .HogePvo = New StrPvo(Nothing)})
            End Sub

            <Test> Public Sub Equal_型_PVOなら内部型_が違うなら一見同値でも一致しないものと扱う()
                dao.IgnoreCaseOfCriteria = True
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogeId, "123")))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogeId, New StrPvo("123"))))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.FugaPvo, "789")))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.FugaPvo, New StrPvo("789"))))
            End Sub

            <Test> Public Sub Equal_IgnoreCaseTrueなら_大文字小文字を同一と見做す()
                dao.IgnoreCaseOfCriteria = True
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogeName, "ABC")))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogeName, New StrPvo("ABC"))))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogePvo, "Xyz")))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogePvo, New StrPvo("Xyz"))))
            End Sub

            <Test> Public Sub Equal_IgnoreCaseFalseなら_大文字小文字は別の文字と見做す()
                dao.IgnoreCaseOfCriteria = False
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogeName, "ABC")))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogeName, New StrPvo("ABC"))))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogePvo, "Xyz")))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogePvo, New StrPvo("Xyz"))))
            End Sub

            <TestCase("", True)>
            <TestCase("", False)>
            <TestCase(Nothing, True)>
            <TestCase(Nothing, False)>
            Public Sub Equal_空の値も取得できる(value As String, ignoreCase As Boolean)
                dao.IgnoreCaseOfCriteria = ignoreCase
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogeName, value)))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogeName, New StrPvo(value))))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogePvo, value)))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Equal(vo.HogePvo, New StrPvo(value))))
            End Sub

            <Test> Public Sub Any_型_PVOなら内部型_が違うなら一見同値でも一致しないものと扱う()
                dao.IgnoreCaseOfCriteria = True
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogeId, New String() {"123"})))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogeId, New StrPvo() {New StrPvo("123")})))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.FugaPvo, New String() {"789"})))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.FugaPvo, New StrPvo() {New StrPvo("789")})))
            End Sub

            <Test> Public Sub Any_IgnoreCaseTrueなら_大文字小文字を同一と見做す()
                dao.IgnoreCaseOfCriteria = True
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogeName, New String() {"ABC"})))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogeName, New StrPvo() {New StrPvo("ABC")})))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogePvo, New String() {"Xyz"})))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogePvo, New StrPvo() {New StrPvo("Xyz")})))
            End Sub

            <Test> Public Sub Any_IgnoreCaseFalseなら_大文字小文字は別の文字と見做す()
                dao.IgnoreCaseOfCriteria = False
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogeName, New String() {"ABC"})))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogeName, New StrPvo() {New StrPvo("ABC")})))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogePvo, New String() {"Xyz"})))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogePvo, New StrPvo() {New StrPvo("Xyz")})))
            End Sub

            <TestCase("", True)>
            <TestCase("", False)>
            <TestCase(Nothing, True)>
            <TestCase(Nothing, False)>
            Public Sub Any_空の値も取得できる(value As String, ignoreCase As Boolean)
                dao.IgnoreCaseOfCriteria = ignoreCase
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogeName, New String() {value})))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogeName, New StrPvo() {New StrPvo(value)})))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogePvo, New String() {value})))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Any(vo.HogePvo, New StrPvo() {New StrPvo(value)})))
            End Sub

            <Test> Public Sub Like_IgnoreCaseTrueなら_大文字小文字を同一と見做す()
                dao.IgnoreCaseOfCriteria = True
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogeName, "AB%")))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogePvo, "X_z")))
            End Sub

            <Test> Public Sub Like_IgnoreCaseFalseなら_大文字小文字は別の文字と見做す()
                dao.IgnoreCaseOfCriteria = False
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogeName, "AB%")))
                Assert.IsNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogePvo, "X_z")))
            End Sub

            <TestCase("", True)>
            <TestCase("", False)>
            <TestCase(Nothing, True)>
            <TestCase(Nothing, False)>
            Public Sub Like_空の値も取得できる(value As String, ignoreCase As Boolean)
                dao.IgnoreCaseOfCriteria = ignoreCase
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogeName, value)))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogeName, New StrPvo(value))))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogePvo, value)))
                Assert.IsNotNull(dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogePvo, New StrPvo(value))))
            End Sub
        End Class

        Public Class CloneRecordTest : Inherits FakeTableDaoTest
            Private Function CloneThatDeepCopy(ByVal vo As FooVo) As FooVo
                Dim result As FooVo = VoUtil.NewInstance(Of FooVo)(vo)
                result.Bar = VoUtil.NewInstance(Of BarVo)(vo.Bar)
                Return result
            End Function

            Private Shadows dao As FakeTableDao(Of FooVo)

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                dao = New FakeTableDao(Of FooVo)
                dao.InitializeTable(New FooVo With {.Id = 1, .Bar = New BarVo With {.Name = "ONE"}}, _
                                    New FooVo With {.Id = 2, .Bar = New BarVo With {.Name = "TWO"}}, _
                                    New FooVo With {.Id = 3, .Bar = New BarVo With {.Name = "THREE"}})
            End Sub

            <Test()> Public Sub FindForObject_実行の度にrecordは別インスタンスになる()
                Dim actual1 As FooVo = dao.FindForObject(Function(vo) 1.Equals(vo.Id))
                Dim actual2 As FooVo = dao.FindForObject(Function(vo) 1.Equals(vo.Id))
                Assert.That(actual1, [Is].Not.SameAs(actual2), "recordは別々のインスタンス")
                Assert.That(actual1.Bar, [Is].SameAs(actual2.Bar), "けれどShallowコピーなのでproperty値は同じインスタンス")
            End Sub
            <Test()> Public Sub FindForObject_実行の度にrecordは別インスタンスになる_CloneRecordに個別指定すれば_DeepCopyも可能()
                dao.CloneRecord = AddressOf CloneThatDeepCopy
                Dim actual1 As FooVo = dao.FindForObject(Function(vo) 1.Equals(vo.Id))
                Dim actual2 As FooVo = dao.FindForObject(Function(vo) 1.Equals(vo.Id))
                Assert.That(actual1, [Is].Not.SameAs(actual2), "recordは別々のインスタンス")
                Assert.That(actual1.Bar, [Is].Not.SameAs(actual2.Bar), "デリゲートをDeepCopyにしたので別々のインスタンス")
            End Sub
            <Test()> Public Sub FindForObject_実行の度にrecordは別インスタンスになる_CloneRecordに個別指定すれば_Cloneしないことも可能()
                dao.CloneRecord = Function(vo) vo
                Dim actual1 As FooVo = dao.FindForObject(Function(vo) 1.Equals(vo.Id))
                Dim actual2 As FooVo = dao.FindForObject(Function(vo) 1.Equals(vo.Id))
                Assert.That(actual1, [Is].SameAs(actual2), "Cloneしてないので同じインスタンス")
            End Sub

            <Test()> Public Sub InitializeTable_別インスタンスになる()
                Dim actual As New FooVo With {.Id = 1, .Bar = New BarVo With {.Name = "ONE"}}
                dao.InitializeTable(actual)
                Assert.That(dao.Records.First, [Is].Not.SameAs(actual), "recordsと元のVoは別インスタンス")
            End Sub

            <Test()> Public Sub InitializeTable_CloneRecordを指定すれば同インスタンスにできる()
                Dim actual As New FooVo With {.Id = 1, .Bar = New BarVo With {.Name = "ONE"}}
                dao.CloneRecord = Function(vo) vo
                dao.InitializeTable(actual)
                Assert.That(dao.Records.First, [Is].SameAs(actual), "CloneRecordsでそのまま返す様にしてるから同じインスタンス")
            End Sub

        End Class

        Public Class MakeTheMakeKeyCallbackItsOnlyHasValueTest : Inherits FakeTableDaoTest

            <Test()> Public Sub 入力値でのMakeKeyCallbackを作成する_1項目()
                Dim vo As New SampleAdvVo With {.HogeName = "a"}
                Assert.That(FakeTableDao(Of SampleAdvVo).MakeTheMakeKeyCallbackItsOnlyHasValue(vo).Invoke(vo), [Is].EqualTo("a"))
            End Sub

            <Test()> Public Sub 入力値でのMakeKeyCallbackを作成する_有効値無し()
                Dim vo As New SampleAdvVo
                Assert.That(FakeTableDao(Of SampleAdvVo).MakeTheMakeKeyCallbackItsOnlyHasValue(vo).Invoke(vo), [Is].Null)
            End Sub

            <Test()> Public Sub 入力値でのMakeKeyCallbackを作成する_n項目()
                Dim vo As New SampleAdvVo With {.HogeId = 1, .HogeDecimal = 2.34D, .HogeDate = CDate("2012/03/04")}
                Assert.That(FakeTableDao(Of SampleAdvVo).MakeTheMakeKeyCallbackItsOnlyHasValue(vo).Invoke(vo), [Is].EqualTo(EzUtil.MakeKey(vo.HogeId, vo.HogeDate, vo.HogeDecimal)))
            End Sub

        End Class

        Public Class MakeTheMatchCallbackItsOnlyHasValueTest : Inherits FakeTableDaoTest

            <Test()> Public Sub 入力値でのMatchCallbackを作成する_1項目()
                Dim vo As New SampleAdvVo With {.HogeName = "a"}
                Dim result As FakeTableDao(Of SampleAdvVo).MatchCallback = FakeTableDao(Of SampleAdvVo).MakeTheMatchCallbackItsOnlyHasValue(vo)
                Assert.That(result.Invoke(vo), [Is].True)
                Assert.That(result.Invoke(New SampleAdvVo With {.HogeName = "a"}), [Is].True, "インスタンスが別でも一致")
            End Sub

            <Test()> Public Sub 入力値でのMatchCallbackを作成する_有効値無し()
                Dim vo As New SampleAdvVo
                Dim result As FakeTableDao(Of SampleAdvVo).MatchCallback = FakeTableDao(Of SampleAdvVo).MakeTheMatchCallbackItsOnlyHasValue(vo)
                Assert.That(result.Invoke(vo), [Is].True)
                Assert.That(result.Invoke(New SampleAdvVo With {.HogeName = "a"}), [Is].True, "値なしだから、全一致")
                Assert.That(result.Invoke(New SampleAdvVo With {.HogeId = 1}), [Is].True, "値なしだから、全一致")
                Assert.That(result.Invoke(New SampleAdvVo With {.HogeId = 2, .HogeDecimal = 2.3D, .HogeName = "w"}), [Is].True, "値なしだから、全一致")
            End Sub

            <Test()> Public Sub 入力値でのMatchCallbackを作成する_n項目()
                Dim vo As New SampleAdvVo With {.HogeId = 1, .HogeDecimal = 2.34D, .HogeDate = CDate("2012/03/04")}
                Dim result As FakeTableDao(Of SampleAdvVo).MatchCallback = FakeTableDao(Of SampleAdvVo).MakeTheMatchCallbackItsOnlyHasValue(vo)
                Assert.That(result.Invoke(vo), [Is].True)
                Assert.That(result.Invoke(New SampleAdvVo With {.HogeId = 1, .HogeDecimal = 2.34D, .HogeDate = CDate("2012/03/04")}), [Is].True, "インスタンスが別でも一致")
            End Sub

        End Class

        Public Class MakeTheMatchCallbackTest : Inherits FakeTableDaoTest

            <TestCase("Abc", True)>
            <TestCase("ABC", False)>
            <TestCase("abc", False)>
            <TestCase(Nothing, False)>
            Public Sub 検索条件でのMatchCallbackを作成する_1項目(hogeName As String, expected As Boolean)
                Dim vo As New SampleAdvVo
                Dim criteria As New Criteria(Of SampleAdvVo)(vo)
                criteria.Equal(vo.HogeName, hogeName)

                Dim result As FakeTableDao(Of SampleAdvVo).MatchCallback = FakeTableDao(Of SampleAdvVo).MakeTheMatchCallback(criteria, ignoreCase:=False)

                Assert.That(result.Invoke(New SampleAdvVo With {.HogeName = "Abc"}), [Is].EqualTo(expected))
            End Sub

            <Test()> Public Sub 検索条件でのMatchCallbackを作成する_条件未設定()
                Dim vo As New SampleAdvVo
                Dim criteria As New Criteria(Of SampleAdvVo)(vo)

                Dim result As FakeTableDao(Of SampleAdvVo).MatchCallback = FakeTableDao(Of SampleAdvVo).MakeTheMatchCallback(criteria, ignoreCase:=False)

                Assert.True(result.Invoke(New SampleAdvVo With {.HogeName = "a"}))
                Assert.True(result.Invoke(New SampleAdvVo With {.HogeId = 1}))
                Assert.True(result.Invoke(New SampleAdvVo With {.HogeId = 2, .HogeDecimal = 2.3D, .HogeName = "w"}))
            End Sub

            <Test()> Public Sub 検索条件でのMatchCallbackを作成する_複数項目()
                Dim vo As New SampleAdvVo
                Dim criteria As New Criteria(Of SampleAdvVo)(vo)
                criteria.Equal(vo.HogeId, 1).Equal(vo.HogeDecimal, 2.34D).Equal(vo.HogeDate, CDate("2012/03/04"))

                Dim result As FakeTableDao(Of SampleAdvVo).MatchCallback = FakeTableDao(Of SampleAdvVo).MakeTheMatchCallback(criteria, ignoreCase:=False)

                Assert.True(result.Invoke(New SampleAdvVo With {.HogeId = 1, .HogeDecimal = 2.34D, .HogeDate = CDate("2012/03/04")}))
            End Sub
        End Class

        Public Class ConvSqlWildcardToDotnetWildcardTest : Inherits FakeTableDaoTest

            <Test()> Public Sub 任意の1文字を表すキーワードを変換できる()
                Assert.That(TestingDao.ConvSqlWildcardToDotnetWildcard("AB_CD_EF"), [Is].EqualTo("AB?CD?EF"))
            End Sub

            <Test()> Public Sub 任意の文字列を表すキーワードを変換できる()
                Assert.That(TestingDao.ConvSqlWildcardToDotnetWildcard("あ%いう%え"), [Is].EqualTo("あ*いう*え"))
            End Sub

            <Test()> Public Sub エスケープ処理されたSQLのワイルドカードは_そのまま出力する()
                Assert.That(TestingDao.ConvSqlWildcardToDotnetWildcard("A?_B?%C"), [Is].EqualTo("A_B%C"))
            End Sub

            <Test()> Public Sub dotNetのワイルドカードは_エスケープ処理する()
                Assert.That(TestingDao.ConvSqlWildcardToDotnetWildcard("あ?い*う#え"), [Is].EqualTo("あ[?]い[*]う[#]え"))
            End Sub
        End Class

        Public Class InsertByTest : Inherits FakeTableDaoTest

            <SetUp()> Public Overrides Sub SetUp()
                MyBase.SetUp()
                dao.MakePrimaryUniqueKey = Nothing
            End Sub

            <Test()> Public Sub フラグを指定すれば_プライマリキーの設定が無くてもInsertByできる()
                dao.InsertBy(New SampleAdvVo With {.HogeId = 1}, ignoresPrimaryKey:=True)
                Assert.Pass()
            End Sub

            <Test()> Public Sub フラグを指定しなかったら_プライマリキーの設定がないと例外になる()
                Try
                    dao.InsertBy(New SampleAdvVo With {.HogeId = 1})
                    Assert.Fail()
                Catch ex As NullReferenceException
                    Assert.That(ex.Message, [Is].EqualTo("#MakePrimaryUniqueKey が必要"))
                End Try
            End Sub

        End Class

        Public Class FindForXxxxTest_SelectionCallbackで取得 : Inherits FakeTableDaoTest

            Private Overloads Function ToString(ParamArray values As SampleAdvVo()) As String
                Dim maker As New DebugStringMaker(Of SampleAdvVo)(
                    Function(defineBy As IDebugStringRuleBinder, vo As SampleAdvVo)
                        Return defineBy.Bind(vo.HogeId, vo.HogeName, vo.HogeDate, vo.HogeDecimal, vo.HogeEnum)
                    End Function)
                Return maker.MakeString(values)
            End Function

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                dao.InitializeTable(New SampleAdvVo With {.HogeId = 1, .HogeName = "one", .HogeDate = CDate("2012/03/04"), .HogeDecimal = 11.11D},
                                    New SampleAdvVo With {.HogeId = 2, .HogeName = "two", .HogeDate = CDate("05:06:07"), .HogeDecimal = 222D},
                                    New SampleAdvVo With {.HogeId = 3, .HogeName = "three", .HogeDate = CDate("2012/08/09 10:11:12")})
            End Sub

            <Test> Public Sub 指定した項目の値だけにして_1件取得できる()
                Dim actual As SampleAdvVo = dao.FindForObject(New SampleAdvVo With {.HogeId = 2},
                                                              selectionCallback:=Function(selection, vo) selection.Is(vo.HogeId, vo.HogeName))

                Assert.That(ToString(actual), [Is].EqualTo(
                            "HogeId HogeName HogeDate HogeDecimal HogeEnum" & vbCrLf &
                            "     2 'two'    null     null        null    "))
            End Sub

            <Test> Public Sub 指定した項目の値だけにして_1件取得できる_criteria()
                Dim actual As SampleAdvVo = dao.FindForObject(Function(criteria, vo) criteria.Like(vo.HogeName, "o%"),
                                                              selectionCallback:=Function(selection, vo) selection.Is(vo.HogeId, vo.HogeName))

                Assert.That(ToString(actual), [Is].EqualTo(
                            "HogeId HogeName HogeDate HogeDecimal HogeEnum" & vbCrLf &
                            "     1 'one'    null     null        null    "))
            End Sub

            <Test> Public Sub 指定した項目の値だけにして_n件取得できる()
                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(
                    Function(vo) vo.HogeId.HasValue AndAlso (vo.HogeId.Value = 1 OrElse vo.HogeId.Value = 2),
                    selectionCallback:=Function(selection, vo) selection.Is(vo.HogeId, vo.HogeDate))

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                  "HogeId HogeName HogeDate             HogeDecimal HogeEnum" & vbCrLf &
                  "     1 null     '2012/03/04 0:00:00' null        null    " & vbCrLf &
                  "     2 null     '5:06:07'            null        null    "))

            End Sub

            <Test> Public Sub 指定した項目の値だけにして_n件取得できる_criteria()
                Dim actuals As List(Of SampleAdvVo) = dao.FindForList(Function(criteria, vo) criteria.Like(vo.HogeName, "t%"),
                    selectionCallback:=Function(selection, vo) selection.Is(vo.HogeId, vo.HogeDate))

                Assert.That(ToString(actuals.ToArray()), [Is].EqualTo(
                  "HogeId HogeName HogeDate              HogeDecimal HogeEnum" & vbCrLf &
                  "     2 null     '5:06:07'             null        null    " & vbCrLf &
                  "     3 null     '2012/08/09 10:11:12' null        null    "))

            End Sub

        End Class

    End Class
End Namespace