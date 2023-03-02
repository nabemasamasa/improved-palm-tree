Imports System.Collections.Generic
Imports NUnit.Framework
Imports System.Data.SqlServerCe
Imports System.Data.SQLite
Imports Fhi.Fw.TestUtil.DebugString

Namespace Db.Impl
    Public MustInherit Class DaoEachTableImplTest

        Public Class NorequireDbTest : Inherits DaoEachTableImplTest

            <Test()> Public Sub THogeFugaDaoのPKは_HogeIdとHogeSubを設定している()
                Dim actual As THogeFugaVo = (New THogeFugaDao).MakePkVo(1, "2")
                Assert.That(actual.HogeId, [Is].EqualTo(1))
                Assert.That(actual.HogeSub, [Is].EqualTo("2"))
            End Sub

            <TestCase(Nothing, "a")>
            <TestCase("1", Nothing)>
            Public Sub NullをPKに含めてのDeleteByPkはエラーになる(hogeId As Object, hogeSub As String)
                Dim id As Integer? = Nothing
                If IsNumeric(hogeId) Then
                    id = CInt(hogeId)
                End If
                Try
                    Call (New THogeFugaDao).DeleteByPk(id, hogeSub)
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("PKにnullが含まれており危険です."))
                End Try
            End Sub

            Private Overloads Function ToString(records As IEnumerable(Of THogeFugaVo)) As String
                Return (New DebugStringMaker(Of THogeFugaVo)(Function(define, vo) define.Bind(vo.HogeId, vo.HogeSub, vo.HogeName, vo.HogeDate))).MakeString(records)
            End Function

            <Test()>
            Public Sub PkのみのVoに抽出できる()
                Dim param As THogeFugaVo
                param = New THogeFugaVo With {.HogeId = 1, .HogeName = "nm", .HogeSub = "1", .HogeDate = DateUtil.ConvDateValueToDateTime("2020/12/1")}

                Dim actual As THogeFugaVo = (New THogeFugaDao).ExtractHasOnlyPkVo(param)
                Assert.That(ToString({param, actual}), [Is].EqualTo(
                    "HogeId HogeSub HogeName HogeDate            " & vbCrLf &
                    "     1 '1'     'nm'     '2020/12/01 0:00:00'" & vbCrLf &
                    "     1 '1'     null     null                "))
            End Sub

            Private Class SimpleDao : Inherits DaoEachTableImpl(Of SimpleVo)
                Protected Overrides Function NewDbClient(dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient
                    Return New EmptyDbClient(dbFieldNameByPropertyName)
                End Function
                Protected Overrides Sub SettingPkField(table As PkTable(Of SimpleVo))
                    Throw New NotImplementedException
                End Sub
            End Class
            Private Class IrregularColumnNameDao : Inherits SimpleDao
                Protected Overrides Sub SettingFieldNameCamelizeIrregular(bind As BindTable)
                    MyBase.SettingFieldNameCamelizeIrregular(bind)
                    Dim vo As New SimpleVo
                    bind.IsA(vo).Bind(vo.Value, "MUSA")
                End Sub
            End Class
            Private Class SimpleVo
                Public Property Id As Integer
                Public Property SubId As Integer
                Public Property Name As String
                Public Property Value As String
            End Class
            Private Class EmptyDbClient : Inherits DbClient
                Public Sub New(dbFieldNameByPropertyName As Dictionary(Of String, String))
                    MyBase.New(DbProvider.SqlServer, "empty", dbFieldNameByPropertyName)
                End Sub
                Protected Overrides Sub SettingFieldNameCamelizeIrregular(bind As BindTable)
                    Dim vo As New SimpleVo
                    bind.IsA(vo).Bind(vo.Name, "NAME_9")
                End Sub
            End Class

            <Test()>
            Public Sub 設定されているVoを元に列名を特定することができる()
                Dim columns As String() = (New SimpleDao).DetectColumnNamesByVoProperties()
                Assert.That(columns, [Is].EquivalentTo({"ID", "SUB_ID", "NAME_9", "VALUE"}))
            End Sub

            <Test()>
            Public Sub 例外ルールが設定済みなら_例外ルールに則って列名を特定できる()
                Dim columns As String() = (New IrregularColumnNameDao).DetectColumnNamesByVoProperties()
                Assert.That(columns, [Is].EquivalentTo({"ID", "SUB_ID", "NAME_9", "MUSA"}))
            End Sub

            <Test()> Public Sub ConvertToDbColumnNames_指定項目のテーブル列名を取得できる_Dao_DbClientの個別適用名も反映する()
                Dim actuals As String() = (New IrregularColumnNameDao).ConvertToDbColumnNames(Function(selection, vo) selection.Is(vo.Id, vo.SubId, vo.Name, vo.Value))
                Assert.That(actuals, [Is].EqualTo({"ID", "SUB_ID", "NAME_9", "MUSA"}))
            End Sub

            <Test()> Public Sub ConvertToDbColumnNames_voの全項目をテーブル列名にして取得できる_Dao_DbClientの個別適用名も反映する()
                Dim actuals As String() = (New IrregularColumnNameDao).ConvertToDbColumnNames(Function(selection, vo) selection.Is(vo))
                Assert.That(actuals, [Is].EqualTo({"ID", "SUB_ID", "NAME_9", "MUSA"}))
            End Sub

            <Test()>
            Public Sub MatchVoPropsBetweenDbColumns_DBの列名と一致する場合_True()
                Dim dao As New SimpleDao
                dao.InternalFetchTableColumnsFromDb = Function() {"ID", "SUB_ID", "NAME_9", "VALUE"}
                Assert.IsTrue(dao.MatchesVoPropertiesBetweenDbColumns)
            End Sub

            <Test()>
            Public Sub MatchVoPropsBetweenDbColumns_取得したDB列名の順番に関わらず_判定できる()
                Dim dao As New SimpleDao
                dao.InternalFetchTableColumnsFromDb = Function() {"VALUE", "NAME_9", "SUB_ID", "ID"}
                Assert.IsTrue(dao.MatchesVoPropertiesBetweenDbColumns)
            End Sub

            <Test()>
            Public Sub MatchVoPropsBetweenDbColumns_DBの列名と不一致の場合_False()
                Dim dao As New SimpleDao
                dao.InternalFetchTableColumnsFromDb = Function() {"ID", "SUB_ID", "NAME_9", "FUGA"}
                Assert.IsFalse(dao.MatchesVoPropertiesBetweenDbColumns)
            End Sub

            <Test()>
            Public Sub MatchVoPropsBetweenDbColumns_DBの列名とで過不足がある場合_False(<Values({"ID", "SUB_ID", "NAME_9", "VALUE", "FUGA"},
                                                                                  {"ID", "SUB_ID", "NAME_9"})> dbColumns As String())
                Dim dao As New SimpleDao
                dao.InternalFetchTableColumnsFromDb = Function() dbColumns
                Assert.IsFalse(dao.MatchesVoPropertiesBetweenDbColumns)
            End Sub

        End Class

    End Class
    Public MustInherit Class THogeFugaDaoTest : Inherits DaoEachTableImplTest

        Protected MustOverride Function NewDao() As THogeFugaDao
        Protected MustOverride Function NewTFunyaDao() As TFunyaDao

        Protected value As THogeFugaVo
        Protected dao As THogeFugaDao

        <SetUp()> Public Sub Setup()
            dao = NewDao()
            dao.CreateTable()
        End Sub

        <TearDown()> Public Overridable Sub TearDown()
            dao.DropTable()
            dao = Nothing
        End Sub

        <Test()> Public Sub 最初は0件のはず()
            Dim dao As THogeFugaDao = NewDao()
            Assert.AreEqual(0, dao.CountBy(Nothing))
        End Sub

        <Test()> Public Overridable Sub InsertしてFindBy()

            Dim value As New THogeFugaVo
            value.HogeId = 12345
            value.HogeSub = "A"
            value.HogeName = "HogehogeName"
            value.HogeDecimal = 3.14D ' DB (8,2)
            value.HogeDate = CDate("2010/07/07 10:20:30")
            value.IsHoge = True
            value.UpdatedUserId = "U1234"
            value.UpdatedDate = "2010-02-02"
            value.UpdatedTime = "11:22:33"
            dao.InsertBy(value)

            Dim actuals As List(Of THogeFugaVo) = dao.FindAll()

            Assert.AreEqual(1, actuals.Count)

            Assert.AreEqual(12345, actuals(0).HogeId)
            Assert.AreEqual("A", actuals(0).HogeSub)
            Assert.AreEqual("HogehogeName", actuals(0).HogeName)
            Assert.AreEqual(3.14D, actuals(0).HogeDecimal)
            Assert.AreEqual(CDate("2010/07/07 10:20:30"), actuals(0).HogeDate)
            Assert.IsTrue(actuals(0).IsHoge.HasValue AndAlso actuals(0).IsHoge.Value)
            Assert.AreEqual("U1234", actuals(0).UpdatedUserId)
            Assert.AreEqual("2010-02-02", actuals(0).UpdatedDate)
            Assert.AreEqual("11:22:33", actuals(0).UpdatedTime)
        End Sub

        <Test()> Public Overridable Sub InsertしてFindBy_n件()

            dao.InsertBy(New THogeFugaVo With {.HogeId = 123, .HogeSub = "a", .HogeName = "n1", .HogeDecimal = 1.23D, .HogeDate = CDate("2010/07/07 10:20:30"), .IsHoge = True},
                         New THogeFugaVo With {.HogeId = 234, .HogeSub = "b", .HogeName = "n2", .HogeDecimal = 2.34D, .HogeDate = CDate("2010/08/08 11:21:31"), .IsHoge = False},
                         New THogeFugaVo With {.HogeId = 345, .HogeSub = "c", .HogeName = "n3", .HogeDecimal = 3.45D, .HogeDate = CDate("2010/09/09 12:22:32"), .IsHoge = True})

            Dim actuals As List(Of THogeFugaVo) = dao.FindAll()

            Assert.That(actuals.Select(Function(a) a.HogeId).ToArray, [Is].EqualTo({123, 234, 345}))
            Assert.That(actuals.Count, [Is].EqualTo(3))
        End Sub

        <Test()> Public Overridable Sub InsertしてFindByできる_Enum値も取得できる()

            Dim value As New THogeFugaVo
            value.HogeId = 1
            value.HogeSub = "b"
            value.HogeEnum = THogeFugaVo.Piyo.TWO
            dao.InsertBy(value)

            Dim actuals As List(Of THogeFugaVo) = dao.FindAll()

            Assert.AreEqual(THogeFugaVo.Piyo.TWO, actuals(0).HogeEnum)
        End Sub

        <Test()> Public Overridable Sub InsertしてFindByできる_Enum値Nothing()

            Dim value As New THogeFugaVo
            value.HogeId = 1
            value.HogeSub = "b"
            value.HogeEnum = Nothing
            dao.InsertBy(value)

            Dim actuals As List(Of THogeFugaVo) = dao.FindAll()

            Assert.AreEqual(Nothing, actuals(0).HogeEnum)
        End Sub

        <Test()> Public Overridable Sub InsertしてFindByできる_Enum不正値でも取得できる()

            Const INVALID_ENUM_VALUE As Integer = 999
            Dim value As New THogeFugaVo
            value.HogeId = 1
            value.HogeSub = "b"
            value.HogeEnum = DirectCast([Enum].ToObject(GetType(THogeFugaVo.Piyo), INVALID_ENUM_VALUE), THogeFugaVo.Piyo)
            dao.InsertBy(value)

            Dim actuals As List(Of THogeFugaVo) = dao.FindAll()

            Assert.AreEqual([Enum].ToObject(GetType(THogeFugaVo.Piyo), INVALID_ENUM_VALUE), actuals(0).HogeEnum)
        End Sub

        <Test()> Public Sub Transaction無しでinsert()

            Assert.AreEqual(0, dao.CountBy(Nothing))

            dao.InsertBy(NewSampleVo("A"))

            Assert.AreEqual(1, dao.CountBy(Nothing))

            dao.InsertBy(NewSampleVo("1"))

            Assert.AreEqual(2, dao.CountBy(Nothing))
        End Sub

        <Test()> Public Sub Transaction有でCommit後は_常時commit()

            Using db As DbClient = dao.InternalNewClient
                db.BeginTransaction()

                Assert.AreEqual(0, dao.CountBy(Nothing))
                dao.InsertBy(NewSampleVo("X"))
                Assert.AreEqual(1, dao.CountBy(Nothing))

                db.Commit()

                dao.InsertBy(NewSampleVo("Y"))

                Assert.AreEqual(2, dao.CountBy(Nothing))
            End Using

            '' "Y"のRollbackはかからない
            Assert.AreEqual(2, dao.CountBy(Nothing))
        End Sub

        <Test()> Public Sub Transaction有で別々のDaoをRollbackできる()
            Dim funyaDao As TFunyaDao = NewTFunyaDao()
            funyaDao.CreateTable()
            Try
                Using db As DbClient = dao.InternalNewClient
                    db.BeginTransaction()

                    Assert.AreEqual(0, dao.CountBy(Nothing))
                    Assert.AreEqual(0, funyaDao.CountBy(Nothing))

                    dao.InsertBy(NewSampleVo("X"))
                    funyaDao.InsertBy(NewFunyaVo(123&, "funyaFunya"))
                    funyaDao.InsertBy(NewFunyaVo(456&, "fugaFuga"))

                    Assert.AreEqual(1, dao.CountBy(Nothing))
                    Assert.AreEqual(2, funyaDao.CountBy(Nothing))

                    db.Rollback()

                    Assert.AreEqual(0, dao.CountBy(Nothing))
                    Assert.AreEqual(0, funyaDao.CountBy(Nothing))
                End Using
            Finally
                funyaDao.DropTable()
            End Try
        End Sub

        <Test()> Public Sub FindBy_Criteriaを使って検索()

            dao.InsertBy(NewSampleVo("E"))
            dao.InsertBy(NewSampleVo("F"))
            dao.InsertBy(NewSampleVo("G"))
            dao.InsertBy(NewSampleVo("H"))

            Dim vo As New THogeFugaVo
            Dim criteria As New Criteria(Of THogeFugaVo)(vo)
            criteria.GreaterEqual(vo.HogeSub, "F").LessThan(vo.HogeSub, "H")

            Dim actuals As List(Of THogeFugaVo) = dao.FindBy(criteria)
            Assert.That(actuals.Count(), [Is].EqualTo(2))
            Assert.That(actuals(0).HogeSub, [Is].EqualTo("F"))
            Assert.That(actuals(1).HogeSub, [Is].EqualTo("G"))
        End Sub

        <Test()> Public Sub FindBy_Criteriaを作るラムダ式を使って検索()

            dao.InsertBy(NewSampleVo("E"))
            dao.InsertBy(NewSampleVo("F"))
            dao.InsertBy(NewSampleVo("G"))
            dao.InsertBy(NewSampleVo("H"))

            Dim actuals As List(Of THogeFugaVo) = dao.FindBy(Function(criteria, vo) criteria.GreaterEqual(vo.HogeSub, "F").LessThan(vo.HogeSub, "H"))

            Assert.That(actuals.Count(), [Is].EqualTo(2))
            Assert.That(actuals(0).HogeSub, [Is].EqualTo("F"))
            Assert.That(actuals(1).HogeSub, [Is].EqualTo("G"))
        End Sub

        <Test()> Public Sub FindByPk_普通に検索()

            dao.InsertBy(NewSampleVo("E"))
            dao.InsertBy(NewSampleVo("F"))
            dao.InsertBy(NewSampleVo("G"))

            Dim actual As THogeFugaVo = dao.FindByPk(12345, "F")
            Assert.AreEqual(12345, actual.HogeId)
            Assert.AreEqual("F", actual.HogeSub)
        End Sub

        <Test()> Public Sub FindByPk_検索条件にNullがあったら該当しない()

            dao.InsertBy(NewSampleVo("E"))
            dao.InsertBy(NewSampleVo("F"))
            dao.InsertBy(NewSampleVo("G"))

            Dim actual As THogeFugaVo = dao.FindByPk(12345, Nothing)
            Assert.IsNull(actual)
        End Sub

        <Test()> Public Sub UpdateByPk_通常の更新処理()

            dao.InsertBy(NewSampleVo("L"))
            dao.InsertBy(NewSampleVo("M"))
            dao.InsertBy(NewSampleVo("N"))

            Dim actual1 As THogeFugaVo = dao.FindByPk(12345, "M")
            Dim actual2 As THogeFugaVo = dao.FindByPk(12345, "N")
            Assert.IsNull(actual1.HogeName)
            Assert.IsNull(actual2.HogeName)

            Dim param As THogeFugaVo = NewSampleVo("M")
            param.HogeName = "hogehogeName"
            Dim actual3 As Integer = dao.UpdateByPk(param)
            Assert.AreEqual(1, actual3)

            Dim actual4 As THogeFugaVo = dao.FindByPk(12345, "M")
            Dim actual5 As THogeFugaVo = dao.FindByPk(12345, "N")
            Assert.AreEqual("hogehogeName", actual4.HogeName)
            Assert.IsNull(actual5.HogeName) '' 違うレコードは更新していないこと
        End Sub
        <Test()> Public Sub DeleteBy_IDとNameで削除()
            dao.InsertBy(NewSampleVo("O"))
            dao.InsertBy(NewSampleVo("P"))
            Dim actual As Integer = dao.DeleteBy(NewConditionVo(12345, "O", Nothing))
            Assert.AreEqual(1, actual)
        End Sub

        <Test()> Public Sub DeleteBy_IDだけで削除()
            dao.InsertBy(NewSampleVo("Q"))
            dao.InsertBy(NewSampleVo("R"))
            Dim actual As Integer = dao.DeleteBy(NewConditionVo(12345, Nothing, Nothing))
            Assert.AreEqual(2, actual)
        End Sub

        <Test()> Public Sub DeleteBy_IDだけで削除して_他レコードが消えていない事()
            dao.InsertBy(NewSampleVo("A"))
            dao.InsertBy(NewSampleVo("S"))
            Dim vo As THogeFugaVo = NewSampleVo("D")
            vo.HogeId = 987
            vo.HogeName = "Poo"
            dao.InsertBy(vo)

            Dim actual As Integer = dao.DeleteBy(NewConditionVo(12345, Nothing, Nothing))
            Assert.AreEqual(2, actual)

            Dim actuals As List(Of THogeFugaVo) = dao.FindBy(NewConditionVo(987, Nothing, Nothing))
            Assert.AreEqual(1, actuals.Count)
            Assert.AreEqual("Poo", actuals(0).HogeName)
        End Sub

        <Test()> Public Sub SetForUpdate_普通にデータを検索()
            dao.InsertBy(NewSampleVo("E"))

            dao.SetForUpdate(True)
            Using db As DbClient = dao.InternalNewClient
                db.BeginTransaction()
                Dim actual As THogeFugaVo = dao.FindByPk(12345, "E")
                Assert.IsNotNull(actual)
            End Using
        End Sub

        <Test()> Public Sub SetForUpdate_トランザクションを開始していない場合は例外()
            dao.SetForUpdate(True)
            Try
                dao.FindByPk(123, "sub")
                Assert.Fail("トランザクション開始していない時のSelectForUpdateで例外になっていない")
            Catch expected As InvalidOperationException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub FetchTableColumnsFromDb_実DBから列名を取得できる()
            Dim actual As String() = dao.FetchTableColumnsFromDb()
            Dim expect As String() = VoUtil.GetPropertyNames(Of THogeFugaVo).Select(Function(name) StringUtil.Decamelize(name)).ToArray
            Assert.That(actual, [Is].EquivalentTo(expect))
        End Sub

        Protected Function NewSampleVo(ByVal hogeSub As String) As THogeFugaVo
            Dim value As New THogeFugaVo
            value.HogeId = 12345
            value.HogeSub = hogeSub
            value.UpdatedUserId = "111"
            value.UpdatedDate = "2010-01-01"
            value.UpdatedTime = "10:20:30"
            Return value
        End Function
        Protected Function NewFunyaVo(ByVal FunyaId As Long, ByVal FunyaName As String) As TFunyaVo
            Dim value As New TFunyaVo
            value.FunyaId = FunyaId
            value.FunyaName = FunyaName
            value.UpdatedUserId = "111"
            value.UpdatedDate = "2010-01-01"
            value.UpdatedTime = "10:20:30"
            Return value
        End Function
        Protected Function NewConditionVo(ByVal hogeId As Nullable(Of Int32), ByVal hogeSub As String, ByVal hogeName As String) As THogeFugaVo
            Dim value As New THogeFugaVo
            If hogeId IsNot Nothing Then
                value.HogeId = hogeId
            End If
            If hogeSub IsNot Nothing Then
                value.HogeSub = hogeSub
            End If
            If hogeName IsNot Nothing Then
                value.HogeName = hogeName
            End If
            Return value
        End Function
    End Class

    <Category("RequireDB")> Public Class THogeFugaDaoSqlServerTest : Inherits THogeFugaDaoTest

        Protected Overrides Function NewDao() As THogeFugaDao
            Return New THogeFugaDao()
        End Function

        Protected Overrides Function NewTFunyaDao() As TFunyaDao
            Return New TFunyaDao()
        End Function

        <Test()> Public Overridable Sub 小数桁overをinsertすると_切り捨てられてinsertできる()

            dao.InsertBy(New THogeFugaVo With {.HogeId = 2, .HogeSub = "3", .HogeDecimal = 3.14159D})

            Dim actuals As List(Of THogeFugaVo) = dao.FindAll()

            Assert.That(actuals(0).HogeDecimal, [Is].EqualTo(3.14D))
            Assert.That(actuals.Count, [Is].EqualTo(1))
        End Sub

        <Test()> Public Overridable Sub Transaction有のネスト_ReadCommittedSnapshotがon_Commit前に別Txから参照してもBeginTrans時点での確定データを参照できる()
            Using db As DbClient = dao.InternalNewClient
                db.BeginTransaction()

                Assert.AreEqual(0, dao.CountBy(Nothing))

                dao.InsertBy(NewSampleVo("X"))

                Using db2 As DbClient = dao.InternalNewClient
                    db2.LockTimeout = 10
                    db2.BeginTransaction()
                    Assert.AreEqual(0, dao.CountBy(Nothing), "最初のTxでInsert中だがブロックされずに件数を取得できる")
                End Using

                db.Commit()
                db.BeginTransaction()

                dao.UpdateByPk(NewSampleVo("Y"))

                Using db2 As DbClient = dao.InternalNewClient
                    db2.LockTimeout = 10
                    db2.BeginTransaction()
                    Assert.AreEqual(1, dao.CountBy(Nothing), "最初のTxでUpdate中だがブロックされずに件数を取得できる")
                End Using
            End Using
        End Sub

        <Test()> Public Overridable Sub Transaction有のネスト_Commit後のネストからは参照問題無し()
            ' Suspend/Resumeの確認

            Using db As DbClient = dao.InternalNewClient
                db.BeginTransaction()

                Assert.AreEqual(0, dao.CountBy(Nothing))

                dao.InsertBy(NewSampleVo("X"))

                Assert.AreEqual(1, dao.CountBy(Nothing))

                db.Commit()
                db.BeginTransaction()

                Using db2 As DbClient = dao.InternalNewClient
                    db2.BeginTransaction()
                    Assert.AreEqual(1, dao.CountBy(Nothing))
                    dao.InsertBy(NewSampleVo("Y"))
                    Assert.AreEqual(2, dao.CountBy(Nothing))
                End Using

                Assert.AreEqual(1, dao.CountBy(Nothing))

                dao.InsertBy(NewSampleVo("Z"))

                Assert.AreEqual(2, dao.CountBy(Nothing))
            End Using

            Assert.AreEqual(1, dao.CountBy(Nothing))
        End Sub

    End Class

    <Category("RequireDB")> Public Class UpdateByPkExceptionTest : Inherits DaoEachTableImplTest

        Private Class THogeFugaNonPkDao : Inherits THogeFugaDao
            Protected Overrides Sub SettingPkField(ByVal table As PkTable(Of THogeFugaVo))
                'nop
            End Sub
        End Class

        Private dao As THogeFugaDao

        <SetUp()> Public Sub Setup()
            dao = New THogeFugaNonPkDao()
            dao.CreateTable()
        End Sub

        <TearDown()> Public Overridable Sub TearDown()
            dao.DropTable()
            dao = Nothing
        End Sub

        <Test> Public Sub UpdateByPk_PrimaryKeyが無ければ_例外()
            Try
                dao.UpdateByPk(New THogeFugaVo With {.HogeId = 1, .HogeName = "1"})
                Assert.Fail("例外にならない＝全更新してしまう")
            Catch expected As InvalidOperationException
                Assert.That(expected.Message, [Is].EqualTo("PrimaryKeyがありません。"))
            End Try
        End Sub

        <Test> Public Sub UpdateIgnoreNullByPk_PrimaryKeyが無ければ_例外()
            Try
                dao.UpdateIgnoreNullByPk(New THogeFugaVo With {.HogeId = 1, .HogeName = "1"})
                Assert.Fail("例外にならない＝全更新してしまう")
            Catch expected As InvalidOperationException
                Assert.That(expected.Message, [Is].EqualTo("PrimaryKeyがありません。"))
            End Try
        End Sub
    End Class

    <Category("RequireDB")> Public Class THogeFugaDaoSqlServerCeTest : Inherits THogeFugaDaoTest

        Public Class THogeFugaBehaviorImpl : Implements THogeFugaDao.IBehavior

            Public Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient Implements THogeFugaDao.IBehavior.NewDbClient
                Return New TestSqlCeDbClient(dbFieldNameByPropertyName)
            End Function

            Public Sub CreateTableTHogeFuga() Implements THogeFugaDao.IBehavior.CreateTableTHogeFuga
                Dim sql As String = "CREATE TABLE T_HOGE_FUGA (" _
                                    & "HOGE_ID INT NOT NULL, " _
                                    & "HOGE_SUB nvarchar(1) NOT NULL, " _
                                    & "HOGE_NAME nvarchar(122) NULL, " _
                                    & "HOGE_DATE Datetime NULL, " _
                                    & "HOGE_DECIMAL NUMERIC(8,2) NULL, " _
                                    & "IS_HOGE BIT NULL, " _
                                    & "HOGE_ENUM INT NULL, " _
                                    & "UPDATED_USER_ID nvarchar(10) NULL, " _
                                    & "UPDATED_DATE nvarchar(10) NULL, " _
                                    & "UPDATED_TIME nvarchar(8) NULL, " _
                                    & "CONSTRAINT [PK_T_HOGE_FUGA] PRIMARY KEY (" _
                                    & "HOGE_ID , HOGE_SUB " _
                                    & ") " _
                                    & ") "
                NewDbClient(Nothing).Update(sql)
            End Sub

            Public Sub DropTableTHogeFuga() Implements THogeFugaDao.IBehavior.DropTableTHogeFuga
                NewDbClient(Nothing).Update("DROP TABLE T_HOGE_FUGA")
            End Sub
        End Class

        Public Class TFunyaBehaviorImpl : Implements TFunyaDao.IBehavior

            Public Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient Implements TFunyaDao.IBehavior.NewDbClient
                Return New TestSqlCeDbClient(dbFieldNameByPropertyName)
            End Function

            Public Sub CreateTableTHogeFuga() Implements TFunyaDao.IBehavior.CreateTableTHogeFuga
                Dim sql As String = "CREATE TABLE T_FUNYA (" _
                                    & "FUNYA_ID NUMERIC(12,0) NOT NULL, " _
                                    & "FUNYA_NAME nvarchar(122) NULL, " _
                                    & "UPDATED_USER_ID nvarchar(10) NOT NULL, " _
                                    & "UPDATED_DATE nvarchar(10) NOT NULL, " _
                                    & "UPDATED_TIME nvarchar(8) NOT NULL, " _
                                    & "CONSTRAINT [PK_T_FUNYA] PRIMARY KEY (" _
                                    & "FUNYA_ID " _
                                    & ") " _
                                    & ") "
                NewDbClient(Nothing).Update(sql)
            End Sub

            Public Sub DropTableTHogeFuga() Implements TFunyaDao.IBehavior.DropTableTHogeFuga
                NewDbClient(Nothing).Update("DROP TABLE T_FUNYA")
            End Sub
        End Class

        <TestFixtureSetUp()> Public Sub SetUpOnce()
            DbTestInitializer.InitializeSqlServerCe()
        End Sub

        Protected Overrides Function NewDao() As THogeFugaDao
            Return New THogeFugaDao(New THogeFugaBehaviorImpl)
        End Function

        Protected Overrides Function NewTFunyaDao() As TFunyaDao
            Return New TFunyaDao(New TFunyaBehaviorImpl)
        End Function

        <Test()> Public Overrides Sub InsertしてFindBy_n件()
            Try
                MyBase.InsertしてFindBy_n件()
                Assert.Fail()
            Catch expected As InvalidOperationException
                ' （未調査だけどたぶん）SQLServerCEは複数行一括insertできない
                Assert.That(expected.Message, [Is].EqualTo("複数行一括insertできるDBではない"))
            End Try
        End Sub

        <Test()> Public Overridable Sub 小数桁overをinsertすると_切り捨てられてinsertできる()

            dao.InsertBy(New THogeFugaVo With {.HogeId = 2, .HogeSub = "3", .HogeDecimal = 3.14159D})

            Dim actuals As List(Of THogeFugaVo) = dao.FindAll()

            Assert.That(actuals(0).HogeDecimal, [Is].EqualTo(3.14D))
            Assert.That(actuals.Count, [Is].EqualTo(1))
        End Sub

        <Test()> Public Overridable Sub Transaction有のネスト_CEは入れ子トランザクションをサポートしない_がCommit前に別Txから参照はできる()
            Using db As DbClient = dao.InternalNewClient
                db.BeginTransaction()

                Assert.AreEqual(0, dao.CountBy(Nothing))

                dao.InsertBy(NewSampleVo("X"))

                Using db2 As DbClient = dao.InternalNewClient
                    db2.LockTimeout = 10
                    db2.BeginTransaction()
                    Assert.AreEqual(0, dao.CountBy(Nothing), "最初のTxでInsert中だがブロックされずに件数を取得できる")
                End Using

                db.Commit()
                db.BeginTransaction()

                dao.UpdateByPk(NewSampleVo("Y"))

                Using db2 As DbClient = dao.InternalNewClient
                    db2.LockTimeout = 10
                    db2.BeginTransaction()
                    Assert.AreEqual(1, dao.CountBy(Nothing), "最初のTxでUpdate中だがブロックされずに件数を取得できる")
                End Using
            End Using
        End Sub

        <Test()> Public Overridable Sub Transaction有のネスト_CEは入れ子トランザクションをサポートしない_Commit前に別Txでinsertしたらinsertできずにロックタイムアウトエラー()
            Using db As DbClient = dao.InternalNewClient
                db.BeginTransaction()

                Assert.AreEqual(0, dao.CountBy(Nothing))

                dao.InsertBy(NewSampleVo("X"))

                Using db2 As DbClient = dao.InternalNewClient
                    db2.LockTimeout = 10
                    db2.BeginTransaction()
                    Try
                        ' TODO 10msにロックタイムアウトを設定しているつもりだが、実際には2000msくらい。設定方法を要確認
                        dao.InsertBy(NewSampleVo("Y"))
                        Assert.Fail("ブロックされずにinsertできている。おかしい")
                    Catch expected As SqlCeLockTimeoutException
                        Assert.IsTrue(True, "最初のTxがブロックして、ロックタイムアウトエラーになる")
                    End Try
                End Using
            End Using
        End Sub
    End Class

    <Ignore("x64用DLLに差し替えないと動かない感じ ※Ignore期間 2021/3/31まで")>
    Public Class THogeFugaDaoSQLiteTest : Inherits THogeFugaDaoTest

        Public Class THogeFugaBehaviorImpl : Implements THogeFugaDao.IBehavior

            Public Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient Implements THogeFugaDao.IBehavior.NewDbClient
                Return New TestSQLiteDbClient(dbFieldNameByPropertyName)
            End Function

            Public Sub CreateTableTHogeFuga() Implements THogeFugaDao.IBehavior.CreateTableTHogeFuga
                Dim sql As String = "CREATE TABLE T_HOGE_FUGA (" _
                                    & "HOGE_ID INTEGER NOT NULL, " _
                                    & "HOGE_SUB TEST NOT NULL, " _
                                    & "HOGE_NAME TEXT NULL, " _
                                    & "HOGE_DATE TEXT NULL, " _
                                    & "HOGE_DECIMAL NUMERIC(8,2) NULL, " _
                                    & "IS_HOGE BIT NULL, " _
                                    & "HOGE_ENUM INT NULL, " _
                                    & "UPDATED_USER_ID TEXT NULL, " _
                                    & "UPDATED_DATE TEXT NULL, " _
                                    & "UPDATED_TIME TEXT NULL, " _
                                    & "CONSTRAINT [PK_T_HOGE_FUGA] PRIMARY KEY (" _
                                    & "HOGE_ID , HOGE_SUB " _
                                    & ") " _
                                    & ") "
                NewDbClient(Nothing).Update(sql)
            End Sub

            Public Sub DropTableTHogeFuga() Implements THogeFugaDao.IBehavior.DropTableTHogeFuga
                NewDbClient(Nothing).Update("DROP TABLE T_HOGE_FUGA")
            End Sub
        End Class

        Public Class TFunyaBehaviorImpl : Implements TFunyaDao.IBehavior

            Public Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient Implements TFunyaDao.IBehavior.NewDbClient
                Return New TestSQLiteDbClient(dbFieldNameByPropertyName)
            End Function

            Public Sub CreateTableTHogeFuga() Implements TFunyaDao.IBehavior.CreateTableTHogeFuga
                Dim sql As String = "CREATE TABLE T_FUNYA (" _
                                    & "FUNYA_ID NUMERIC(12,0) NOT NULL, " _
                                    & "FUNYA_NAME TEXT NULL, " _
                                    & "UPDATED_USER_ID TEXT NOT NULL, " _
                                    & "UPDATED_DATE TEXT NOT NULL, " _
                                    & "UPDATED_TIME TEXT NOT NULL, " _
                                    & "CONSTRAINT [PK_T_FUNYA] PRIMARY KEY (" _
                                    & "FUNYA_ID " _
                                    & ") " _
                                    & ") "
                NewDbClient(Nothing).Update(sql)
            End Sub

            Public Sub DropTableTHogeFuga() Implements TFunyaDao.IBehavior.DropTableTHogeFuga
                NewDbClient(Nothing).Update("DROP TABLE T_FUNYA")
            End Sub
        End Class

        Protected Overrides Function NewDao() As THogeFugaDao
            Return New THogeFugaDao(New THogeFugaBehaviorImpl)
        End Function

        Protected Overrides Function NewTFunyaDao() As TFunyaDao
            Return New TFunyaDao(New TFunyaBehaviorImpl)
        End Function

        Public Overrides Sub TearDown()
            MyBase.TearDown()
            ' 以下Aのテスト実行後、次に起動するテストでT_HOGE_FUGAテーブルが残っているエラーになる
            ' SQLite固有の問題っぽく、GCすると解決するので、SQLiteは常にGCする
            ' A) Transaction有のネスト_SQLiteは入れ子トランザクションをサポートしない_Commit前に別Txでinsertしたらinsertできずにロックタイムアウトエラー()
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Sub

        <Test()> Public Overrides Sub InsertしてFindBy_n件()
            Try
                MyBase.InsertしてFindBy_n件()
                Assert.Fail()
            Catch expected As InvalidOperationException
                ' （未対応だから）SQLiteは複数行一括insertできない
                Assert.That(expected.Message, [Is].EqualTo("複数行一括insertできるDBではない"))
            End Try
        End Sub

        <Test()> Public Overridable Sub 小数桁overをinsertしても_SQLiteはReal型だから_浮動小数点で保持してる_insertできる()

            dao.InsertBy(New THogeFugaVo With {.HogeId = 2, .HogeSub = "3", .HogeDecimal = 3.14159D})

            Dim actuals As List(Of THogeFugaVo) = dao.FindAll()

            Assert.That(actuals(0).HogeDecimal, [Is].EqualTo(3.14159D))
            Assert.That(actuals.Count, [Is].EqualTo(1))
        End Sub

        <Test()> Public Overridable Sub Transaction有のネスト_SQLiteは入れ子トランザクションをサポートしない_Commit前に別Txでinsertしたらinsertできずにロックタイムアウトエラー()
            Using db As DbClient = dao.InternalNewClient
                db.BeginTransaction()

                Assert.AreEqual(0, dao.CountBy(Nothing))

                dao.InsertBy(NewSampleVo("X"))

                Using db2 As DbClient = dao.InternalNewClient
                    db2.LockTimeout = 1000 ' 秒単位じゃないと設定できない
                    db2.BeginTransaction()
                    Try
                        dao.InsertBy(NewSampleVo("Y"))
                        Assert.Fail("ブロックされずにinsertできている。おかしい")
                    Catch ex As SQLiteException
                        log4net.LogManager.GetLogger(Me.GetType).Debug("↑このInsertエラーは「ロックタイムアウトするテスト」で意図したエラーだから問題なし")
                        Assert.IsTrue(True, "最初のTxがブロックして、ロックタイムアウトエラーになる")
                    End Try
                End Using
            End Using
        End Sub
    End Class
End Namespace