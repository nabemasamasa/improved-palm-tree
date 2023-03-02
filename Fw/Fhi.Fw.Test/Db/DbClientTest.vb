Imports NUnit.Framework
Imports System.Threading

Namespace Db
    <TestFixture()> <Category("RequireDB")> Public MustInherit Class DbClientTest

#Region "Testing classes..."
        Private Class TestingDbAccess : Inherits DbClient
            Public Sub New()
                Me.New(Nothing)
            End Sub
            Public Sub New(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String))
                MyBase.New(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG, dbFieldNameByPropertyName)
            End Sub
        End Class

        Private Class TestingDbAccessAnother : Inherits DbClient
            Public Sub New()
                MyBase.New(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG_ANOTHOER, Nothing)
            End Sub
        End Class
#End Region

        Public Class IsBeginningTransactionTest : Inherits DbClientTest

            <Test()> Public Sub IsBeginningTransaction_インスタンス生成直後はfalse()
                Dim db As New TestingDbAccess
                Assert.IsFalse(db.IsBeginningTransaction)
            End Sub

            <Test()> Public Sub IsBeginningTransaction_インスタンス生成直後はfalse_Using版()
                Using db As New TestingDbAccess
                    Assert.IsFalse(db.IsBeginningTransaction)
                End Using
            End Sub

            <Test()> Public Sub IsBeginningTransaction_トランザクション開始すればtrue()
                Using db As New TestingDbAccess
                    db.BeginTransaction()
                    Assert.IsTrue(db.IsBeginningTransaction)
                End Using
            End Sub

            <Test()> Public Sub IsBeginningTransaction_トランザクション開始すれば_別インスタンスのTestingDbAccessでもtrue()
                Using db As New TestingDbAccess
                    db.BeginTransaction()

                    Dim hogeDb As New TestingDbAccess
                    Assert.IsTrue(hogeDb.IsBeginningTransaction)
                End Using
            End Sub

            <Test()> Public Sub IsBeginningTransaction_トランザクション開始しても_別接続先のTestingDbAccessではfalse()
                Using db As New TestingDbAccess
                    db.BeginTransaction()

                    Dim hogeDb As New TestingDbAccessAnother
                    Assert.IsFalse(hogeDb.IsBeginningTransaction)
                End Using
            End Sub

            <Test()> Public Sub IsBeginningTransaction_トランザクション開始後rollbackすればfalse()
                Using db As New TestingDbAccess
                    db.BeginTransaction()
                    db.Rollback()
                    Assert.IsFalse(db.IsBeginningTransaction)
                End Using
            End Sub

        End Class

        Public Class IsWriteLockHeldTest : Inherits DbClientTest
            <Test()> Public Sub IsWriteLockHeld_インスタンス生成直後はfalse()
                Dim rwLock As New ReaderWriterLockSlim
                Assert.IsFalse(rwLock.IsWriteLockHeld)
            End Sub

            <Test()> Public Sub IsWriteLockHeld_インスタンス生成直後はfalse_Using版()
                Using rwLock As New ReaderWriterLockSlim
                    Assert.IsFalse(rwLock.IsWriteLockHeld)
                End Using
            End Sub

            <Test()> Public Sub IsWriteLockHeld_書き込みロック中はtrue()
                Using rwLock As New ReaderWriterLockSlim
                    rwLock.EnterWriteLock()
                    Assert.IsTrue(rwLock.IsWriteLockHeld)
                    rwLock.ExitWriteLock()
                End Using
            End Sub

            <Test()> Public Sub IsWriteLockHeld_書き込みロック解除後はfalse()
                Using rwLock As New ReaderWriterLockSlim
                    rwLock.EnterWriteLock()
                    rwLock.ExitWriteLock()
                    Assert.IsFalse(rwLock.IsWriteLockHeld)
                End Using
            End Sub
        End Class

        Public Class CamelizeIrregularTest : Inherits DbClientTest

            Private Class FooVo
                Private _abc2 As String
                Private _ghi4 As String
                Public Property Abc2() As String
                    Get
                        Return _abc2
                    End Get
                    Set(ByVal value As String)
                        _abc2 = value
                    End Set
                End Property
                Public Property Ghi4() As String
                    Get
                        Return _ghi4
                    End Get
                    Set(ByVal value As String)
                        _ghi4 = value
                    End Set
                End Property
            End Class
            Private Class FooSubVo : Inherits FooVo
                Private _def3 As String
                Public Property Def3() As String
                    Get
                        Return _def3
                    End Get
                    Set(ByVal value As String)
                        _def3 = value
                    End Set
                End Property
            End Class
            Private Class FooSubExVo : Inherits FooSubVo
            End Class
            Private Class BarVo
                Private _abcTwo As String
                Public Property AbcTwo() As String
                    Get
                        Return _abcTwo
                    End Get
                    Set(ByVal value As String)
                        _abcTwo = value
                    End Set
                End Property
            End Class
            Private Class BazVo
                Private _abc2 As String
                Public Property Abc2() As String
                    Get
                        Return _abc2
                    End Get
                    Set(ByVal value As String)
                        _abc2 = value
                    End Set
                End Property
            End Class
            Private Class TestingDbClient2 : Inherits DbClient
                Public Sub New()
                    MyBase.New(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG, Nothing)
                End Sub
                Protected Overrides Sub SettingFieldNameCamelizeIrregular(ByVal bind As BindTable)
                    MyBase.SettingFieldNameCamelizeIrregular(bind)
                    Dim foo As New FooVo
                    bind.IsA(foo).Bind(foo.Abc2, "ABC2")
                    Dim bar As New BarVo
                    bind.IsA(bar).Bind(bar.AbcTwo, "ABC2")
                End Sub
            End Class

            <Test()> Public Sub 異なるテーブルVoで同じIrregular項目名ABC2を設定しても_それぞれ取得できる()
                Using db As New TestingDbClient2
                    Dim actual As FooVo = db.QueryForObject(Of FooVo)("select 'Foo' ABC2")
                    Assert.That(actual.Abc2, [Is].EqualTo("Foo"))
                    Dim actual2 As BarVo = db.QueryForObject(Of BarVo)("select 'Bar' ABC2")
                    Assert.That(actual2.AbcTwo, [Is].EqualTo("Bar"))
                End Using
            End Sub

            <Test()> Public Sub 異なるテーブルVoで同じIrregular項目名ABC2を設定しても_それぞれ取得できる_けどBazVoは未設定だから取得できない()
                Using db As New TestingDbClient2
                    Dim actual As FooVo = db.QueryForObject(Of FooVo)("select 'Foo' ABC2")
                    Assert.That(actual.Abc2, [Is].EqualTo("Foo"))
                    Dim actual2 As BazVo = db.QueryForObject(Of BazVo)("select 'Baz' ABC2")
                    Assert.That(actual2.Abc2, [Is].Null, "取得できない")
                End Using
            End Sub

            <Test()> Public Sub コンストラクタ引数のIrregular項目定義は_型に依存しないので全適用される()
                Dim dbFieldNameByPropertyName As New Dictionary(Of String, String)
                dbFieldNameByPropertyName.Add("Abc2", "ABC_TWO")
                Using db As New TestingDbAccess(dbFieldNameByPropertyName)
                    Dim actual As FooVo = db.QueryForObject(Of FooVo)("select 'Foo' ABC_TWO")
                    Assert.That(actual.Abc2, [Is].EqualTo("Foo"))
                    Dim actual2 As BazVo = db.QueryForObject(Of BazVo)("select 'Baz' ABC_TWO")
                    Assert.That(actual2.Abc2, [Is].EqualTo("Baz"))
                End Using
            End Sub

            Public Class SettingFieldNameCamelizeIrregularTest : Inherits DbClientTest

                Private Class TestingDbClient_NoneMapping : Inherits DbClient
                    Public Sub New()
                        MyBase.New(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG, Nothing)
                    End Sub
                    Protected Overrides Sub SettingFieldNameCamelizeIrregular(ByVal bind As BindTable)
                        MyBase.SettingFieldNameCamelizeIrregular(bind)
                    End Sub
                End Class

                Private Class TestingDbClient_MappingSubClass : Inherits DbClient
                    Public Sub New()
                        MyBase.New(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG, Nothing)
                    End Sub
                    Protected Overrides Sub SettingFieldNameCamelizeIrregular(ByVal bind As BindTable)
                        MyBase.SettingFieldNameCamelizeIrregular(bind)
                        Dim fooSub As New FooSubVo
                        bind.IsA(fooSub).Bind(fooSub.Abc2, "ABC2") _
                                        .Bind(fooSub.Def3, "DEF3")
                    End Sub
                End Class

                Private Class TestingDbClient_MappingBoth : Inherits DbClient
                    Public Sub New()
                        MyBase.New(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG, Nothing)
                    End Sub
                    Protected Overrides Sub SettingFieldNameCamelizeIrregular(ByVal bind As BindTable)
                        MyBase.SettingFieldNameCamelizeIrregular(bind)
                        Dim foo As New FooVo
                        bind.IsA(foo).Bind(foo.Abc2, "ABC2")
                        Dim fooSub As New FooSubVo
                        bind.IsA(fooSub).Bind(fooSub.Abc2, "ABC2") _
                                        .Bind(fooSub.Def3, "DEF3")
                    End Sub
                End Class

                Private Class TestingDbClient_MappingBoth2 : Inherits DbClient
                    Public Sub New()
                        MyBase.New(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG, Nothing)
                    End Sub
                    Protected Overrides Sub SettingFieldNameCamelizeIrregular(ByVal bind As BindTable)
                        MyBase.SettingFieldNameCamelizeIrregular(bind)
                        Dim foo As New FooVo
                        bind.IsA(foo).Bind(foo.Abc2, "ABC2")
                        Dim fooSub As New FooSubVo
                        bind.IsA(fooSub).Bind(fooSub.Def3, "DEF3")
                    End Sub
                End Class

                <Test()> Public Sub マッピングしていなければ_取得できない()
                    Using db As New TestingDbClient_NoneMapping
                        Dim actual As FooVo = db.QueryForObject(Of FooVo)("select 'Foo' ABC2")
                        Assert.That(actual.Abc2, [Is].Null, "取得できない")
                        Dim actual2 As FooSubVo = db.QueryForObject(Of FooSubVo)("select 'FooSub' ABC2")
                        Assert.That(actual2.Abc2, [Is].Null, "取得できない")
                    End Using
                End Sub

                <Test()> Public Sub スーパークラスを手動マッピングしたら_サブクラスにも引き継がれる()
                    Using db As New TestingDbClient2
                        Dim actual As FooVo = db.QueryForObject(Of FooVo)("select 'Foo' ABC2")
                        Assert.That(actual.Abc2, [Is].EqualTo("Foo"))
                        Dim actual2 As FooSubVo = db.QueryForObject(Of FooSubVo)("select 'FooSub' ABC2")
                        Assert.That(actual2.Abc2, [Is].EqualTo("FooSub"))
                        Dim actual3 As FooSubExVo = db.QueryForObject(Of FooSubExVo)("select 'FooSubEx' ABC2")
                        Assert.That(actual3.Abc2, [Is].EqualTo("FooSubEx"))
                    End Using
                End Sub

                <Test()> Public Sub スーパークラスを手動マッピングしても_サブクラスの項目はマッピングしていないので取得できない()
                    Using db As New TestingDbClient2
                        Dim actual As FooSubVo = db.QueryForObject(Of FooSubVo)("select 'FooSub' ABC2")
                        Assert.That(actual.Abc2, [Is].EqualTo("FooSub"))
                        Dim actual2 As FooSubVo = db.QueryForObject(Of FooSubVo)("select 'FooSub' DEF3")
                        Assert.That(actual2.Def3, [Is].Null, "取得できない")
                    End Using
                End Sub

                <Test()> Public Sub サブクラスをマッピングしたら_サブクラスの項目を取得できる_がスーパークラスの項目は取得できない()
                    Using db As New TestingDbClient_MappingSubClass
                        Dim actual As FooVo = db.QueryForObject(Of FooVo)("select 'Foo' ABC2")
                        Assert.That(actual.Abc2, [Is].Null, "取得できない")
                        Dim actual2 As FooSubVo = db.QueryForObject(Of FooSubVo)("select 'FooSub' ABC2")
                        Assert.That(actual2.Abc2, [Is].EqualTo("FooSub"))
                        Dim actual3 As FooSubVo = db.QueryForObject(Of FooSubVo)("select 'FooSub' DEF3")
                        Assert.That(actual3.Def3, [Is].EqualTo("FooSub"))
                    End Using
                End Sub

                <Test()> Public Sub スーパークラスをマッピングし_サブクラスをマッピングしたら_スーパークラスの項目を取得でき_サブクラスの項目を取得できる()
                    Using db As New TestingDbClient_MappingBoth
                        Dim actual As FooVo = db.QueryForObject(Of FooVo)("select 'Foo' ABC2")
                        Assert.That(actual.Abc2, [Is].EqualTo("Foo"))
                        Dim actual2 As FooSubVo = db.QueryForObject(Of FooSubVo)("select 'FooSub' ABC2")
                        Assert.That(actual2.Abc2, [Is].EqualTo("FooSub"))
                        Dim actual3 As FooSubVo = db.QueryForObject(Of FooSubVo)("select 'FooSub' DEF3")
                        Assert.That(actual3.Def3, [Is].EqualTo("FooSub"))
                    End Using
                End Sub

                <Test()> Public Sub スーパークラスをマッピングし_サブクラスは追加された項目だけマッピングしたら_サブクラスでもスーパークラスの項目を取得できる()
                    Using db As New TestingDbClient_MappingBoth2
                        Dim actual As FooVo = db.QueryForObject(Of FooVo)("select 'Foo' ABC2")
                        Assert.That(actual.Abc2, [Is].EqualTo("Foo"))
                        Dim actual2 As FooSubVo = db.QueryForObject(Of FooSubVo)("select 'FooSub' ABC2")
                        Assert.That(actual2.Abc2, [Is].EqualTo("FooSub"))
                        Dim actual3 As FooSubVo = db.QueryForObject(Of FooSubVo)("select 'FooSub' DEF3")
                        Assert.That(actual3.Def3, [Is].EqualTo("FooSub"))
                    End Using
                End Sub
            End Class

            Public Class SettingFieldNameCamelizeIrregularTest_複数回継承されたクラスVoの場合 : Inherits DbClientTest
                Private Class HogeVo
                    Private _aaa1 As String
                    Private _bbb2 As String
                    Private _ccc3 As String
                    Private _ddd4 As String
                    Private _eee5 As String
                    Public Property Aaa1() As String
                        Get
                            Return _aaa1
                        End Get
                        Set(ByVal value As String)
                            _aaa1 = value
                        End Set
                    End Property
                    Public Property Bbb2() As String
                        Get
                            Return _bbb2
                        End Get
                        Set(ByVal value As String)
                            _bbb2 = value
                        End Set
                    End Property
                    Public Property Ccc3() As String
                        Get
                            Return _ccc3
                        End Get
                        Set(ByVal value As String)
                            _ccc3 = value
                        End Set
                    End Property
                    Public Property Ddd4() As String
                        Get
                            Return _ddd4
                        End Get
                        Set(ByVal value As String)
                            _ddd4 = value
                        End Set
                    End Property
                    Public Property Eee5() As String
                        Get
                            Return _eee5
                        End Get
                        Set(ByVal value As String)
                            _eee5 = value
                        End Set
                    End Property
                End Class
                Private Class HogeSubVo : Inherits HogeVo
                End Class
                Private Class HogeSubExVo : Inherits HogeSubVo
                End Class

                Private Class TestingDbClient_VoMapping : Inherits DbClient
                    Public Sub New()
                        MyBase.New(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG, Nothing)
                    End Sub
                    Protected Overrides Sub SettingFieldNameCamelizeIrregular(ByVal bind As BindTable)
                        MyBase.SettingFieldNameCamelizeIrregular(bind)
                        Dim vo As New HogeVo
                        bind.IsA(vo).Bind(vo.Bbb2, "BBB2") _
                                    .Bind(vo.Ccc3, "CCC3") _
                                    .Bind(vo.Ddd4, "DDD4") _
                                    .Bind(vo.Eee5, "EEE5")
                        Dim subVo As New HogeSubVo
                        bind.IsA(subVo).Bind(subVo.Aaa1, "AAA1_SUB") _
                                       .Bind(subVo.Ccc3, "CCC3_SUB") _
                                       .Bind(subVo.Eee5, "EEE5_SUB")
                        Dim subExVo As New HogeSubExVo
                        bind.IsA(subExVo).Bind(subExVo.Aaa1, "AAA1_SUB_EX") _
                                         .Bind(subExVo.Bbb2, "BBB2_SUB_EX") _
                                         .Bind(subExVo.Eee5, "EEE5_SUB_EX")
                    End Sub
                End Class

                <Test()> Public Sub 親クラスをマッピングせず_中間クラスと子クラスをマッピングしたら_子クラスでは子クラスの設定が取得できる()
                    Using db As New TestingDbClient_VoMapping
                        Dim actual As HogeSubExVo = db.QueryForObject(Of HogeSubExVo)("select 'HogeSubEx' AAA1_SUB")
                        Assert.That(actual.Aaa1, [Is].Null, "取得できない")
                        Dim actual2 As HogeSubExVo = db.QueryForObject(Of HogeSubExVo)("select 'HogeSubEx' AAA1_SUB_EX")
                        Assert.That(actual2.Aaa1, [Is].EqualTo("HogeSubEx"))
                    End Using
                End Sub

                <Test()> Public Sub 中間クラスをマッピングせず_親クラスと子クラスをマッピングしたら_子クラスでは子クラスの設定が取得できる()
                    Using db As New TestingDbClient_VoMapping
                        Dim actual As HogeSubExVo = db.QueryForObject(Of HogeSubExVo)("select 'HogeSubEx' BBB2")
                        Assert.That(actual.Bbb2, [Is].Null, "取得できない")
                        Dim actual2 As HogeSubExVo = db.QueryForObject(Of HogeSubExVo)("select 'HogeSubEx' BBB2_SUB_EX")
                        Assert.That(actual2.Bbb2, [Is].EqualTo("HogeSubEx"))
                    End Using
                End Sub

                <Test()> Public Sub 中間クラスをマッピングせず_親クラスと子クラスをマッピングしたら_中間クラスでは親クラスの設定が取得できる()
                    Using db As New TestingDbClient_VoMapping
                        Dim actual As HogeSubVo = db.QueryForObject(Of HogeSubVo)("select 'HogeSub' BBB2")
                        Assert.That(actual.Bbb2, [Is].EqualTo("HogeSub"))
                        Dim actual2 As HogeSubVo = db.QueryForObject(Of HogeSubVo)("select 'HogeSub' BBB2_SUB_EX")
                        Assert.That(actual2.Bbb2, [Is].Null, "取得できない")
                    End Using
                End Sub

                <Test()> Public Sub 子クラスをマッピングせず_親クラスと中間クラスをマッピングしたら_子クラスでは中間クラスの設定が取得できる()
                    Using db As New TestingDbClient_VoMapping
                        Dim actual As HogeSubExVo = db.QueryForObject(Of HogeSubExVo)("select 'HogeSubEx' CCC3")
                        Assert.That(actual.Ccc3, [Is].Null, "取得できない")
                        Dim actual2 As HogeSubExVo = db.QueryForObject(Of HogeSubExVo)("select 'HogeSubEx' CCC3_SUB")
                        Assert.That(actual2.Ccc3, [Is].EqualTo("HogeSubEx"))
                    End Using
                End Sub

                <Test()> Public Sub 中間クラスと子クラスをマッピングせず_親クラスだけマッピングしたら_子クラスでは親クラスの設定が取得できる()
                    Using db As New TestingDbClient_VoMapping
                        Dim actual As HogeSubExVo = db.QueryForObject(Of HogeSubExVo)("select 'HogeSubEx' DDD4")
                        Assert.That(actual.Ddd4, [Is].EqualTo("HogeSubEx"))
                    End Using
                End Sub

                <Test()> Public Sub 親クラスと中間クラスと子クラスをマッピングしたら_子クラスでは子クラスの設定が取得できる()
                    Using db As New TestingDbClient_VoMapping
                        Dim actual As HogeSubExVo = db.QueryForObject(Of HogeSubExVo)("select 'HogeSubEx' EEE5")
                        Assert.That(actual.Eee5, [Is].Null, "取得できない")
                        Dim actual2 As HogeSubExVo = db.QueryForObject(Of HogeSubExVo)("select 'HogeSubEx' EEE5_SUB")
                        Assert.That(actual2.Eee5, [Is].Null, "取得できない")
                        Dim actual3 As HogeSubExVo = db.QueryForObject(Of HogeSubExVo)("select 'HogeSubEx' EEE5_SUB_EX")
                        Assert.That(actual3.Eee5, [Is].EqualTo("HogeSubEx"))
                    End Using
                End Sub

                <Test()> Public Sub 親クラスと中間クラスと子クラスをマッピングしたら_中間クラスでは中間クラスの設定が取得できる()
                    Using db As New TestingDbClient_VoMapping
                        Dim actual As HogeSubVo = db.QueryForObject(Of HogeSubVo)("select 'HogeSub' EEE5")
                        Assert.That(actual.Eee5, [Is].Null, "取得できない")
                        Dim actual2 As HogeSubVo = db.QueryForObject(Of HogeSubVo)("select 'HogeSub' EEE5_SUB")
                        Assert.That(actual2.Eee5, [Is].EqualTo("HogeSub"))
                        Dim actual3 As HogeSubVo = db.QueryForObject(Of HogeSubVo)("select 'HogeSub' EEE5_SUB_EX")
                        Assert.That(actual3.Eee5, [Is].Null, "取得できない")
                    End Using
                End Sub             
            End Class
        End Class

        Public Class SQLタイムアウトテスト : Inherits DbClientTest
            <Test()> Public Sub インスタンス生成直後はデフォルト値になっている()
                Dim db As New TestingDbAccess
                Assert.That(db.CommandTimeout, [Is].EqualTo(300), "SQLSERVERの初期値は300秒")
            End Sub

            <Test()> Public Sub USINGの場合の場合でもインスタンス生成直後はデフォルト値になっている()
                Using db As New TestingDbAccess
                    Assert.That(db.CommandTimeout, [Is].EqualTo(300), "SQLSERVERの初期値は300秒")
                End Using
            End Sub

            <Test()> Public Sub SQLタイムアウト時間を変更できる()
                Dim db As New TestingDbAccess
                db.CommandTimeout = 123
                Assert.That(db.CommandTimeout, [Is].EqualTo(123))
            End Sub

            <Test()> Public Sub USINGの場合でもSQLタイムアウト時間を変更できる()
                Using db As New TestingDbAccess
                    db.CommandTimeout = 123
                    Assert.That(db.CommandTimeout, [Is].EqualTo(123))
                End Using
            End Sub

            <Test()> Public Sub 変更後にインスタンスを再度生成するとデフォルト値になる()
                Dim db As New TestingDbAccess
                db.CommandTimeout = 123
                Assert.That(db.CommandTimeout, [Is].EqualTo(123), "変更された")
                db = New TestingDbAccess
                Assert.That(db.CommandTimeout, [Is].EqualTo(300), "初期値になる")
            End Sub

            <Test()> Public Sub USINGでも変更後にインスタンスを再度生成するとデフォルト値になる()
                Using db As New TestingDbAccess
                    db.CommandTimeout = 123
                    Assert.That(db.CommandTimeout, [Is].EqualTo(123), "変更された")
                End Using

                Using db As New TestingDbAccess
                    Assert.That(db.CommandTimeout, [Is].EqualTo(300), "初期値になる")
                End Using
            End Sub

            <Test()> Public Sub USING内で変更後にインスタンスを生成するとそれぞれが値を持つ()
                Using oldDb As New TestingDbAccess
                    oldDb.CommandTimeout = 123
                    Assert.That(oldDb.CommandTimeout, [Is].EqualTo(123), "変更され")
                    Dim newDb As New TestingDbAccess
                    Assert.That(newDb.CommandTimeout, [Is].EqualTo(300), "初期値になる")
                    Assert.That(oldDb.CommandTimeout, [Is].EqualTo(123), "変更されない")
                End Using
            End Sub

            <Test()> Public Sub USING内で変更後にUsingでインスタンス生成するとそれぞれが値を持つ()
                Using oldDb As New TestingDbAccess
                    oldDb.CommandTimeout = 123
                    Assert.That(oldDb.CommandTimeout, [Is].EqualTo(123), "変更される")
                    Using newDb As New TestingDbAccess
                        Assert.That(newDb.CommandTimeout, [Is].EqualTo(300), "初期値")
                        Assert.That(oldDb.CommandTimeout, [Is].EqualTo(123), "変更されない")
                    End Using
                End Using
            End Sub
        End Class
    End Class
End Namespace
