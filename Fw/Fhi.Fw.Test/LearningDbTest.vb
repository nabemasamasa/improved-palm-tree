Imports System.Data.Common
Imports Fhi.Fw.Db
Imports Fhi.Fw.Db.Impl.LinkServer
Imports NUnit.Framework
Imports System.Data.SqlClient

<Category("RequireDB")> Public Class LearningDbTest

    <Test()> Public Sub バインド変数_シャープは大丈夫_但しSqlServer限定()
        'DbTestInitializer.Initialize()
        Using dt As New DataTable
            Using db As New DbAccess(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG)
                db.Open()
                db.AddParameter("@Hoge#0", "aaa")
                db.AddParameter("@Hoge#1", "bbbb")
                db.Fill("SELECT @Hoge#0 as xx, @Hoge#1 as yy", dt)
            End Using

            Assert.AreEqual(1, dt.Rows.Count)
            Assert.AreEqual("aaa", dt.Rows(0).Item(0))
            Assert.AreEqual("bbbb", dt.Rows(0).Item(1))
        End Using
    End Sub

    <Test()> Public Sub バインド変数_ダラーも大丈夫_但しSqlServer限定()
        'DbTestInitializer.Initialize()
        Using dt As New DataTable
            Using db As New DbAccess(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG)
                db.Open()
                db.AddParameter("@Hoge$0", "aaa")
                db.AddParameter("@Hoge$1", "bbbb")
                db.Fill("SELECT @Hoge$0 as xx, @Hoge$1 as yy", dt)
            End Using

            Assert.AreEqual(1, dt.Rows.Count)
            Assert.AreEqual("aaa", dt.Rows(0).Item(0))
            Assert.AreEqual("bbbb", dt.Rows(0).Item(1))
        End Using
    End Sub

    <Test()> Public Sub COLUMN句に同名Columnがあると二つ目以降はindex値が付与されて別column名になる_但しSqlServer限定()
        'DbTestInitializer.Initialize()
        Using dt As New DataTable
            Using db As New DbAccess(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG)
                'Using db As New MdbAccess("Provider=Microsoft.Jet.OleDb.4.0:Data Source=c:\aaa.mdb")
                'Using db As New MdbAccess("Provider=Microsoft.ACE.OLEDB.12.0:Data Source=aaa.accdb")
                'Using db As New SqlCeAccess("Data Source='aaa.sdf';Persist Security Info=True; Password=test")
                db.Open()
                db.Fill("SELECT 'FIRST' AS NAME, 'SECOND' AS NAME, 'THIRD' AS NAME", dt)
            End Using

            Assert.AreEqual(1, dt.Rows.Count)
            Assert.AreEqual(3, dt.Columns.Count)

            ' 二つ目以降はindex値が付与される
            Assert.AreEqual("NAME", dt.Columns(0).ColumnName)
            Assert.AreEqual("NAME1", dt.Columns(1).ColumnName)
            Assert.AreEqual("NAME2", dt.Columns(2).ColumnName)

            Assert.AreEqual("FIRST", dt.Rows(0).Item("NAME"))
            Assert.AreEqual("SECOND", dt.Rows(0).Item("NAME1"))
            Assert.AreEqual("THIRD", dt.Rows(0).Item("NAME2"))
        End Using
    End Sub

    <Test()> Public Sub リンクサーバー経由でIBMAS05へ接続した場合_パラメータの型指定し無いと処理がタイムアウトする_正常()
        'DbTestInitializer.Initialize()

        Dim result As New DataTable

        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(DbProvider.SqlServer.ProviderString)
        Using connection As DbConnection = factory.CreateConnection(), _
                dbCommand As DbCommand = factory.CreateCommand

            connection.ConnectionString = TestLinkServerDbClient.CONNECTION_STRING
            connection.Open()

            dbCommand.Connection = connection
            dbCommand.CommandText = "SELECT P02SR_SKF FROM IBMAS05_SKLIBF.IBMAS05.SKLIBF.SKPF02 WHERE P02BHBA10 = @Hoge"

            Dim parameter As DbParameter = factory.CreateParameter
            parameter.ParameterName = "@Hoge"
            parameter.Value = "A"
            parameter.DbType = DbType.AnsiString

            dbCommand.Parameters.Add(parameter)
            dbCommand.CommandTimeout = 3

            Dim reader As DbDataReader = dbCommand.ExecuteReader()
            Do While reader.Read
                'reader.
            Loop
        End Using
    End Sub

    <Test()> Public Sub リンクサーバー経由でIBMAS05へ接続した場合_パラメータの型指定し無いと処理がタイムアウトする_タイムアウト()
        'DbTestInitializer.Initialize()

        Dim result As New DataTable

        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(DbProvider.SqlServer.ProviderString)
        Using connection As DbConnection = factory.CreateConnection(), _
                dbCommand As DbCommand = factory.CreateCommand

            connection.ConnectionString = TestLinkServerDbClient.CONNECTION_STRING
            connection.Open()

            dbCommand.Connection = connection
            dbCommand.CommandText = "SELECT P02SR_SKF FROM IBMAS05_SKLIBF.IBMAS05.SKLIBF.SKPF02 WHERE P02BHBA10 = @Hoge"

            Dim parameter As DbParameter = factory.CreateParameter
            parameter.ParameterName = "@Hoge"
            parameter.Value = "A"

            dbCommand.Parameters.Add(parameter)
            dbCommand.CommandTimeout = 3

            Try
                dbCommand.ExecuteReader()
                Assert.Fail()
            Catch expect As SqlException
                Assert.AreEqual(-2, expect.Number)
            End Try
        End Using
    End Sub

End Class
