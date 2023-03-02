Imports System.Data.Common
Imports Fhi.Fw.Db
Imports NUnit.Framework
Imports System.Data.SqlClient

Namespace Db
    <Category("RequireDB")> Public Class SqlAccessTest

        Private Function NewDb() As DbAccess
            Dim db As New DbAccess(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG)
            db.Open()
            Return db
        End Function

        Private db As DbAccess

        <SetUp()> Public Sub SetUp()
            db = NewDb()
        End Sub

        <TearDown()> Public Sub TearDown()
            db.Close()
        End Sub

        <Test()> Public Sub LockTimeout_ExecuteReaderだとタイムアウト例外が発生しない()
            Dim tableName As String = "SQL_ACCESS_TEST"
            db.ExecuteNonQuery("CREATE TABLE " & tableName & " (ID INT NOT NULL, NAME VARCHAR(20) NULL)")
            Try
                db.BeginTransaction()
                db.ExecuteNonQuery("insert into " & tableName & " values (1,'hoge')")

                Using reader As DbDataReader = db.ExecuteReader("select count(*) from " & tableName)
                    reader.Read()
                    Assert.AreEqual(1, reader.GetInt32(0))
                End Using

                Using db2 As DbAccess = NewDb()
                    db2.ExecuteNonQuery("set lock_timeout 10")

                    Using reader As DbDataReader = db2.ExecuteReader("select count(*) from " & tableName)
                        Assert.Pass("ExecuteReaderだとタイムアウト例外が発生しない")
                    End Using
                End Using
            Finally
                db.Rollback()
                db.ExecuteNonQuery("DROP TABLE SQL_ACCESS_TEST")
            End Try

            Assert.Fail()
        End Sub

        '<Test()> Public Sub LockTimeout_ExecuteReaderだとタイムアウト例外が発生しない_けどDbDataReader_Readメソッドでタイムアウト例外が発生_ReadCommittedSnapshotがoffのとき()
        '    Dim tableName As String = "SQL_ACCESS_TEST"
        '    db.ExecuteNonQuery("CREATE TABLE " & tableName & " (ID INT NOT NULL, NAME VARCHAR(20) NULL)")
        '    Try
        '        db.BeginTransaction()
        '        db.ExecuteNonQuery("insert into " & tableName & " values (1,'hoge')")

        '        Using reader As DbDataReader = db.ExecuteReader("select count(*) from " & tableName)
        '            reader.Read()
        '            Assert.AreEqual(1, reader.GetInt32(0))
        '        End Using

        '        Using db2 As DbAccess = NewDb()
        '            db2.ExecuteNonQuery("set lock_timeout 10")

        '            Using reader As DbDataReader = db2.ExecuteReader("select count(*) from " & tableName)
        '                Try
        '                    reader.Read()
        '                    Assert.Fail()
        '                Catch expected As SqlException
        '                    Assert.Pass("DbDataReader#Readを呼んだらタイムアウト例外が発生する")
        '                End Try
        '            End Using
        '        End Using
        '    Finally
        '        db.Rollback()
        '        db.ExecuteNonQuery("DROP TABLE SQL_ACCESS_TEST")
        '    End Try
        'End Sub

        '<Test()> Public Sub LockTimeout_DbDataAdapterのFillならタイムアウト例外が発生する_ReadCommittedSnapshotがoffのとき()
        '    Dim tableName As String = "SQL_ACCESS_TEST"
        '    db.ExecuteNonQuery("CREATE TABLE " & tableName & " (ID INT NOT NULL, NAME VARCHAR(20) NULL)")
        '    Try
        '        db.BeginTransaction()
        '        db.ExecuteNonQuery("insert into " & tableName & " values (1,'hoge')")

        '        Using reader As DbDataReader = db.ExecuteReader("select count(*) from " & tableName)
        '            reader.Read()
        '            Assert.AreEqual(1, reader.GetInt32(0))
        '        End Using

        '        Using db2 As DbAccess = NewDb()
        '            db2.ExecuteNonQuery("set lock_timeout 10")

        '            Dim dt As New DataTable
        '            Try
        '                db2.Fill("select count(*) from " & tableName, dt)
        '                Assert.Fail()
        '            Catch expected As SqlException
        '                Assert.Pass("DbDataAdapterのFillならタイムアウト例外が発生する")
        '            End Try
        '        End Using
        '    Finally
        '        db.Rollback()
        '        db.ExecuteNonQuery("DROP TABLE SQL_ACCESS_TEST")
        '    End Try
        'End Sub
    End Class
End Namespace