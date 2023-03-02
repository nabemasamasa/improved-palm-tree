Imports System.IO
Imports System.Data.SqlServerCe
Imports Fhi.Fw.Db
Imports System.Threading

Namespace Db
    ''' <summary>
    ''' テスト用のDBへ、接続する為の初期設定を担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class DbTestInitializer

        ''' <summary>どこかの Sql Server CE のファイル名</summary>
        Public Const TEST_SQL_SERVER_CE_FILE_NAME As String = "TestCE"

        ''' <summary>どこかの Sql Server に接続する接続文字列</summary>
        Public Const TEST_SQL_SERVER_CONNECT_STIRNG As String = "SERVER=Auto3tsv71;UID=common;PWD=commonmanager;DATABASE=B-Raku_Unit_Test;"

        Public Const TEST_ORIGIN_SQL_SERVER_CONNECT_STIRNG As String = "SERVER=Auto3tsv71;UID=common;PWD=commonmanager;DATABASE=_TEST_B-RAKU;"

        ''' <summary>もう一つの Sql Server に接続する接続文字列</summary>
        Public Const TEST_SQL_SERVER_CONNECT_STIRNG_ANOTHOER As String = "SERVER=Auto3tsv71;UID=common;PWD=commonmanager;DATABASE=B-RAKU_TEST;"

        ''' <summary>どこかの Sql Server CE に接続する接続文字列</summary>
        Public Const TEST_SQL_SERVER_CE_CONNECT_STIRNG As String = "Data Source=" & TEST_SQL_SERVER_CE_FILE_NAME

        Public Shared Sub InitializeSqlServerCe()
            If Not File.Exists(TEST_SQL_SERVER_CE_FILE_NAME) Then
                Dim engine As New SqlCeEngine("Data Source=" & TEST_SQL_SERVER_CE_FILE_NAME)
                engine.CreateDatabase()
            End If
        End Sub

        Public Shared Sub CreateTestDb_SqlServer()
            If ExistsDb_SqlServer(TEST_ORIGIN_SQL_SERVER_CONNECT_STIRNG) Then
                Return
            End If
            Dim dbName As String = ConnectionUtil.GetDatabaseName(TEST_ORIGIN_SQL_SERVER_CONNECT_STIRNG)
            ExecuteNonQuery_OnMasterOfTestDb(String.Format("CREATE DATABASE [{0}] ", dbName))

            EzUtil.logDebug("Database作成中...")
            While Not ExistsDb_SqlServer(TEST_ORIGIN_SQL_SERVER_CONNECT_STIRNG)
                ' 出来るまでwait... 5秒弱!?
                Thread.Sleep(100)
            End While
            EzUtil.logDebug("Database作成完了")
        End Sub

        Public Shared Sub DropTestDb_SqlServer()
            If Not ExistsDb_SqlServer(TEST_ORIGIN_SQL_SERVER_CONNECT_STIRNG) Then
                Return
            End If
            Dim dbName As String = ConnectionUtil.GetDatabaseName(TEST_ORIGIN_SQL_SERVER_CONNECT_STIRNG)
            Dim flag As Boolean = True
            While flag
                Try
                    ExecuteNonQuery_OnMasterOfTestDb(String.Format("DROP DATABASE [{0}] ", dbName))
                    flag = False
                Catch ex As Exception
                    Thread.Sleep(100)
                End Try
            End While
        End Sub

        ''' <summary>
        ''' データベース master でDDLを実行する
        ''' </summary>
        ''' <param name="ddl">DDL</param>
        ''' <remarks></remarks>
        Public Shared Sub ExecuteNonQuery_OnMasterOfTestDb(ByVal ddl As String)
            Using db As New DbAccess(DbProvider.SqlServer, GetConnectStringOnMasterOfTestDb)
                db.Open()
                db.ExecuteNonQuery(ddl)
            End Using
        End Sub

        Private Shared Function GetConnectStringOnMasterOfTestDb() As String

            Dim dbName As String = ConnectionUtil.GetDatabaseName(TEST_ORIGIN_SQL_SERVER_CONNECT_STIRNG)
            Return TEST_ORIGIN_SQL_SERVER_CONNECT_STIRNG.Replace("DATABASE=" & dbName, "DATABASE=master")
        End Function

        Public Shared Function ExistsDb_SqlServer(ByVal connectionString As String) As Boolean
            Try
                Using db As New DbAccess(DbProvider.SqlServer, TEST_ORIGIN_SQL_SERVER_CONNECT_STIRNG)
                    db.Open()
                End Using
            Catch ignore As Exception
                Return False
            End Try
            Return True
        End Function
    End Class
End Namespace
