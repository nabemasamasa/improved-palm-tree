Imports Fhi.Fw.Db

Namespace Db
    Public Class TestSqlDbClient : Inherits DbClient

        Public Sub New()
            Me.New(Nothing)
        End Sub

        Public Sub New(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String))
            MyBase.New(DbProvider.SqlServer, DbTestInitializer.TEST_SQL_SERVER_CONNECT_STIRNG, dbFieldNameByPropertyName)
            AllowUpdate = True
        End Sub
    End Class
End Namespace