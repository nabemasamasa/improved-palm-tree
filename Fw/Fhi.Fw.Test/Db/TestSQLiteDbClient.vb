Imports Fhi.Fw.Db

Namespace Db
    Public Class TestSQLiteDbClient : Inherits DbClient

        Public Sub New()
            Me.New(Nothing)
        End Sub

        Public Sub New(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String))
            MyBase.New(DbProvider.SQLite, "Data Source=TestDB;Pooling=true;", dbFieldNameByPropertyName)
            AllowUpdate = True
        End Sub
    End Class
End Namespace