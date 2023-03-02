Namespace Db.Impl.LinkServer
    Public Class TestLinkServerDbClient : Inherits DbClient

        ''' <summary>テスト用AS400リンクサーバー接続文字列</summary>
        Public Const CONNECTION_STRING As String = "Server=Auto3tacs;Uid=sa;Pwd=It@manager;"

        Public Sub New()
            Me.New(Nothing)
        End Sub

        Public Sub New(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String))
            MyBase.New(DbProvider.LinkServerOnSqlServer, CONNECTION_STRING, dbFieldNameByPropertyName)
            LinkServerName = "IBMAS13_CYJTESTF"
            AllowUpdate = True
        End Sub

        Protected Overrides Sub SettingFieldNameCamelizeIrregular(ByVal bind As BindTable)
            Dim vo As New Testpf01Vo
            bind.IsA(vo).Bind(vo.Fld001, "FLD001").Bind(vo.Fld002, "FLD002")
        End Sub
    End Class
End Namespace