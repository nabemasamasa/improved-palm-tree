Namespace Db.Impl.LinkServer

    Public Class Testpf01DaoImpl : Inherits DaoEachTableImpl(Of Testpf01Vo)

        Protected Overrides Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient
            Return New TestLinkServerDbClient
        End Function

        Protected Overrides Sub SettingPkField(ByVal table As PkTable(Of Testpf01Vo))
            ' 本当は主キーなしのテーブルだが #UpdateByPk()の確認用に
            Dim vo As New Testpf01Vo
            table.IsA(vo).PkField(vo.Fld001)
        End Sub

        Public Function DeleteByPk(ByVal fld001 As String) As Integer
            Return DeleteByPkMain(fld001)
        End Function

    End Class
End Namespace