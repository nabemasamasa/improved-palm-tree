Imports Fhi.Fw.Db
Imports Fhi.Fw.Db.Impl

Namespace Db.Impl
    Public Class TFunyaDao : Inherits DaoEachTableImpl(Of TFunyaVo)

        Public Interface IBehavior
            Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient
            Sub CreateTableTHogeFuga()
            Sub DropTableTHogeFuga()
        End Interface

#Region "Default behavior"
        Public Class SqlServerTFunyaBehavior : Implements TFunyaDao.IBehavior

            Public Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient Implements TFunyaDao.IBehavior.NewDbClient
                Return New TestSqlDbClient(dbFieldNameByPropertyName)
            End Function

            Public Sub CreateTableTHogeFuga() Implements TFunyaDao.IBehavior.CreateTableTHogeFuga
                Dim sql As String = "CREATE TABLE T_FUNYA (" _
                                    & "FUNYA_ID NUMERIC(12,0) NOT NULL, " _
                                    & "FUNYA_NAME varchar(122) NULL, " _
                                    & "UPDATED_USER_ID varchar(10) NOT NULL, " _
                                    & "UPDATED_DATE varchar(10) NOT NULL, " _
                                    & "UPDATED_TIME varchar(8) NOT NULL, " _
                                    & "CONSTRAINT [PK_T_FUNYA] PRIMARY KEY CLUSTERED (" _
                                    & "FUNYA_ID ASC" _
                                    & ") WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" _
                                    & ") ON [PRIMARY]"
                NewDbClient(Nothing).Update(sql)
            End Sub

            Public Sub DropTableTHogeFuga() Implements TFunyaDao.IBehavior.DropTableTHogeFuga
                NewDbClient(Nothing).Update("DROP TABLE T_FUNYA")
            End Sub
        End Class
#End Region

        Private ReadOnly behavior As IBehavior

        Public Sub New()
            Me.new(New SqlServerTFunyaBehavior)
        End Sub

        Public Sub New(ByVal behavior As IBehavior)
            Me.behavior = behavior
        End Sub

        Friend Function InternalNewClient() As DbClient
            Return NewDbClient(Nothing)
        End Function

        Protected Overrides Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient
            Return behavior.NewDbClient(dbFieldNameByPropertyName)
        End Function

        Public Sub CreateTable()
            behavior.CreateTableTHogeFuga()
        End Sub

        Public Sub DropTable()
            behavior.DropTableTHogeFuga()
        End Sub

        Protected Overrides Sub SettingPkField(ByVal table As PkTable(Of TFunyaVo))
            Dim vo As New TFunyaVo
            table.IsA(vo).PkField(vo.FunyaId)
        End Sub
    End Class
End Namespace