Imports Fhi.Fw.Db
Imports Fhi.Fw.Db.Impl

Namespace Db.Impl
    Public Class THogeFugaDao : Inherits DaoEachTableImpl(Of THogeFugaVo)

        Public Interface IBehavior
            Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient
            Sub CreateTableTHogeFuga()
            Sub DropTableTHogeFuga()
        End Interface

#Region "Default behavior"
        Public Class SqlServerTHogeFugaBehavior : Implements THogeFugaDao.IBehavior

            Public Function NewDbClient(ByVal dbFieldNameByPropertyName As Dictionary(Of String, String)) As DbClient Implements THogeFugaDao.IBehavior.NewDbClient
                Return New TestSqlDbClient(dbFieldNameByPropertyName)
            End Function

            Public Sub CreateTableTHogeFuga() Implements THogeFugaDao.IBehavior.CreateTableTHogeFuga
                Dim sql As String = "CREATE TABLE T_HOGE_FUGA (" _
                                    & "HOGE_ID INT NOT NULL, " _
                                    & "HOGE_SUB char(1) NOT NULL, " _
                                    & "HOGE_NAME varchar(122) NULL, " _
                                    & "HOGE_DATE Datetime NULL, " _
                                    & "HOGE_DECIMAL NUMERIC(8,2) NULL, " _
                                    & "IS_HOGE BIT NULL, " _
                                    & "HOGE_ENUM INT NULL, " _
                                    & "UPDATED_USER_ID varchar(10) NULL, " _
                                    & "UPDATED_DATE varchar(10) NULL, " _
                                    & "UPDATED_TIME varchar(8) NULL, " _
                                    & "CONSTRAINT [PK_T_HOGE_FUGA] PRIMARY KEY CLUSTERED (" _
                                    & "HOGE_ID ASC, HOGE_SUB ASC" _
                                    & ") WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" _
                                    & ") ON [PRIMARY]"
                NewDbClient(Nothing).Update(sql)
            End Sub

            Public Sub DropTableTHogeFuga() Implements THogeFugaDao.IBehavior.DropTableTHogeFuga
                NewDbClient(Nothing).Update("DROP TABLE T_HOGE_FUGA")
            End Sub
        End Class
#End Region

        Private ReadOnly behavior As IBehavior

        Public Sub New()
            Me.new(New SqlServerTHogeFugaBehavior)
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

        Protected Overrides Sub SettingPkField(ByVal table As PkTable(Of THogeFugaVo))
            Dim vo As New THogeFugaVo
            table.IsA(vo).PkField(vo.HogeId).PkField(vo.HogeSub)
        End Sub

        Public Function FindByPk(ByVal hogeId As Nullable(Of Integer), ByVal hogeSub As String) As THogeFugaVo
            Return FindByPkMain(hogeId, hogeSub)
        End Function

        Public Function DeleteByPk(ByVal hogeId As Nullable(Of Integer), ByVal hogeSub As String) As Integer
            Return DeleteByPkMain(hogeId, hogeSub)
        End Function
    End Class
End Namespace