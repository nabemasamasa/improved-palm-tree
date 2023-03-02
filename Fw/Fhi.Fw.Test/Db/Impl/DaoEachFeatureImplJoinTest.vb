Imports Fhi.Fw.Db
Imports Fhi.Fw.Db.Impl
Imports NUnit.Framework

Namespace Db.Impl

    <TestFixture()> <Category("RequireDB")> Public Class DaoEachFeatureImplJoinTest

        Private Class TestingDaoEachFeatureImpl : Inherits DaoEachFeatureImpl
            Public Function FindHogeById(ByVal hogeId As String) As THogeFugaVo
                Dim sql As String = "SELECT * FROM T_HOGE_FUGA WHERE HOGE_ID = @Value AND HOGE_ID = @Value"
                Dim dao As New THogeFugaDao
                Dim db As DbClient = dao.InternalNewClient
                Return db.QueryForObject(Of THogeFugaVo)(sql, hogeId)
            End Function
        End Class

        <SetUp()> Public Sub Setup()
            Dim dao As New THogeFugaDao
            dao.CreateTable()
        End Sub

        <TearDown()> Public Sub TearDown()
            Dim dao As New THogeFugaDao
            dao.DropTable()
        End Sub

        <Test()> Public Sub 同じ埋め込みパラメータ名があっても動く()
            Dim dao As New TestingDaoEachFeatureImpl
            Dim result As THogeFugaVo = dao.FindHogeById("12345")
            Assert.IsNull(result)
        End Sub

    End Class
End Namespace
