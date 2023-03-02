Imports NUnit.Framework

Namespace Db
    Public Class FakeAbstractDaoEachTableTest

        Private Class TestingFakeDao : Inherits FakeAbstractDaoEachTable(Of SampleVo)

        End Class

        Private Function NewSampleVo(ByVal HogeId As Int32?, ByVal HogeName As String, ByVal HogeDate As DateTime, ByVal HogeDecimal As Decimal) As SampleVo
            Dim result As New SampleVo
            result.HogeId = HogeId
            result.HogeId = HogeId
            result.HogeDate = HogeDate
            result.HogeDecimal = HogeDecimal
            Return result
        End Function

        Private dao As TestingFakeDao
        <SetUp()> Public Sub Setup()
            dao = New TestingFakeDao
        End Sub

        <Test()> Public Sub FindBy_未設定ならException()
            Dim nullVo As SampleVo = Nothing
            Try
                dao.FindBy(nullVo)
                Assert.Fail()
            Catch expected As NotImplementedException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub FindBy_Resultを設定したら使用可能()
            Dim result As List(Of SampleVo) = New List(Of SampleVo)
            Dim param As SampleVo = NewSampleVo(1, "2", CDate("2010/10/10"), CDec("3.14159"))

            dao.ResultFindBy = result
            Dim actual As List(Of SampleVo) = dao.FindBy(param)

            Assert.AreSame(result, actual)
            Assert.AreEqual(1, dao.ParamFindBy.Count)
            Assert.AreSame(param, dao.ParamFindBy(0))
        End Sub

        <Test()> Public Sub FindByAll_未設定ならException()
            Try
                dao.FindAll()
                Assert.Fail()
            Catch expected As NotImplementedException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub FindByAll_Resultを設定したら使用可能()
            Dim result As List(Of SampleVo) = New List(Of SampleVo)

            dao.ResultFindByAll = result
            Dim actual As List(Of SampleVo) = dao.FindAll()

            Assert.AreSame(result, actual)
        End Sub

        <Test()> Public Sub CountBy_未設定ならException()
            Try
                dao.CountBy(Nothing)
                Assert.Fail()
            Catch expected As NotImplementedException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub CountBy_Resultを設定したら使用可能()
            Dim result As List(Of SampleVo) = New List(Of SampleVo)
            Dim param As SampleVo = NewSampleVo(1, "2", CDate("2010/10/10"), CDec("3.14159"))

            dao.ResultCountBy = 33
            Dim actual As Integer = dao.CountBy(param)

            Assert.AreEqual(33, actual)
            Assert.AreEqual(1, dao.ParamCountBy.Count)
            Assert.AreSame(param, dao.ParamCountBy(0))
        End Sub

        <Test()> Public Sub DeleteBy_未設定ならException()
            Try
                dao.DeleteBy(Nothing)
                Assert.Fail()
            Catch expected As NotImplementedException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub DeleteBy_Resultを設定したら使用可能()
            Dim result As List(Of SampleVo) = New List(Of SampleVo)
            Dim param As SampleVo = NewSampleVo(1, "2", CDate("2010/10/10"), CDec("3.14159"))

            dao.ResultDeleteBy = 44
            Dim actual As Integer = dao.DeleteBy(param)

            Assert.AreEqual(44, actual)
            Assert.AreEqual(1, dao.ParamDeleteBy.Count)
            Assert.AreSame(param, dao.ParamDeleteBy(0))
        End Sub

        <Test()> Public Sub InsertBy_未設定ならException()
            Try
                dao.InsertBy(Nothing)
                Assert.Fail()
            Catch expected As NotImplementedException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub InsertBy_Resultを設定したら使用可能()
            Dim result As List(Of SampleVo) = New List(Of SampleVo)
            Dim param As SampleVo = NewSampleVo(1, "2", CDate("2010/10/10"), CDec("3.14159"))

            dao.ResultInsertBy = 55
            Dim actual As Integer = dao.InsertBy(param)

            Assert.AreEqual(55, actual)
            Assert.AreEqual(1, dao.ParamInsertBy.Count)
            Assert.AreSame(param, dao.ParamInsertBy(0))
        End Sub

        <Test()> Public Sub MakePkVo_未設定ならException()
            Try
                dao.MakePkVo(Nothing)
                Assert.Fail()
            Catch expected As NotImplementedException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub MakePkVo_Resultを設定したら使用可能()
            Dim result As SampleVo = NewSampleVo(1, "2", CDate("2010/10/10"), CDec("3.14159"))

            dao.ResultMakePkVo = result
            Dim actual As SampleVo = dao.MakePkVo("1", "2", "x", "y")

            Assert.AreSame(result, actual)
            Assert.AreEqual(1, dao.ParamMakePkVo.Count)
            Assert.AreEqual(4, dao.ParamMakePkVo(0).Length)
            Assert.AreEqual("y", dao.ParamMakePkVo(0)(3))
        End Sub

        <Test()> Public Sub SetForUpdate_未設定ならException()
            Try
                dao.SetForUpdate(Nothing)
                Assert.Fail()
            Catch expected As NotImplementedException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub UpdateByPk_未設定ならException()
            Try
                dao.UpdateByPk(Nothing)
                Assert.Fail()
            Catch expected As NotImplementedException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub UpdateByPk_Resultを設定したら使用可能()
            Dim result As List(Of SampleVo) = New List(Of SampleVo)
            Dim param As SampleVo = NewSampleVo(1, "2", CDate("2010/10/10"), CDec("3.14159"))

            dao.ResultUpdateByPk = 66
            Dim actual As Integer = dao.UpdateByPk(param)

            Assert.AreEqual(66, actual)
            Assert.AreEqual(1, dao.ParamUpdateByPk.Count)
            Assert.AreSame(param, dao.ParamUpdateByPk(0))
        End Sub

    End Class
End Namespace