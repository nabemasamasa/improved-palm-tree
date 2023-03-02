Imports System.Collections.Generic
Imports Fhi.Fw.Db
Imports NUnit.Framework

Namespace Db

    <TestFixture()> Public Class TableExclusionTest

        Private Class FakeTShisakuEventDao
            Inherits FakeAbstractDaoEachTable(Of TShisakuEventVo)
            Implements TShisakuEventDao

            Public Overrides Function MakePkVo(ByVal ParamArray values As Object()) As TShisakuEventVo
                Return Nothing
            End Function

            Private ForUpdate As Boolean
            Private saveVo As TShisakuEventVo
            Private checkVo As TShisakuEventVo
            Public Sub New(ByVal saveVo As TShisakuEventVo, ByVal checkVo As TShisakuEventVo)
                Me.saveVo = saveVo
                Me.checkVo = checkVo
            End Sub
            Public Overrides Function FindBy(where As TShisakuEventVo, Optional selectionCallback As Func(Of SelectionField(Of TShisakuEventVo), TShisakuEventVo, Object) = Nothing) As List(Of TShisakuEventVo)
                Dim results As New List(Of TShisakuEventVo)
                If ForUpdate Then
                    If checkVo IsNot Nothing Then
                        results.Add(checkVo)
                    End If
                Else
                    results.Add(saveVo)
                End If
                Return results
            End Function

            Friend updateValue As TShisakuEventVo
            Public Overrides Function UpdateByPk(ByVal pkWhereAndValue As TShisakuEventVo) As Integer
                updateValue = pkWhereAndValue
                Return -1
            End Function

            Public Overrides Sub SetForUpdate(ByVal ForUpdate As Boolean)
                Me.ForUpdate = ForUpdate
            End Sub

            Public Function DeleteByPk(ByVal shisakuEventCode As String) As Integer Implements TShisakuEventDao.DeleteByPk
                Return -1
            End Function

            Public Function FindByPk(ByVal shisakuEventCode As String) As TShisakuEventVo Implements TShisakuEventDao.FindByPk
                Return Nothing
            End Function

        End Class

        <Test()> Public Sub WasUpdatedBySomeone_開始と終了で変化が無ければfalse()
            Dim vo1 As New TShisakuEventVo
            vo1.UpdatedUserId = "HOGE"
            vo1.UpdatedDate = "1234-56-78"
            vo1.UpdatedTime = "11:22:33"

            Dim dao As TShisakuEventDao = New FakeTShisakuEventDao(vo1, vo1)
            Dim exclusion As New TableExclusion(Of TShisakuEventVo)(dao, New FakeSystemDate("2010/05/05 11:22:33"))
            exclusion.Save("Dummy")
            Assert.IsFalse(exclusion.WasUpdatedBySomeone)
        End Sub

        <Test()> Public Sub WasUpdatedBySomeone_開始と終了でUserIdが違ったらtrue()
            Dim vo1 As New TShisakuEventVo
            vo1.UpdatedUserId = "HOGE"
            vo1.UpdatedDate = "1234-56-78"
            vo1.UpdatedTime = "11:22:33"
            Dim vo2 As New TShisakuEventVo
            vo2.UpdatedUserId = "FUGA"
            vo2.UpdatedDate = "1234-56-78"
            vo2.UpdatedTime = "11:22:33"

            Dim dao As TShisakuEventDao = New FakeTShisakuEventDao(vo1, vo2)
            Dim exclusion As New TableExclusion(Of TShisakuEventVo)(dao, New FakeSystemDate("2010/05/05 11:22:33"))
            exclusion.Save("Dummy")
            Assert.IsTrue(exclusion.WasUpdatedBySomeone)
        End Sub

        <Test()> Public Sub WasUpdatedBySomeone_開始と終了でDateが違ったらtrue()
            Dim vo1 As New TShisakuEventVo
            vo1.UpdatedUserId = "HOGE"
            vo1.UpdatedDate = "1234-56-78"
            vo1.UpdatedTime = "11:22:33"
            Dim vo2 As New TShisakuEventVo
            vo2.UpdatedUserId = "HOGE"
            vo2.UpdatedDate = "1234-56-79"
            vo2.UpdatedTime = "11:22:33"

            Dim dao As TShisakuEventDao = New FakeTShisakuEventDao(vo1, vo2)
            Dim exclusion As New TableExclusion(Of TShisakuEventVo)(dao, New FakeSystemDate("2010/05/05 11:22:33"))
            exclusion.Save("Dummy")
            Assert.IsTrue(exclusion.WasUpdatedBySomeone)
        End Sub

        <Test()> Public Sub WasUpdatedBySomeone_開始と終了でTimeが違ったらtrue()
            Dim vo1 As New TShisakuEventVo
            vo1.UpdatedUserId = "HOGE"
            vo1.UpdatedDate = "1234-56-78"
            vo1.UpdatedTime = "11:22:33"
            Dim vo2 As New TShisakuEventVo
            vo2.UpdatedUserId = "HOGE"
            vo2.UpdatedDate = "1234-56-78"
            vo2.UpdatedTime = "11:22:34"

            Dim dao As TShisakuEventDao = New FakeTShisakuEventDao(vo1, vo2)
            Dim exclusion As New TableExclusion(Of TShisakuEventVo)(dao, New FakeSystemDate("2010/05/05 11:22:33"))
            exclusion.Save("Dummy")
            Assert.IsTrue(exclusion.WasUpdatedBySomeone)
        End Sub

        <Test()> Public Sub WasUpdatedBySomeone_終了でレコードが無かったらtrue()
            Dim vo1 As New TShisakuEventVo
            vo1.UpdatedUserId = "HOGE"
            vo1.UpdatedDate = "1234-56-78"
            vo1.UpdatedTime = "11:22:33"

            Dim dao As TShisakuEventDao = New FakeTShisakuEventDao(vo1, Nothing)
            Dim exclusion As New TableExclusion(Of TShisakuEventVo)(dao, New FakeSystemDate("2010/05/05 11:22:33"))
            exclusion.Save("Dummy")
            Assert.IsTrue(exclusion.WasUpdatedBySomeone)
        End Sub

        <Test()> Public Sub Update_動作確認()
            Dim vo1 As New TShisakuEventVo
            vo1.UpdatedUserId = "HOGE"
            vo1.UpdatedDate = "1234-56-78"
            vo1.UpdatedTime = "10:20:30"

            Dim dao As FakeTShisakuEventDao = New FakeTShisakuEventDao(vo1, vo1)
            Dim exclusion As New TableExclusion(Of TShisakuEventVo)(dao, New FakeSystemDate("2010/05/05 11:22:33"))
            exclusion.Save("Dummy")
            exclusion.WasUpdatedBySomeone()
            exclusion.Update("hoge")

            Assert.AreEqual("hoge", dao.updateValue.UpdatedUserId)
            Assert.AreEqual("2010-05-05", dao.updateValue.UpdatedDate)
            Assert.AreEqual("11:22:33", dao.updateValue.UpdatedTime)
        End Sub

        <Test()> Public Sub UpdateAndSave_動作確認()
            Dim vo1 As New TShisakuEventVo
            vo1.UpdatedUserId = "HOGE"
            vo1.UpdatedDate = "1234-56-78"
            vo1.UpdatedTime = "10:20:30"

            Dim dao As FakeTShisakuEventDao = New FakeTShisakuEventDao(vo1, vo1)
            Dim exclusion As New TableExclusion(Of TShisakuEventVo)(dao, New FakeSystemDate("2010/05/05 11:22:33"))
            exclusion.Save("Dummy")
            exclusion.WasUpdatedBySomeone()
            exclusion.UpdateAndSave("hoge")

            Assert.AreEqual("hoge", dao.updateValue.UpdatedUserId)
            Assert.AreEqual("2010-05-05", dao.updateValue.UpdatedDate)
            Assert.AreEqual("11:22:33", dao.updateValue.UpdatedTime)
        End Sub

        <Test()> Public Sub GetUpdatedUserId_開始と終了でUserIdが違う()
            Dim vo1 As New TShisakuEventVo
            vo1.UpdatedUserId = "HOGE"
            vo1.UpdatedDate = "1234-56-78"
            vo1.UpdatedTime = "11:22:33"
            Dim vo2 As New TShisakuEventVo
            vo2.UpdatedUserId = "FUGA"
            vo2.UpdatedDate = "1234-56-78"
            vo2.UpdatedTime = "11:22:33"

            Dim dao As TShisakuEventDao = New FakeTShisakuEventDao(vo1, vo2)
            Dim exclusion As New TableExclusion(Of TShisakuEventVo)(dao, New FakeSystemDate("2012/03/04 05:06:07"))
            exclusion.Save("Dummy")

            Assert.AreEqual("HOGE", exclusion.GetUpdatedUserId)

            exclusion.WasUpdatedBySomeone()

            Assert.AreEqual("FUGA", exclusion.GetUpdatedUserId)
        End Sub

    End Class
End Namespace
