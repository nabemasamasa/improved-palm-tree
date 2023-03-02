Imports Fhi.Fw.Domain
Imports NUnit.Framework

Namespace Db.Sql
    Public MustInherit Class SqlParameterTest
#Region "Fake Classes..."
        Private Class FakeBehavior : Implements SqlParameter.IBehavior
            Public ReadOnly ParamsForLogDebug As New List(Of String)
            Public Sub LogDebug(message As String) Implements SqlParameter.IBehavior.LogDebug
                ParamsForLogDebug.Add(message)
            End Sub
        End Class
#End Region
        Private sut As SqlParameter
        Private behavior As FakeBehavior
        <SetUp()>
        Public Overridable Sub Setup()
            behavior = New FakeBehavior
            sut = New SqlParameter(behavior)
        End Sub

        Public Class AddAllToTest : Inherits SqlParameterTest
            Private Class StringPVO : Inherits PrimitiveValueObject(Of String)
                Public Sub New(value As String)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class IntPVO : Inherits PrimitiveValueObject(Of Integer?)
                Public Sub New(value As Integer?)
                    MyBase.New(value)
                End Sub
            End Class
            Private Enum HogeEnum
                First
                Second
            End Enum
            Private Overloads Function ToString(messages As List(Of String)) As String
                Return Join(messages.ToArray, vbCrLf)
            End Function

            Private testDbAccess As DbAccess
            Public Overrides Sub Setup()
                MyBase.Setup()
                testDbAccess = New DbAccess(DbProvider.SqlServer)
            End Sub

            <Test()>
            Public Sub DbAccessへパラメータセットする際にデバッグ出力できる()
                sut.Add("@Id", 1)
                sut.Add("@Name", "musa")
                sut.Add("@Date", DateUtil.ConvDateValueToDateTime("2020/01/23 12:34:56.123"))
                sut.AddAllTo(testDbAccess)

                Assert.That(ToString(behavior.ParamsForLogDebug), [Is].EqualTo(
                    "@Id:1" & vbCrLf &
                    "@Name:musa" & vbCrLf &
                    "@Date:2020/01/23 12:34:56"))
            End Sub

            <Test()>
            Public Sub NullはNullとわかるように出力される()
                sut.Add("@Id", DirectCast(Nothing, Integer?))
                sut.Add("@Name", DirectCast(Nothing, String))
                sut.AddAllTo(testDbAccess)

                Assert.That(ToString(behavior.ParamsForLogDebug), [Is].EqualTo(
                    "@Id:<null>" & vbCrLf &
                    "@Name:<null>"))
            End Sub

            <Test()>
            Public Sub Enumなら名前と値をセットで出力できる()
                sut.Add("@Enum1", HogeEnum.First)
                sut.Add("@Enum2", HogeEnum.Second)
                sut.AddAllTo(testDbAccess)

                Assert.That(ToString(behavior.ParamsForLogDebug), [Is].EqualTo(
                    "@Enum1:0 (First)" & vbCrLf &
                    "@Enum2:1 (Second)"))
            End Sub

            <Test()>
            Public Sub PVOも出力できる()
                sut.Add("@Pvo1", New StringPVO("ぴーぶいおー"))
                sut.Add("@EmptyPvo", New StringPVO(""))
                sut.Add("@Pvo2", New IntPVO(1234))
                sut.AddAllTo(testDbAccess)

                Assert.That(ToString(behavior.ParamsForLogDebug), [Is].EqualTo(
                    "@Pvo1:ぴーぶいおー" & vbCrLf &
                    "@EmptyPvo:" & vbCrLf &
                    "@Pvo2:1234"))
            End Sub

            <Test()>
            Public Sub Nullが入ったPVOはNullで出力される()
                sut.Add("@NullPvo1", New StringPVO(Nothing))
                sut.Add("@NullPvo2", New IntPVO(Nothing))
                sut.AddAllTo(testDbAccess)

                Assert.That(ToString(behavior.ParamsForLogDebug), [Is].EqualTo(
                    "@NullPvo1:<null>" & vbCrLf &
                    "@NullPvo2:<null>"))
            End Sub

        End Class

    End Class
End Namespace