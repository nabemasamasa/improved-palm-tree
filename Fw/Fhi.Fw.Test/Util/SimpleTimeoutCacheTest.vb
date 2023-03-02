Imports NUnit.Framework
Imports System.Threading

Namespace Util
    Public MustInherit Class SimpleTimeoutCacheTest
#Region "Testing classes..."
        Private Class TestingVo
            Private _id As Integer
            Private _name As String

            Public Property Id() As Integer
                Get
                    Return _id
                End Get
                Set(ByVal value As Integer)
                    _id = value
                End Set
            End Property

            Public Property Name() As String
                Get
                    Return _name
                End Get
                Set(ByVal value As String)
                    _name = value
                End Set
            End Property
        End Class

        Private Class TestingSimpleTimeoutCache : Inherits SimpleTimeoutCache(Of TestingVo)
            Public MakeTCallback As Func(Of TestingVo)
            Protected Overrides Function MakeT() As TestingVo
                Return MakeTCallback.Invoke
            End Function
            Public Function [Get]() As TestingVo
                Return GetT()
            End Function
            Public Sub SetTimeout(ByVal millis As Long)
                Me.TimeoutMillis = millis
            End Sub
        End Class
#End Region

        <SetUp()> Public Overridable Sub SetUp()

        End Sub

        Public Class DefaultTest : Inherits SimpleTimeoutCacheTest

            <Test()> Public Sub 一度Getした内容はtimeoutされるまで保持されるので_sleep後も同じinstanceになる()
                Dim sut As New TestingSimpleTimeoutCache With {.MakeTCallback = Function() New TestingVo With {.Id = 1, .Name = "2"}}
                Dim first As TestingVo = sut.Get()
                Thread.Sleep(10)
                Dim actual As TestingVo = sut.Get()
                Assert.That(actual, [Is].SameAs(first))
                Assert.That(actual.Id, [Is].EqualTo(1))
                Assert.That(actual.Name, [Is].EqualTo("2"))
            End Sub

            <Test()> Public Sub timeout時間1msで_Getした内容が差し替る()
                Dim sut As New TestingSimpleTimeoutCache With {.MakeTCallback = Function() New TestingVo With {.Id = 3, .Name = "4"}}
                Dim first As TestingVo = sut.Get()
                first.Id = 5
                sut.SetTimeout(1)
                Dim actual As TestingVo = sut.Get()
                While actual Is first
                    Thread.Sleep(1)
                    actual = sut.Get()
                End While
                Assert.That(actual, [Is].Not.SameAs(first))
                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.Id, [Is].EqualTo(3))
                Assert.That(actual.Name, [Is].EqualTo("4"))
            End Sub

        End Class

    End Class
End Namespace
