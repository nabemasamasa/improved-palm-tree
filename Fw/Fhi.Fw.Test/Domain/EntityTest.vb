Imports NUnit.Framework

Namespace Domain
    Public Class EntityTest

#Region "Testing classes..."
        Private Class TestingEntity : Inherits Entity
            Public Shadows Property Id() As Object
                Get
                    Return MyBase.Id
                End Get
                Set(ByVal value As Object)
                    MyBase.Id = value
                End Set
            End Property
            Public Property Name As String
        End Class
        Private Class TestingEntities : Inherits CollectionJudgeerObject(Of TestingEntity)
            Public Sub New()
            End Sub
            Public Sub New(ByVal initialList As IEnumerable(Of TestingEntity))
                MyBase.New(initialList)
            End Sub
            Public Sub New(ByVal src As CollectionJudgeerObject(Of TestingEntity))
                MyBase.New(src)
            End Sub
        End Class
        Private Class TestingIdEntity(Of T) : Inherits Entity(Of T)
            Public Shadows Property Id() As T
                Get
                    Return MyBase.Id
                End Get
                Set(ByVal value As T)
                    MyBase.Id = value
                End Set
            End Property
        End Class
#End Region

        <TestCase(1)>
        <TestCase("ABC")>
        <TestCase(3.4R)>
        Public Sub IDが同じなら同値扱いになる_PrimitiveID(id As Object)
            Dim sut1 As New TestingEntity With {.Id = id}
            Dim sut2 As New TestingEntity With {.Id = id}
            Assert.That(sut1, [Is].EqualTo(sut2))
            Assert.That(sut1, [Is].Not.SameAs(sut2), "インスタンスは別")
        End Sub

        <TestCase(1)>
        <TestCase(12345)>
        Public Sub IDが同じなら同値扱いになる_ID型指定(id As Integer)
            Dim sut1 As New TestingIdEntity(Of Integer) With {.Id = id}
            Dim sut2 As New TestingIdEntity(Of Integer) With {.Id = id}
            Assert.That(sut1, [Is].EqualTo(sut2))
            Assert.That(sut1, [Is].Not.SameAs(sut2), "インスタンスは別")
        End Sub

        <Test()> Public Sub IDが同じなら同値扱いになる_Objectインスタンス値()
            Dim objId As New Object
            Dim sut1 As New TestingEntity With {.Id = objId}
            Dim sut2 As New TestingEntity With {.Id = objId}
            Assert.That(sut1, [Is].EqualTo(sut2))
            Assert.That(sut1, [Is].Not.SameAs(sut2), "インスタンスは別")
        End Sub

        <Test()> Public Sub 同一インスタンスだから_ID無しでも同値扱いになる()
            Dim sut As New TestingEntity
            Assert.That(sut, [Is].EqualTo(sut))
            Assert.That(sut, [Is].SameAs(sut), "インスタンス同じ")
        End Sub

        <Test()> Public Sub 別インスタンスだから_ID無しは別値扱いになる()
            Dim sut1 As New TestingEntity
            Dim sut2 As New TestingEntity
            Assert.That(sut1, [Is].Not.EqualTo(sut2))
            Assert.That(sut1, [Is].Not.SameAs(sut2), "インスタンスは別")
        End Sub

        <Test()> Public Sub IDがnull同士は_別値扱いになる()
            Dim objId As Object = Nothing
            Dim sut1 As New TestingEntity With {.Id = objId}
            Dim sut2 As New TestingEntity With {.Id = objId}
            Assert.That(sut1, [Is].Not.EqualTo(sut2))
            Assert.That(sut1, [Is].Not.SameAs(sut2), "インスタンスは別")
        End Sub

        <Test()> Public Sub IDの二重設定はエラーにする()
            Dim sut1 As New TestingEntity With {.Id = 1}
            Try
                sut1.Id = 2
                Assert.Fail()
            Catch ex As InvalidOperationException
                Assert.That(ex.Message, [Is].EqualTo("設定済みのEntityIDは変更できない"))
            End Try
        End Sub

        <Test()> Public Sub 一度_equals使ったら_ID設定はエラーにする()
            Dim sut1 As New TestingEntity
            Dim sut2 As New TestingEntity
            Assert.That(sut1, [Is].Not.EqualTo(sut2))
            Try
                sut1.Id = 3
                Assert.Fail()
            Catch ex As InvalidOperationException
                Assert.That(ex.Message, [Is].EqualTo("EntityID設定前にHashCode利用がある. 先にEntityIDを設定すべき"))
            End Try
        End Sub

        <Test()> Public Sub NotifySaveEndでも_IDエラーにならない()
            Dim sut As New TestingEntities({New TestingEntity With {.Id = "3", .Name = "hoge"}})
            sut(0).Name = "fuga"
            sut.SetUpdatedItems()
            sut.NotifySaveEnd()
            sut.SetUpdatedItems()
            Assert.That(sut.HasChanged, [Is].False)
        End Sub

    End Class
End Namespace