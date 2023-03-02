Imports NUnit.Framework

Namespace Domain
    Public MustInherit Class CollectionValueObjectTest

#Region "Nested classes..."
        Private Class TestingValueObject : Inherits ValueObjectTest.TestingValueObject
            Public Sub New(ByVal name As String, ByVal zipCode As String, ByVal address As String)
                MyBase.New(name, zipCode, address)
            End Sub
            Public Overrides Function ToString() As String
                Return Join(({Name, ZipCode, Address}).ToArray, ",")
            End Function
        End Class
        Private Class NestedValueObject : Inherits ValueObject
            Public ReadOnly StaffName As String
            Public ReadOnly Customer As TestingValueObject
            Public Sub New(staffName As String, customer As TestingValueObject)
                Me.StaffName = staffName
                Me.Customer = customer
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {Me.StaffName, Me.Customer}
            End Function
        End Class
        Private Class NestedValueCollection : Inherits CollectionValueObject(Of NestedValueObject)
            Public Sub New(ByVal list As IEnumerable(Of NestedValueObject))
                MyBase.New(list)
            End Sub
            Public Sub New(ByVal src As CollectionValueObject(Of NestedValueObject))
                MyBase.New(src)
            End Sub
        End Class
        Private Class DisableChangeJudgeerCollectionValueObject(Of T As ValueObject) : Inherits CollectionValueObject(Of T)
            Public Sub New(ByVal initialList As IEnumerable(Of T))
                MyBase.New(initialList)
            End Sub
            Public Sub New(ByVal src As CollectionObject(Of T))
                MyBase.New(src)
            End Sub
        End Class
#End Region

        <SetUp()> Public Overridable Sub SetUp()
        End Sub

        Public Class NestクラスでTest : Inherits CollectionValueObjectTest

            Private sut As NestedValueCollection

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                sut = New NestedValueCollection({New NestedValueObject("c0", New TestingValueObject("c1", "c2", "c3")),
                                                 New NestedValueObject("d0", New TestingValueObject("d1", "d2", "d3"))})
            End Sub

            <Test()> Public Sub Selectできる_ただしValueObject型にしかできない()
                Dim actual As CollectionValueObject(Of TestingValueObject) = sut.Select(Function(v) v.Customer)
                Assert.That(actual.Item(0).ToString, [Is].EqualTo("c1,c2,c3"))
                Assert.That(actual.Item(1).ToString, [Is].EqualTo("d1,d2,d3"))
                Assert.That(actual.Count, [Is].EqualTo(2))
            End Sub

        End Class

        Public Class 重複エラーTest : Inherits CollectionValueObjectTest

            <Test()> Public Sub 同一のValueObjectを二重に追加できない()
                Try
                    Dim a As New CollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c"),
                                                                               New TestingValueObject("a", "b", "c")})
                Catch ex As Exception
                    Assert.That(ex.Message, [Is].StringStarting("同一のキーを含む項目が既に追加されています"))
                End Try
            End Sub

            <Test()> Public Sub IsDisabledChangeJudgeerがtrueなら同一のValueObjectを複数追加できる()
                Dim actual As New DisableChangeJudgeerCollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c"),
                                                                                                    New TestingValueObject("a", "b", "c")})
                Assert.That(actual, [Is].Not.Null)
            End Sub

        End Class

        Public Class 同値判定EqualTest : Inherits CollectionValueObjectTest

            <Test()> Public Sub 型と_Collection中身と_サイズが同じなら_同値判定trueになる()
                Dim a As New CollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c")})
                Dim b As New CollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c")})
                Assert.That(a, [Is].Not.SameAs(b), "インスタンスは違うけど")
                Assert.That(a, [Is].EqualTo(b), "同値判定できる")
            End Sub

            <Test()> Public Sub 型は同じだけど_中身が違うのでfalseになる()
                Dim a As New CollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c")})
                Dim b As New CollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c'")})
                Assert.That(a, [Is].Not.EqualTo(b), "中身が違うからNotEqual")
            End Sub

            <Test()> Public Sub 中身は同じだけど_型が違うのでfalseになる()
                Dim a As CollectionValueObject(Of TestingValueObject) = New CollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c")})
                Dim b As CollectionValueObject(Of TestingValueObject) = New DisableChangeJudgeerCollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c")})
                Assert.That(a, [Is].Not.EqualTo(b), "型が違うからNotEqual")
            End Sub

            <Test()> Public Sub 型は同じだけど_サイズが違うのでfalseになる()
                Dim a As New DisableChangeJudgeerCollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c")})
                Dim b As New DisableChangeJudgeerCollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c"), New TestingValueObject("b", "c", "c")})
                Assert.That(a, [Is].Not.EqualTo(b), "サイズが違うからNotEqual")
            End Sub

        End Class

        Public Class GetHashCodeTest : Inherits CollectionValueObjectTest

            <Test()> Public Sub 要素が0のときでも_Dictionaryのキーになれる()
                Dim key As New CollectionValueObject(Of TestingValueObject)
                Dim map As New Dictionary(Of CollectionValueObject(Of TestingValueObject), String)
                map.Add(key, "c")
                Assert.That(map(key), [Is].EqualTo("c"))
            End Sub

            <Test()> Public Sub 要素がn個でも_Dictionaryのキーになれる()
                Dim key As New CollectionValueObject(Of TestingValueObject)({New TestingValueObject("a", "b", "c"), New TestingValueObject("x", "y", "z")})
                Dim map As New Dictionary(Of CollectionValueObject(Of TestingValueObject), String)
                map.Add(key, "c")
                Assert.That(map(key), [Is].EqualTo("c"))
            End Sub

        End Class

    End Class
End Namespace