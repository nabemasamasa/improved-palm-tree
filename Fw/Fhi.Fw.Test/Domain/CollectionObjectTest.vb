Imports NUnit.Framework

Namespace Domain
    Public MustInherit Class CollectionObjectTest

#Region "Nested classes..."
        Private Class TestingValueObject : Inherits ValueObjectTest.TestingValueObject
            Public Sub New(ByVal name As String, ByVal zipCode As String, ByVal address As String)
                MyBase.New(name, zipCode, address)
            End Sub
            Public Function UpdateName(name As String) As TestingValueObject
                Return New TestingValueObject(name, Me.ZipCode, Me.Address)
            End Function
            Public Function UpdateZipCode(zipCode As String) As TestingValueObject
                Return New TestingValueObject(Me.Name, zipCode, Me.Address)
            End Function
            Public Function UpdateAddress(address As String) As TestingValueObject
                Return New TestingValueObject(Me.Name, Me.ZipCode, address)
            End Function
            Public Overrides Function ToString() As String
                Return Join(({Name, ZipCode, Address}).ToArray, ",")
            End Function
        End Class
        Private Class TestingCollection : Inherits CollectionObject(Of TestingValueObject)
            Public Sub New()
            End Sub
            Public Sub New(ByVal list As IEnumerable(Of TestingValueObject))
                MyBase.New(list)
            End Sub
            Public Sub New(ByVal src As CollectionObject(Of TestingValueObject))
                ' 引数が親クラスの型のまま（ReSharperで作るとコレ）
                MyBase.New(src)
            End Sub
            Public Overloads Function UpdateItem(index As Integer, updateCallback As Func(Of TestingValueObject, TestingValueObject)) As TestingCollection
                Return MyBase.UpdateItem(Of TestingCollection)(index, updateCallback)
            End Function
            Public Overloads Function UpdateItem(index As Integer, detail As TestingValueObject) As TestingCollection
                Return MyBase.UpdateItem(Of TestingCollection)(index, detail)
            End Function
            Public Overloads Function Add(detail As TestingValueObject) As TestingCollection
                Return MyBase.Add(Of TestingCollection)(detail)
            End Function
            Public Overloads Function AddRange(details As IEnumerable(Of TestingValueObject)) As TestingCollection
                Return MyBase.AddRange(Of TestingCollection)(details)
            End Function
            Public Overloads Function Insert(ByVal index As Integer, ByVal item As TestingValueObject) As TestingCollection
                Return MyBase.Insert(Of TestingCollection)(index, item)
            End Function
            Public Overloads Function Remove(ByVal item As TestingValueObject) As TestingCollection
                Return MyBase.Remove(Of TestingCollection)(item)
            End Function
            Public Overloads Function RemoveAt(ByVal index As Integer) As TestingCollection
                Return MyBase.RemoveAt(Of TestingCollection)(index)
            End Function
            Public Overloads Function Sort(ByVal comparer As IComparer(Of TestingValueObject)) As TestingCollection
                Return MyBase.Sort(Of TestingCollection)(comparer)
            End Function
            Public Overloads Function Sort(ByVal comparison As Comparison(Of TestingValueObject)) As TestingCollection
                Return MyBase.Sort(Of TestingCollection)(comparison)
            End Function
            Public Overloads Function Where(ByVal predicate As Func(Of TestingValueObject, Boolean)) As TestingCollection
                Return MyBase.Where(Of TestingCollection)(predicate)
            End Function
            Public Overloads Function Exchange(ByVal x As Integer, ByVal y As Integer) As TestingCollection
                Return MyBase.Exchange(Of TestingCollection)(x, y)
            End Function
        End Class
        Private Class MySelfConstructorCollection : Inherits CollectionObject(Of TestingValueObject)
            Public Sub New(ByVal list As IEnumerable(Of TestingValueObject))
                MyBase.New(list)
            End Sub
            Public Sub New(ByVal src As MySelfConstructorCollection)
                ' 引数が自分自身の型
                MyBase.New(src)
            End Sub
            Public Shadows Function UpdateItem(index As Integer, updateCallback As Func(Of TestingValueObject, TestingValueObject)) As MySelfConstructorCollection
                Return DirectCast(MyBase.UpdateItem(index, updateCallback), MySelfConstructorCollection)
            End Function
            Public Shadows Function Add(detail As TestingValueObject) As MySelfConstructorCollection
                Return DirectCast(MyBase.Add(detail), MySelfConstructorCollection)
            End Function
        End Class
        Private Class TakeOverCollection : Inherits CollectionObject(Of TestingValueObject)

            Public ReadOnly TakeOverObj As Object

            Public Sub New()
                Me.TakeOverObj = New Object
            End Sub

            Public Sub New(ByVal list As IEnumerable(Of TestingValueObject))
                MyBase.New(list)
                Me.TakeOverObj = New Object
            End Sub

            Public Sub New(ByVal src As TakeOverCollection)
                MyBase.New(src)
                Me.TakeOverObj = If(src.TakeOverObj, New Object)
            End Sub

            Public Shadows Function UpdateItem(index As Integer, updateCallback As Func(Of TestingValueObject, TestingValueObject)) As TakeOverCollection
                Return DirectCast(MyBase.UpdateItem(index, updateCallback), TakeOverCollection)
            End Function
            Public Shadows Function Add(detail As TestingValueObject) As TakeOverCollection
                Return DirectCast(MyBase.Add(detail), TakeOverCollection)
            End Function
        End Class
        Private Class NoConstructorCollection : Inherits CollectionObject(Of TestingValueObject)
            Public Sub New(ByVal list As IEnumerable(Of TestingValueObject))
                MyBase.New(list)
            End Sub
            ' 自身を引数に持つコンストラクタを実装していない（コメントアウト）
            'Public Sub New(ByVal src As CollectionObject(Of TestingValueObject))
            '    MyBase.New(src)
            'End Sub
            Public Overloads Function Add(ByVal item As TestingValueObject) As CollectionObject(Of TestingValueObject)
                Return MyBase.Add(item)
            End Function
            Public Overloads Function UpdateItem(ByVal index As Integer, ByVal updateCallback As Func(Of TestingValueObject, TestingValueObject)) As CollectionObject(Of TestingValueObject)
                Return MyBase.UpdateItem(index, updateCallback)
            End Function
        End Class
#End Region

        <SetUp()> Public Overridable Sub SetUp()
        End Sub

        Public Class 親クラス型のコンストラクタでTest : Inherits CollectionObjectTest

            Private sut As TestingCollection

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                sut = New TestingCollection({New TestingValueObject("c1", "c2", "c3"),
                                             New TestingValueObject("d1", "d2", "d3")})
            End Sub

            <Test()> Public Sub 初期件数は2件()
                Assert.That(sut.Count, [Is].EqualTo(2))
            End Sub

            <Test()> Public Sub Addで追加できるが_追加後のインスタンスは別()
                Dim actual As TestingCollection = sut.Add(New TestingValueObject("e1", "e2", "e3"))
                Assert.That(actual.Item(2).ToString, [Is].EqualTo("e1,e2,e3"))
                Assert.That(actual.Count, [Is].EqualTo(3))
                Assert.That(sut.Count, [Is].EqualTo(2), "不変なので追加元は2件のまま")
            End Sub

            <Test()> Public Sub AddRangeで追加できるが_追加後のインスタンスは別()
                Dim actual As TestingCollection = sut.AddRange({New TestingValueObject("e1", "e2", "e3"),
                                                                New TestingValueObject("f1", "f2", "f3")})
                Assert.That(actual.Item(2).ToString, [Is].EqualTo("e1,e2,e3"))
                Assert.That(actual.Item(3).ToString, [Is].EqualTo("f1,f2,f3"))
                Assert.That(actual.Count, [Is].EqualTo(4))
                Assert.That(sut.Count, [Is].EqualTo(2), "不変なので追加元は2件のまま")
            End Sub

            <Test()> Public Sub 同じindexの入力値を変更できる_変更後のインスタンスは別_使いにくい()
                Dim actual As TestingCollection = sut.UpdateItem(1, sut.Item(1).UpdateName("xx"))
                Assert.That(actual.Item(1).ToString, [Is].EqualTo("xx,d2,d3"))
                Assert.That(actual.Count, [Is].EqualTo(2))
            End Sub

            <Test()> Public Sub 同じindexの入力値を変更できる_変更後のインスタンスは別_このほうが自然()
                Dim actual As TestingCollection = sut.UpdateItem(0, Function(val) val.UpdateAddress("yy"))
                Assert.That(actual.Item(0).ToString, [Is].EqualTo("c1,c2,yy"))
                Assert.That(actual.Count, [Is].EqualTo(2))
            End Sub

            <Test()> Public Sub Insertで挿入できるが_挿入後のインスタンスは別()
                Dim actual As CollectionObject(Of TestingValueObject) = sut.Insert(1, New TestingValueObject("e1", "e2", "e3"))
                Assert.That(actual.Item(0).ToString, [Is].EqualTo("c1,c2,c3"))
                Assert.That(actual.Item(1).ToString, [Is].EqualTo("e1,e2,e3"))
                Assert.That(actual.Item(2).ToString, [Is].EqualTo("d1,d2,d3"))
                Assert.That(actual.Count, [Is].EqualTo(3))
                Assert.That(sut.Count, [Is].EqualTo(2), "不変なので挿入元は2件のまま")
            End Sub

            <Test()> Public Sub Removeで除去できるが_除去後のインスタンスは別()
                Dim actual As CollectionObject(Of TestingValueObject) = sut.RemoveAt(0)
                Assert.That(actual.Item(0).ToString, [Is].EqualTo("d1,d2,d3"))
                Assert.That(actual.Count, [Is].EqualTo(1))
                Assert.That(sut.Count, [Is].EqualTo(2), "不変なので除去元は2件のまま")
            End Sub

            <Test()> Public Sub Sort後のインスタンスは別()
                Dim actual1 As TestingCollection = sut.Add(New TestingValueObject("a1", "a2", "a3"))
                Dim actual2 As CollectionObject(Of TestingValueObject) = actual1.Sort(Function(x, y) x.Name.CompareTo(y.Name))
                Assert.That(actual2.Item(0).ToString, [Is].EqualTo("a1,a2,a3"))
                Assert.That(actual2.Item(1).ToString, [Is].EqualTo("c1,c2,c3"))
                Assert.That(actual2.Item(2).ToString, [Is].EqualTo("d1,d2,d3"))
                Assert.That(actual2.Count, [Is].EqualTo(3))
                Assert.That(actual1.Item(0).ToString, [Is].EqualTo("c1,c2,c3"))
                Assert.That(actual1.Item(1).ToString, [Is].EqualTo("d1,d2,d3"))
                Assert.That(actual1.Item(2).ToString, [Is].EqualTo("a1,a2,a3"))
                Assert.That(actual1.Count, [Is].EqualTo(3))
            End Sub

            <Test()> Public Sub Where後のインスタンスは別()
                Dim actual As CollectionObject(Of TestingValueObject) = sut.Where(Function(v) "d1".Equals(v.Name))
                Assert.That(actual.Item(0).ToString, [Is].EqualTo("d1,d2,d3"))
                Assert.That(actual.Count, [Is].EqualTo(1))
                Assert.That(sut.Count, [Is].EqualTo(2), "不変なのでWhere元は2件のまま")
            End Sub

            <Test()> Public Sub Exchangeで入れ替えできる()
                sut = sut.Add(New TestingValueObject("e1", "e2", "e3"))
                Dim actual As TestingCollection = sut.Exchange(0, 2)
                Assert.That(actual(2).ToString, [Is].EqualTo(sut(0).ToString))
                Assert.That(actual(1).ToString, [Is].EqualTo(sut(1).ToString))
                Assert.That(actual(0).ToString, [Is].EqualTo(sut(2).ToString))
                Assert.That(actual(0).ToString, [Is].EqualTo("e1,e2,e3"))
                Assert.That(actual(1).ToString, [Is].EqualTo("d1,d2,d3"))
                Assert.That(actual(2).ToString, [Is].EqualTo("c1,c2,c3"))
                Assert.That(actual.Count, [Is].EqualTo(3))
            End Sub

        End Class

        Public Class 自分自身の型のコンストラクタTest : Inherits CollectionObjectTest

            Private sut As MySelfConstructorCollection

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                sut = New MySelfConstructorCollection({New TestingValueObject("c1", "c2", "c3"),
                                                       New TestingValueObject("d1", "d2", "d3")})
            End Sub

            <Test()> Public Sub 初期件数は2件()
                Assert.That(sut.Count, [Is].EqualTo(2))
            End Sub

            <Test()> Public Sub Addで追加できるが_追加後のインスタンスは別()
                Dim actual As MySelfConstructorCollection = sut.Add(New TestingValueObject("e1", "e2", "e3"))
                Assert.That(actual.Item(2).ToString, [Is].EqualTo("e1,e2,e3"))
                Assert.That(actual.Count, [Is].EqualTo(3))
                Assert.That(sut.Count, [Is].EqualTo(2), "不変なので追加元は2件のまま")
            End Sub

            <Test()> Public Sub 同じindexの入力値を変更できる_変更後のインスタンスは別()
                Dim actual As MySelfConstructorCollection = sut.UpdateItem(0, Function(val) val.UpdateAddress("yy"))
                Assert.That(actual.Item(0).ToString, [Is].EqualTo("c1,c2,yy"))
                Assert.That(actual.Count, [Is].EqualTo(2))
            End Sub

        End Class

        Public Class SubClassのreadonlyメンバーに引き継げるかTest : Inherits CollectionObjectTest

            Private sut As TakeOverCollection

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                sut = New TakeOverCollection({New TestingValueObject("c1", "c2", "c3"),
                                              New TestingValueObject("d1", "d2", "d3")})
            End Sub

            <Test()> Public Sub Addの戻り値に_引き継げている()
                Dim actual As TakeOverCollection = sut.Add(New TestingValueObject("e1", "e2", "e3"))
                Assert.That(actual.TakeOverObj, [Is].SameAs(sut.TakeOverObj))
            End Sub

            <Test()> Public Sub UpdateItemの戻り値に_引き継げている()
                Dim actual As TakeOverCollection = sut.UpdateItem(0, Function(val) val.UpdateAddress("yy"))
                Assert.That(actual.TakeOverObj, [Is].SameAs(sut.TakeOverObj))
            End Sub

        End Class

        Public Class CollectionObject引数のコンストラクタが呼ばれないとCloneできないTest : Inherits CollectionObjectTest

            Private sut As NoConstructorCollection

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                sut = New NoConstructorCollection({New TestingValueObject("c1", "c2", "c3"),
                                                   New TestingValueObject("d1", "d2", "d3")})
            End Sub

            <Test()> Public Sub Cloneできない_Add呼び出し()
                Try
                    sut.Add(New TestingValueObject("e1", "e2", "e3"))
                    Assert.Fail()
                Catch expected As InvalidProgramException
                    Assert.That(expected.Message, [Is].StringContaining("NoConstructorCollection に.ctor(NoConstructorCollection) が無いと動作しない"))
                End Try
            End Sub

            <Test()> Public Sub Cloneできない_UpdateItem呼び出し()
                Try
                    sut.UpdateItem(0, Function(val) val.UpdateAddress("yy"))
                    Assert.Fail()
                Catch expected As InvalidProgramException
                    Assert.That(expected.Message, [Is].StringContaining("NoConstructorCollection に.ctor(NoConstructorCollection) が無いと動作しない"))
                End Try
            End Sub

        End Class

    End Class
End Namespace
