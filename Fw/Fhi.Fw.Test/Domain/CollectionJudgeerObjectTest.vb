Imports Fhi.Fw.TestUtil.DebugString
Imports NUnit.Framework
Imports Fhi.Fw.Util

Namespace Domain
    Public MustInherit Class CollectionJudgeerObjectTest

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
        Private Class TestingCollection : Inherits CollectionJudgeerObject(Of TestingValueObject)
            Public Sub New()
            End Sub
            Public Sub New(ByVal list As IEnumerable(Of TestingValueObject))
                MyBase.New(list)
            End Sub
            Public Sub New(ByVal src As CollectionJudgeerObject(Of TestingValueObject))
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
        End Class
        Private Class MySelfConstructorCollection : Inherits CollectionJudgeerObject(Of TestingValueObject)
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
        Private Class TakeOverCollection : Inherits CollectionJudgeerObject(Of TestingValueObject)

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
        Private Class NoConstructorCollection : Inherits CollectionJudgeerObject(Of TestingValueObject)
            Public Sub New(ByVal list As IEnumerable(Of TestingValueObject))
                MyBase.New(list)
            End Sub
            ' 自身を引数に持つコンストラクタを実装していない（コメントアウト）
            'Public Sub New(ByVal src As CollectionJudgeerObject(Of TestingValueObject))
            '    MyBase.New(src)
            'End Sub
            Public Overloads Function Add(ByVal item As TestingValueObject) As CollectionObject(Of TestingValueObject)
                Return MyBase.Add(item)
            End Function
            Public Overloads Function UpdateItem(ByVal index As Integer, ByVal updateCallback As Func(Of TestingValueObject, TestingValueObject)) As CollectionObject(Of TestingValueObject)
                Return MyBase.UpdateItem(index, updateCallback)
            End Function
        End Class
        Private Class TestingGenericCollection(Of T) : Inherits CollectionJudgeerObject(Of T)
            Public Sub New()
            End Sub
            Public Sub New(ByVal list As IEnumerable(Of T))
                MyBase.New(list)
            End Sub
            Public Sub New(ByVal src As CollectionJudgeerObject(Of T))
                ' 引数が親クラスの型のまま（ReSharperで作るとコレ）
                MyBase.New(src)
            End Sub
            Public MakeKeyCallback As Func(Of T, Object) = Nothing
            Protected Overrides Function MakeKeyForChangeJudgeer(ByVal obj As T) As Object
                If MakeKeyCallback Is Nothing Then
                    Return MyBase.MakeKeyForChangeJudgeer(obj)
                End If
                Return MakeKeyCallback.Invoke(obj)
            End Function
            Public Overloads Function Add(ByVal item As T) As TestingGenericCollection(Of T)
                Return MyBase.Add(Of TestingGenericCollection(Of T))(item)
            End Function
            Public Overloads Function RemoveAt(ByVal index As Integer) As TestingGenericCollection(Of T)
                Return MyBase.RemoveAt(Of TestingGenericCollection(Of T))(index)
            End Function
        End Class
        Private Class HogeVo
            Public Property Id As Integer
            Public Property Name As String
        End Class
        Private Class TestingIgnoreAttributeCollection(Of T) : Inherits CollectionJudgeerObject(Of T)
            Public Sub New(ByVal list As IEnumerable(Of T))
                MyBase.New(list)
            End Sub
            Public Sub New(ByVal src As CollectionJudgeerObject(Of T))
                ' 引数が親クラスの型のまま（ReSharperで作るとコレ）
                MyBase.New(src)
            End Sub
            Public GetAttributeTypesToIgnorePropertyFromChangeCallback As Func(Of Type())
            Protected Overrides Function GetAttributeTypesToIgnorePropertyFromChange() As Type()
                If GetAttributeTypesToIgnorePropertyFromChangeCallback Is Nothing Then
                    Return MyBase.GetAttributeTypesToIgnorePropertyFromChange()
                End If
                Return GetAttributeTypesToIgnorePropertyFromChangeCallback.Invoke
            End Function
        End Class
        Private Class TestingNameAttribute : Inherits Attribute
        End Class
        Private Class AttrVo
            Public Property Id As Integer
            <TestingName()>
            Public Property Name As String
        End Class
#End Region

        <SetUp()> Public Overridable Sub SetUp()
        End Sub

        Public Class 親クラス型のコンストラクタでTest : Inherits CollectionJudgeerObjectTest

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
                Dim actual As CollectionJudgeerObject(Of TestingValueObject) = sut.Insert(1, New TestingValueObject("e1", "e2", "e3"))
                Assert.That(actual.Item(0).ToString, [Is].EqualTo("c1,c2,c3"))
                Assert.That(actual.Item(1).ToString, [Is].EqualTo("e1,e2,e3"))
                Assert.That(actual.Item(2).ToString, [Is].EqualTo("d1,d2,d3"))
                Assert.That(actual.Count, [Is].EqualTo(3))
                Assert.That(sut.Count, [Is].EqualTo(2), "不変なので挿入元は2件のまま")
            End Sub

            <Test()> Public Sub Removeで除去できるが_除去後のインスタンスは別()
                Dim actual As CollectionJudgeerObject(Of TestingValueObject) = sut.RemoveAt(0)
                Assert.That(actual.Item(0).ToString, [Is].EqualTo("d1,d2,d3"))
                Assert.That(actual.Count, [Is].EqualTo(1))
                Assert.That(sut.Count, [Is].EqualTo(2), "不変なので除去元は2件のまま")
            End Sub

            <Test()> Public Sub Sort後のインスタンスは別()
                Dim actual1 As TestingCollection = sut.Add(New TestingValueObject("a1", "a2", "a3"))
                Dim actual2 As CollectionJudgeerObject(Of TestingValueObject) = actual1.Sort(Function(x, y) x.Name.CompareTo(y.Name))
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
                Dim actual As CollectionJudgeerObject(Of TestingValueObject) = sut.Where(Function(v) "d1".Equals(v.Name))
                Assert.That(actual.Item(0).ToString, [Is].EqualTo("d1,d2,d3"))
                Assert.That(actual.Count, [Is].EqualTo(1))
                Assert.That(sut.Count, [Is].EqualTo(2), "不変なのでWhere元は2件のまま")
            End Sub

        End Class

        Public Class 自分自身の型のコンストラクタTest : Inherits CollectionJudgeerObjectTest

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

        Public Class SubClassのreadonlyメンバーに引き継げるかTest : Inherits CollectionJudgeerObjectTest

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

        Public Class CollectionJudgeerObject引数のコンストラクタが呼ばれないとCloneできないTest : Inherits CollectionJudgeerObjectTest

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

        Public Class コンストラクタの初期値以降の変更を追跡できるTest : Inherits CollectionJudgeerObjectTest

            Private sut As TestingCollection

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                sut = New TestingCollection({New TestingValueObject("c1", "c2", "c3"),
                                             New TestingValueObject("d1", "d2", "d3")})
            End Sub

            <Test()> Public Sub Addで追加しても_追跡できる()
                Dim actual As TestingCollection = sut.Add(New TestingValueObject("e1", "e2", "e3"))
                actual.SetUpdatedItems()
                Assert.That(actual.WasDeleted(), [Is].False)
                Assert.That(actual.WasInserted(), [Is].True)
                Assert.That(actual.WasUpdated(), [Is].False)
            End Sub

            <Test()> Public Sub UpdateItemで変更しても_追跡できる()
                Dim actual As TestingCollection = sut.UpdateItem(0, Function(val) val.UpdateAddress("yy"))
                actual.SetUpdatedItems()
                Assert.That(actual.WasDeleted(), [Is].True)
                Assert.That(actual.WasInserted(), [Is].True)
                Assert.That(actual.WasUpdated(), [Is].False)
                Assert.That(actual.ExtractDeletedItems(0).ToString, [Is].EqualTo("c1,c2,c3"))
                Assert.That(actual.ExtractInsertedItems(0).ToString, [Is].EqualTo("c1,c2,yy"))
            End Sub

            <Test()> Public Sub 初期値なしのCollectionに対し_Addしたら_全件追跡できる()
                sut = New TestingCollection()
                Dim actual As TestingCollection = sut.AddRange({New TestingValueObject("e1", "e2", "e3"),
                                                                New TestingValueObject("f1", "f2", "f3")})
                actual.SetUpdatedItems()
                Assert.That(actual.WasDeleted(), [Is].False)
                Assert.That(actual.WasInserted(), [Is].True)
                Assert.That(actual.WasUpdated(), [Is].False)
                Assert.That(actual.ExtractInsertedItems(0).ToString, [Is].EqualTo("e1,e2,e3"))
                Assert.That(actual.ExtractInsertedItems(1).ToString, [Is].EqualTo("f1,f2,f3"))
            End Sub

        End Class

        Public Class 変化点追跡のキーを指定できるTest : Inherits CollectionJudgeerObjectTest

            Private Overloads Shared Function ToString(ParamArray vos As ChangeJudgeer(Of HogeVo).UpdateInfo()) As String
                Dim vos2 As HogeVo() = vos.SelectMany(Function(info, i) New HogeVo() {info.BeforeVo, info.AfterVo}).ToArray
                Return ToString(vos2)
            End Function
            Private Overloads Shared Function ToString(ParamArray vos As HogeVo()) As String
                Dim maker As New DebugStringMaker(Of HogeVo)(Function(defineBy As IDebugStringRuleBinder, vo As HogeVo) _
                                                       defineBy.Bind(vo.Id, vo.Name))
                Return maker.MakeString(vos)
            End Function

            <Test()> Public Sub 通常使用_同一インスタンスの値変更なら_変更_として追跡する()
                Dim vo As HogeVo = New HogeVo With {.Id = 2, .Name = "two"}
                Dim sut As New TestingGenericCollection(Of HogeVo)({vo})
                vo.Name = "TWO"
                sut.SetUpdatedItems()

                Assert.That(sut.WasDeleted(), [Is].False)
                Assert.That(sut.WasInserted(), [Is].False)
                Assert.That(sut.WasUpdated(), [Is].True)
                Assert.That(sut.ExtractUpdatedItems.Length, [Is].EqualTo(1))
                Assert.That(ToString(sut.ExtractUpdatedItems), [Is].EqualTo(
                            "Id Name " & vbCrLf & _
                            " 2 'two'" & vbCrLf & _
                            " 2 'TWO'"))
            End Sub

            <Test()> Public Sub 通常使用_異なるインスタンスの値変更だと_追加削除_として追跡する()
                Dim vo As HogeVo = New HogeVo With {.Id = 2, .Name = "two"}
                Dim sut As New TestingGenericCollection(Of HogeVo)({vo})
                Dim result As TestingGenericCollection(Of HogeVo) = sut.RemoveAt(0)
                Dim actual As TestingGenericCollection(Of HogeVo) = result.Add(New HogeVo With {.Id = 2, .Name = "TWO"})
                actual.SetUpdatedItems()

                Assert.That(actual.WasDeleted(), [Is].True)
                Assert.That(actual.WasInserted(), [Is].True)
                Assert.That(actual.WasUpdated(), [Is].False)
                Assert.That(ToString(sut.ExtractDeletedItems), [Is].EqualTo(
                            "Id Name " & vbCrLf & _
                            " 2 'two'"))
                Assert.That(ToString(sut.ExtractInsertedItems), [Is].EqualTo(
                            "Id Name " & vbCrLf & _
                            " 2 'TWO'"))
            End Sub

            <Test()> Public Sub 追跡キー指定_異なるインスタンスの値変更でも_変更_として追跡できる()
                Dim vo As HogeVo = New HogeVo With {.Id = 2, .Name = "two"}
                Dim sut As New TestingGenericCollection(Of HogeVo)({vo})
                sut.MakeKeyCallback = Function(v) v.Id
                ' コンストラクタで使った古い追跡キーをクリアする
                sut.SupersedeBeforeUpdatedItems()
                Dim result As TestingGenericCollection(Of HogeVo) = sut.RemoveAt(0)
                Dim actual As TestingGenericCollection(Of HogeVo) = result.Add(New HogeVo With {.Id = 2, .Name = "TWO"})
                actual.SetUpdatedItems()

                Assert.That(sut.WasDeleted(), [Is].False)
                Assert.That(sut.WasInserted(), [Is].False)
                Assert.That(sut.WasUpdated(), [Is].True)
                Assert.That(sut.ExtractUpdatedItems.Length, [Is].EqualTo(1))
                Assert.That(ToString(sut.ExtractUpdatedItems), [Is].EqualTo(
                            "Id Name " & vbCrLf & _
                            " 2 'two'" & vbCrLf & _
                            " 2 'TWO'"))
            End Sub

        End Class

        Public Class 変更点の判定から除外する属性を指定できるTest : Inherits CollectionJudgeerObjectTest

            <Test()> Public Sub 値を変更したら_変化点ありになる()
                Dim vos As AttrVo() = {New AttrVo With {.Id = 2, .Name = "two"},
                                        New AttrVo With {.Id = 3, .Name = "three"}}
                Dim sut As New TestingIgnoreAttributeCollection(Of AttrVo)(vos)
                vos(0).Name = "TWO"
                vos(1).Name = "THREE"

                sut.SetUpdatedItems()

                Assert.That(sut.HasChanged, [Is].True)
            End Sub

            <Test()> Public Sub 変更点の判定から除外したプロパティ値を変更しても_変化点はなしのまま()
                Dim vos As AttrVo() = {New AttrVo With {.Id = 2, .Name = "two"},
                                        New AttrVo With {.Id = 3, .Name = "three"}}
                Dim sut As New TestingIgnoreAttributeCollection(Of AttrVo)(vos) With {
                    .GetAttributeTypesToIgnorePropertyFromChangeCallback = Function() {GetType(TestingNameAttribute)}}
                vos(0).Name = "TWO"
                vos(1).Name = "THREE"

                sut.SetUpdatedItems()

                Assert.That(sut.HasChanged, [Is].False)
            End Sub

        End Class

        Private Class Name : Inherits PrimitiveValueObject(Of String)
            Public Sub New(ByVal value As String)
                MyBase.New(value)
            End Sub
        End Class
        Private Class Lv2Entity : Inherits Entity
            Public Property First As Name
            Public Property Last As Name
            Public Sub New()
            End Sub
            Public Sub New(first As String, last As String)
                Me.First = New Name(first)
                Me.Last = New Name(last)
            End Sub
        End Class
        Private Class Lv2Entities : Inherits CollectionJudgeerObject(Of Lv2Entity)
            Public Sub New()
            End Sub
            Public Sub New(ByVal initialList As IEnumerable(Of Lv2Entity))
                MyBase.New(initialList)
            End Sub
            Public Sub New(ByVal src As CollectionJudgeerObject(Of Lv2Entity))
                MyBase.New(src)
            End Sub
            Public Overloads Function AddRange(items As IEnumerable(Of Lv2Entity)) As Lv2Entities
                Return MyBase.AddRange(Of Lv2Entities)(items)
            End Function
        End Class
        Private Class Lv1Entity : Inherits Entity
            Public Property Middle As Name
            Private _lv2s As Lv2Entities
            Public ReadOnly Property Lv2s As Lv2Entities
                Get
                    Return _lv2s
                End Get
            End Property
            Public Sub New()
                Me.New(Nothing)
            End Sub
            Public Sub New(lv2s As IEnumerable(Of Lv2Entity))
                Me.New(New Name(Nothing), lv2s)
            End Sub
            Public Sub New(middle As String, lv2s As IEnumerable(Of Lv2Entity))
                Me.New(New Name(middle), lv2s)
            End Sub
            Public Sub New(middle As Name, lv2s As IEnumerable(Of Lv2Entity))
                Me.Middle = If(middle, New Name(Nothing))
                _lv2s = If(lv2s Is Nothing, New Lv2Entities, New Lv2Entities(lv2s))
            End Sub
        End Class
        Private Class Lv1Entities : Inherits CollectionJudgeerObject(Of Lv1Entity)
            Public Sub New()
            End Sub
            Public Sub New(ByVal initialList As IEnumerable(Of Lv1Entity))
                MyBase.New(initialList)
            End Sub
            Public Sub New(ByVal src As CollectionJudgeerObject(Of Lv1Entity))
                MyBase.New(src)
            End Sub
            Public Overloads Function AddRange(items As IEnumerable(Of Lv1Entity)) As Lv1Entities
                Return MyBase.AddRange(Of Lv1Entities)(items)
            End Function
        End Class
        Private Class Lv1PropertyEntity : Inherits Entity
            Public Property Middle As Name
            <TestingName()>
            Public Property ALv2 As Lv2Entity
            Public Sub New()
                Me.New(Nothing)
            End Sub
            Public Sub New(aLv2 As Lv2Entity)
                Me.New(New Name(Nothing), aLv2)
            End Sub
            Public Sub New(middle As String, aLv2 As Lv2Entity)
                Me.New(New Name(middle), aLv2)
            End Sub
            Public Sub New(middle As Name, aLv2 As Lv2Entity)
                Me.Middle = If(middle, New Name(Nothing))
                _ALv2 = If(aLv2, New Lv2Entity)
            End Sub
        End Class
        Private Class Lv1PropertyEntities : Inherits CollectionJudgeerObject(Of Lv1PropertyEntity)
            Public Sub New()
            End Sub
            Public Sub New(ByVal initialList As IEnumerable(Of Lv1PropertyEntity))
                MyBase.New(initialList)
            End Sub
            Public Sub New(ByVal src As CollectionJudgeerObject(Of Lv1PropertyEntity))
                MyBase.New(src)
            End Sub
            Public Overloads Function AddRange(items As IEnumerable(Of Lv1PropertyEntity)) As Lv1PropertyEntities
                Return MyBase.AddRange(Of Lv1PropertyEntities)(items)
            End Function
            Protected Overrides Function GetAttributeTypesToIncludePropertyFromChange() As Type()
                Return {GetType(TestingNameAttribute)}
            End Function
        End Class
        Private Class Lv1ReadonlyEntity : Inherits Entity
            Public Property Middle As Name
            Private _RLv2 As Lv2Entity
            <TestingName()>
            Public ReadOnly Property ReadonlyLv2 As Lv2Entity
                Get
                    Return _RLv2
                End Get
            End Property
            Public Sub New()
                Me.New(Nothing)
            End Sub
            Public Sub New(aLv2 As Lv2Entity)
                Me.New(New Name(Nothing), aLv2)
            End Sub
            Public Sub New(middle As String, aLv2 As Lv2Entity)
                Me.New(New Name(middle), aLv2)
            End Sub
            Public Sub New(middle As Name, aLv2 As Lv2Entity)
                Me.Middle = If(middle, New Name(Nothing))
                _RLv2 = If(aLv2, New Lv2Entity)
            End Sub
        End Class
        Private Class Lv1ReadonlyEntities : Inherits CollectionJudgeerObject(Of Lv1ReadonlyEntity)
            Public Sub New()
            End Sub
            Public Sub New(ByVal initialList As IEnumerable(Of Lv1ReadonlyEntity))
                MyBase.New(initialList)
            End Sub
            Public Sub New(ByVal src As CollectionJudgeerObject(Of Lv1ReadonlyEntity))
                MyBase.New(src)
            End Sub
            Public Overloads Function AddRange(items As IEnumerable(Of Lv1ReadonlyEntity)) As Lv1ReadonlyEntities
                Return MyBase.AddRange(Of Lv1ReadonlyEntities)(items)
            End Function
            Protected Overrides Function GetAttributeTypesToIncludePropertyFromChange() As Type()
                Return {GetType(TestingNameAttribute)}
            End Function
        End Class

        Public MustInherit Class Lv2EntitiesでTest : Inherits CollectionJudgeerObjectTest

            Private Overloads Shared Function ToString(entities As IEnumerable(Of Lv2Entity)) As String
                Dim maker As New DebugStringMaker(Of Lv2Entity)(
                    Function(defineBy As IDebugStringRuleBinder, entity As Lv2Entity) _
                                                     defineBy.Bind(entity.First, entity.Last))
                Return maker.MakeString(entities)
            End Function

            Private sut As Lv2Entities

            Public Class 空っぽEntitiesTest : Inherits Lv2EntitiesでTest

                Public Overrides Sub SetUp()
                    MyBase.SetUp()
                    sut = New Lv2Entities
                End Sub

                <Test()> Public Sub インスタンスを追加したら_追加した情報を取得できる_1件()
                    sut = sut.AddRange({New Lv2Entity("vv", "uu")})
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].True)
                    Assert.That(sut.WasUpdated, [Is].False)
                    Assert.That(ToString(sut.ExtractInsertedItems), [Is].EqualTo(
                                "First Last" & vbCrLf & _
                                "'vv'  'uu'"))
                End Sub

                <Test()> Public Sub インスタンスを追加したら_追加した情報を取得できる_n件()
                    sut = sut.AddRange({New Lv2Entity("v1", "u1"),
                                        New Lv2Entity("v2", "u2"),
                                        New Lv2Entity("v3", "u3")})
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].True)
                    Assert.That(sut.WasUpdated, [Is].False)
                    Assert.That(ToString(sut.ExtractInsertedItems), [Is].EqualTo(
                                "First Last" & vbCrLf & _
                                "'v1'  'u1'" & vbCrLf & _
                                "'v2'  'u2'" & vbCrLf & _
                                "'v3'  'u3'"))
                End Sub

            End Class

            Public Class 中身有EntitiesTest : Inherits Lv2EntitiesでTest

                Public Overrides Sub SetUp()
                    MyBase.SetUp()
                    sut = New Lv2Entities({New Lv2Entity("v1", "u1"), New Lv2Entity("v2", "u2"), New Lv2Entity("v3", "u3")})
                End Sub

                <Test()> Public Sub インスタンスを追加したら_追加した情報を取得できる_1件()
                    sut = sut.AddRange({New Lv2Entity("v4", "u4")})
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].True)
                    Assert.That(sut.WasUpdated, [Is].False)
                    Assert.That(ToString(sut.ExtractInsertedItems), [Is].EqualTo(
                                "First Last" & vbCrLf & _
                                "'v4'  'u4'"))
                End Sub

                <Test()> Public Sub インスタンスを追加したら_追加した情報を取得できる_n件()
                    sut = sut.AddRange({New Lv2Entity("v2", "u2"),
                                        New Lv2Entity("v3", "u3"),
                                        New Lv2Entity("v4", "u4")})
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].True)
                    Assert.That(sut.WasUpdated, [Is].False)
                    Assert.That(ToString(sut.ExtractInsertedItems), [Is].EqualTo(
                                "First Last" & vbCrLf & _
                                "'v2'  'u2'" & vbCrLf & _
                                "'v3'  'u3'" & vbCrLf & _
                                "'v4'  'u4'"))
                End Sub

                <Test()> Public Sub インスタンスの内部値を変更したら_変更前後を取得できる_1件()
                    sut(0).First = New Name("v3")
                    sut(0).Last = New Name("u3")
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].False)
                    Assert.That(sut.WasUpdated, [Is].True)
                    Assert.That(ToString(sut.ExtractUpdatedItems.SelectMany(Function(info) New Lv2Entity() {info.BeforeVo, info.AfterVo})), [Is].EqualTo(
                                "First Last" & vbCrLf & _
                                "'v1'  'u1'" & vbCrLf & _
                                "'v3'  'u3'"))
                End Sub

                <Test()> Public Sub インスタンスの内部値を変更したら_変更前後を取得できる_n件()
                    sut(1).First = New Name("v4")
                    sut(1).Last = New Name("u4")
                    sut(2).First = New Name("v5")
                    sut(2).Last = New Name("u5")
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].False)
                    Assert.That(sut.WasUpdated, [Is].True)
                    Assert.That(ToString(sut.ExtractUpdatedItems.SelectMany(Function(info) New Lv2Entity() {info.BeforeVo, info.AfterVo})), [Is].EqualTo(
                                "First Last" & vbCrLf & _
                                "'v2'  'u2'" & vbCrLf & _
                                "'v4'  'u4'" & vbCrLf & _
                                "'v3'  'u3'" & vbCrLf & _
                                "'v5'  'u5'"))
                End Sub

            End Class

        End Class
        Public MustInherit Class Lv1EntitiesでTest : Inherits CollectionJudgeerObjectTest

            Private Overloads Shared Function ToString(entities As IEnumerable(Of Lv1Entity)) As String
                Dim maker As New DebugStringMaker(Of Lv1Entity)(
                    Function(defineBy As IDebugStringRuleBinder, entity As Lv1Entity) _
                        defineBy.Bind(entity.Middle).JoinFromSideBySide(entity.Lv2s, Function(sideBySide As IDebugStringRuleBinder, lv2 As Lv2Entity) _
                                                                                         sideBySide.Bind(lv2.First, lv2.Last)))
                Return maker.MakeString(entities)
            End Function

            Private sut As Lv1Entities

            Public Class 空っぽEntitiesTest : Inherits Lv1EntitiesでTest

                Public Overrides Sub SetUp()
                    MyBase.SetUp()
                    sut = New Lv1Entities
                End Sub

                <Test()> Public Sub インスタンスを追加したら_追加した情報を取得できる_1件()
                    sut = sut.AddRange({New Lv1Entity("m", {New Lv2Entity("f", "l")})})
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].True)
                    Assert.That(sut.WasUpdated, [Is].False)
                    Assert.That(ToString(sut.ExtractInsertedItems), [Is].EqualTo(
                                "Middle" & vbCrLf & _
                                "'m'   "), "readonlyなプロパティは値コピーできないのでSideBySide項目が無い")
                    Assert.That(ToString(sut.ExtractInsertedItems.Select(Function(itm) sut.DetectOriginInstance(itm))), [Is].EqualTo(
                                "Middle L.First#0 L.Last#0" & vbCrLf & _
                                "'m'    'f'       'l'     "))
                End Sub

                <Test()> Public Sub インスタンスを追加したら_追加した情報を取得できる_n件()
                    sut = sut.AddRange({New Lv1Entity("v2", {New Lv2Entity("f2", "l2")}),
                                        New Lv1Entity("v3", {}),
                                        New Lv1Entity("v4", {New Lv2Entity("f4", "l4")})})
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].True)
                    Assert.That(sut.WasUpdated, [Is].False)
                    Assert.That(ToString(sut.ExtractInsertedItems), [Is].EqualTo(
                                "Middle" & vbCrLf & _
                                "'v2'  " & vbCrLf & _
                                "'v3'  " & vbCrLf & _
                                "'v4'  "), "readonlyなプロパティは値コピーできないのでSideBySide項目が無い")
                    Assert.That(ToString(sut.ExtractInsertedItems.Select(Function(itm) sut.DetectOriginInstance(itm))), [Is].EqualTo(
                                "Middle L.First#0 L.Last#0" & vbCrLf & _
                                "'v2'   'f2'      'l2'    " & vbCrLf & _
                                "'v3'   null      null    " & vbCrLf & _
                                "'v4'   'f4'      'l4'    "))
                End Sub

            End Class

            Public Class 中身有EntitiesTest : Inherits Lv1EntitiesでTest

                Public Overrides Sub SetUp()
                    MyBase.SetUp()
                    sut = New Lv1Entities({New Lv1Entity("m1", {New Lv2Entity("f1", "l1")}),
                                           New Lv1Entity("m2", {New Lv2Entity("f2", "l2"), New Lv2Entity("f3", "l3"), New Lv2Entity("f4", "l4")}),
                                           New Lv1Entity("m3", {New Lv2Entity("f5", "l5"), New Lv2Entity("f6", "l6")})})
                End Sub

                <Test()> Public Sub インスタンスを追加したら_追加した情報を取得できる_1件()
                    sut = sut.AddRange({New Lv1Entity("m", {New Lv2Entity("f", "l")})})
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].True)
                    Assert.That(sut.WasUpdated, [Is].False)
                    Assert.That(ToString(sut.ExtractInsertedItems), [Is].EqualTo(
                                "Middle" & vbCrLf & _
                                "'m'   "), "readonlyなプロパティは値コピーできないのでSideBySide項目が無い")
                    Assert.That(ToString(sut.ExtractInsertedItems.Select(Function(itm) sut.DetectOriginInstance(itm))), [Is].EqualTo(
                                "Middle L.First#0 L.Last#0" & vbCrLf & _
                                "'m'    'f'       'l'     "))
                End Sub

                <Test()> Public Sub インスタンスを追加したら_追加した情報を取得できる_n件()
                    sut = sut.AddRange({New Lv1Entity("v2", {New Lv2Entity("f2", "l2")}),
                                        New Lv1Entity("v3", {}),
                                        New Lv1Entity("v4", {New Lv2Entity("f4", "l4")})})
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].True)
                    Assert.That(sut.WasUpdated, [Is].False)
                    Assert.That(ToString(sut.ExtractInsertedItems), [Is].EqualTo(
                                "Middle" & vbCrLf & _
                                "'v2'  " & vbCrLf & _
                                "'v3'  " & vbCrLf & _
                                "'v4'  "), "readonlyなプロパティは値コピーできないのでSideBySide項目が無い")
                    Assert.That(ToString(sut.ExtractInsertedItems.Select(Function(itm) sut.DetectOriginInstance(itm))), [Is].EqualTo(
                                "Middle L.First#0 L.Last#0" & vbCrLf & _
                                "'v2'   'f2'      'l2'    " & vbCrLf & _
                                "'v3'   null      null    " & vbCrLf & _
                                "'v4'   'f4'      'l4'    "))
                End Sub

                <Test()> Public Sub インスタンスの内部値を変更したら_変更前後を取得できる_1件()
                    sut(0).Middle = New Name("m4")
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].False)
                    Assert.That(sut.WasUpdated, [Is].True)
                    Assert.That(ToString(sut.ExtractUpdatedItems.SelectMany(Function(info) New Lv1Entity() {info.BeforeVo, info.AfterVo})), [Is].EqualTo(
                                "Middle" & vbCrLf & _
                                "'m1'  " & vbCrLf & _
                                "'m4'  "))
                End Sub

                <Test()> Public Sub Pulibプロパティのシャローコピーなので_Readonlyプロパティの中身が変更されても_変更を追跡できない()
                    sut(0).Lv2s(0).First = New Name("f4")
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].False)
                    Assert.That(sut.WasUpdated, [Is].False)
                End Sub

            End Class

        End Class
        Public MustInherit Class Lv1プロパティEntityでTest : Inherits CollectionJudgeerObjectTest

            Private Overloads Shared Function ToString(entities As IEnumerable(Of Lv1PropertyEntity)) As String
                Dim maker As New DebugStringMaker(Of Lv1PropertyEntity)(
                    Function(defineBy As IDebugStringRuleBinder, entity As Lv1PropertyEntity) _
                                                                   defineBy.Bind(entity.Middle, entity.ALv2.First, entity.ALv2.Last))
                Return maker.MakeString(entities)
            End Function

            Public Class 空っぽEntitiesTest : Inherits Lv1プロパティEntityでTest

                <Test()> Public Sub Entityプロパティ内の値を変更しても_変更した情報を取得できる()
                    Dim sut As New Lv1PropertyEntities
                    sut = sut.AddRange({New Lv1PropertyEntity("m", New Lv2Entity("f", "l"))})
                    sut.SupersedeBeforeUpdatedItems()
                    sut(0).ALv2.First = New Name("F")
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].False)
                    Assert.That(sut.WasUpdated, [Is].True)
                    Assert.That(ToString(sut.ExtractUpdatedItems.Select(Function(info) info.AfterVo)), [Is].EqualTo(
                        "Middle First Last" & vbCrLf & _
                        "'m'    'F'   'l' "))
                End Sub

                <Test()> Public Sub 読み取り専用Entityプロパティ内の値を変更しても_変更した情報を取得できる()
                    Dim sut As New Lv1ReadonlyEntities
                    sut = sut.AddRange({New Lv1ReadonlyEntity("m", New Lv2Entity("f", "l"))})
                    sut.SupersedeBeforeUpdatedItems()
                    sut(0).ReadonlyLv2.Last = New Name("L")
                    sut.SetUpdatedItems()

                    Assert.That(sut.WasDeleted, [Is].False)
                    Assert.That(sut.WasInserted, [Is].False)
                    Assert.That(sut.WasUpdated, [Is].True)
                    Assert.That(sut.ExtractUpdatedItems.Select(Function(info) info.AfterVo.Middle), [Is].EqualTo({New Name("m")}))
                    Assert.That(sut.ExtractUpdatedItems.Select(Function(info) info.AfterVo.ReadonlyLv2.First), [Is].EqualTo({New Name("f")}))
                    Assert.That(sut.ExtractUpdatedItems.Select(Function(info) info.AfterVo.ReadonlyLv2.Last), [Is].EqualTo({New Name("L")}))
                End Sub

            End Class

        End Class

    End Class
End Namespace
