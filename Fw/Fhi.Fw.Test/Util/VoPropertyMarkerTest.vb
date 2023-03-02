Imports Fhi.Fw.Domain
Imports NUnit.Framework

Namespace Util
    ''' <summary>
    ''' Voのプロパティをマーキングする役割を担うクラスのテストクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class VoPropertyMarkerTest

#Region "テストに使用するVO"
        Protected Class Hoge
            Public Property Id() As Nullable(Of Integer)
            Public Property Name() As String
            Public Property Fx() As Nullable(Of Decimal)
            Public Property Adate() As DateTime?
        End Class

        Protected Class Fuga
            Public Property Id() As String
            Public Property Name() As String
            Public Property Fx() As String
            Public Property Adate() As String
        End Class

        Protected Class FugaVo
            Public Property Id() As Integer
            Public Property Hoge() As Hoge
            Public Property StrArray() As String()
            Public Property StrList() As List(Of String)
            Public Property HogeArray() As Hoge()
            Public Property HogeList() As List(Of Hoge)
        End Class

        Protected Class BooleanVo
            Public Property BoolA() As Boolean?
            Public Property BoolB() As Boolean?
        End Class

        Protected Class BooleanVo2
            Public Property BoolA() As Boolean?
            Public Property BoolB() As Boolean?
            Public Property BoolC() As Boolean?
        End Class

        Protected Class BooleanVo3
            Public Property BoolA() As Boolean?
            Public Property BoolB() As Boolean?
            Public Property BoolC() As Boolean?
            Public Property Name() As String
        End Class

        Protected Class AllVo
            Public Property AInt() As Integer?
            Public Property ALong() As Long?
            Public Property ASingle() As Single?
            Public Property ADouble() As Double?
            Public Property ADecimal() As Decimal?
            Public Property ADateTime() As Date?
            Public Property ABoolean() As Boolean?
            Public Property AString() As String
            Public Property AObject() As Object
            Public Property AByte() As Byte?
            Public Property AIntArray() As Integer?()
            Public Property ALongArray() As Long?()
            Public Property ASingleArray() As Single?()
            Public Property ADoubleArray() As Double?()
            Public Property ADecimalArray() As Decimal?()
            Public Property ADateTimeArray() As Date?()
            Public Property ABooleanArray() As Boolean?()
            Public Property AStringArray() As String()
            Public Property AObjectArray() As Object()
            Public Property AByteArray() As Byte()
            Public Property AIntList() As List(Of Integer?)
            Public Property ALongList() As List(Of Long?)
            Public Property ASingleList() As List(Of Single?)
            Public Property ADoubleList() As List(Of Double?)
            Public Property ADecimalList() As List(Of Decimal?)
            Public Property ADateTimeList() As List(Of Date?)
            Public Property ABooleanList() As List(Of Boolean?)
            Public Property AStringList() As List(Of String)
            Public Property AObjectList() As List(Of Object)
            Public Property ADictionary() As Dictionary(Of String, String)
            Public Property AArrayList() As ArrayList
            Public Property ACollection() As Collection
        End Class

        Protected Enum EnumA
            A_FIRST = 1
        End Enum

        Protected Enum EnumB
            B_FIRST = 1
        End Enum
        Protected Enum EnumC
            A_FIRST = 1
        End Enum

        Protected Class CEnumA
            Public Property First() As EnumA
        End Class
        Protected Class CEnumB
            Public Property First() As EnumB
        End Class

        Private Class EnumVo
            Public Property A() As EnumA
            Public Property B() As EnumB
            Public Property C() As EnumC
        End Class

        Private Class ReadWriteOnly
            Private _idReadonly As Integer?
            Private _nameWriteonly As String

            Public ReadOnly Property IdReadonly() As Integer?
                Get
                    Return _idReadonly
                End Get
            End Property

            Public WriteOnly Property NameWriteonly() As String
                Set(ByVal value As String)
                    _nameWriteonly = value
                End Set
            End Property

            Public Property Value() As String

            Public Sub SetId(ByVal id As Integer)
                _idReadonly = id
            End Sub

            Public Function GetName() As String
                Return _nameWriteonly
            End Function
        End Class

        Private Class LastVo
            Public Property Name() As String
        End Class

        Private Class SecondVo
            Public Property Id() As Integer
            Public Property LastVo() As LastVo
        End Class

        Private Class FirstVo
            Public Property Id() As Integer
            Public Property Name() As String
            Public Property SecondVo() As SecondVo
        End Class

        Private Class InfiniteLoopVo
            Public Property Vo() As InfiniteLoopVo
        End Class

        Private Class PrimitiveInt : Inherits ValueObject
            Public ReadOnly Value As Integer
            Public Sub New(value As Integer)
                Me.Value = value
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {Value}
            End Function
        End Class
        Private Class PrimitiveNullInt : Inherits ValueObject
            Friend ReadOnly Value As Integer?
            Public Sub New(value As Integer?)
                Me.Value = value
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {Value}
            End Function
        End Class
        Private Class PrimitiveString : Inherits ValueObject
            Public ReadOnly Value As String
            Public Sub New(value As String)
                Me.Value = value
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {Value}
            End Function
        End Class
        Private Class PrimitiveAggregate : Inherits ValueObject
            Public ReadOnly Id As PrimitiveInt
            Public ReadOnly Name As PrimitiveString
            Public ReadOnly NullId As PrimitiveNullInt
            Public Sub New(Id As Integer, Name As String, nullId As Integer?)
                Me.Id = New PrimitiveInt(Id)
                Me.Name = New PrimitiveString(Name)
                Me.NullId = New PrimitiveNullInt(nullId)
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {Id, Name, NullId}
            End Function
        End Class
        Private Class PrimitiveAndEmptyAggregate : Inherits PrimitiveAggregate
            Public Sub New()
                Me.New(9, Nothing, Nothing)
            End Sub
            Public Sub New(ByVal Id As Integer, ByVal Name As String, ByVal nullId As Integer?)
                MyBase.New(Id, Name, nullId)
            End Sub
        End Class
        Private Class EmptyAggregate : Inherits PrimitiveAggregate
            Public Sub New()
                MyBase.New(9, Nothing, Nothing)
            End Sub
        End Class
        Private Class ValueObjectAggregate : Inherits ValueObject
            Public ReadOnly Id As PrimitiveInt
            Public ReadOnly Name As PrimitiveString
            Public ReadOnly NullId As PrimitiveNullInt
            Public Sub New(Id As PrimitiveInt, Name As PrimitiveString, nullId As PrimitiveNullInt)
                Me.Id = Id
                Me.Name = Name
                Me.NullId = nullId
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {Id, Name, NullId}
            End Function
        End Class
        Private Class ValueObjectAggregateCollection : Inherits CollectionObject(Of ValueObjectAggregate)
            Public Sub New()
            End Sub
            Public Sub New(ByVal initialList As IEnumerable(Of ValueObjectAggregate))
                MyBase.New(initialList)
            End Sub
            Public Sub New(ByVal src As CollectionObject(Of ValueObjectAggregate))
                MyBase.New(src)
            End Sub
        End Class
        Private Class VoIncludedValueObjectCollection
            Public Property Name As PrimitiveString
            Private _Details As ValueObjectAggregateCollection
            Public Sub New()
                Me.New(Nothing)
            End Sub
            Public Sub New(details As ValueObjectAggregateCollection)
                _Details = If(details, New ValueObjectAggregateCollection)
            End Sub
            Public ReadOnly Property Details As ValueObjectAggregateCollection
                Get
                    Return _Details
                End Get
            End Property
        End Class
        Private Class ValueObjectAndPrimitiveAggregate : Inherits ValueObjectAggregate
            Public Sub New(Id As Integer, Name As String, nullId As Integer?)
                Me.New(New PrimitiveInt(Id), New PrimitiveString(Name), New PrimitiveNullInt(nullId))
            End Sub
            Public Sub New(ByVal Id As PrimitiveInt, ByVal Name As PrimitiveString, ByVal nullId As PrimitiveNullInt)
                MyBase.New(Id, Name, nullId)
            End Sub
        End Class
        Private Class ValueObjectIncludedMySelf : Inherits ValueObject
            ' ↓これが含まれてても問題ないことを精査するので、消さないで
            Public Shared ReadOnly MY_SELF As New ValueObjectIncludedMySelf(Nothing)
            Public ReadOnly Name As PrimitiveString
            Public Sub New(name As String)
                Me.Name = New PrimitiveString(name)
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {Name}
            End Function
        End Class
        Private Class ValueObjectIncludedROProperty : Inherits ValueObject
            Private ReadOnly _name As String
            Public ReadOnly Property Name As String
                Get
                    Return _name
                End Get
            End Property
            Public Sub New(name As String)
                _name = name
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {Name}
            End Function
        End Class

        Private Class VoWithIndexer
            Public Property Name As String
            Private ReadOnly aList As New List(Of Object)
            Public ReadOnly Property Item(i As Integer) As Object
                Get
                    Return aList(i)
                End Get
            End Property
        End Class
        Private Class VoIncludedValueObject
            Public Property Name As PrimitiveString
            Public Property Info As PrimitiveAggregate
        End Class
        Private Class VoIncludeCollection
            Public Property Name As PrimitiveString
            Public Property Details As ICollection(Of SecondVo)
        End Class
        Private Class VoFieldOnlyIncludedPrimitive
            Public Name As PrimitiveString
        End Class

        Private Enum EnumD
            No = 0
            Yes = 1
        End Enum
        Private Class MainConstructorPrivate : Inherits ValueObject
            Public Shared ReadOnly EMPTY As New MainConstructorPrivate(EnumD.No, EnumD.No)
            Public ReadOnly First As EnumD
            Public ReadOnly Last As EnumD
            Public Sub New(first As String, last As String)
                Me.new(If("○".Equals(first), EnumD.Yes, EnumD.No), If("○".Equals(last), EnumD.Yes, EnumD.No))
            End Sub
            Public Sub New(first As Boolean, last As Boolean)
                Me.new(If(first, EnumD.Yes, EnumD.No), If(last, EnumD.Yes, EnumD.No))
            End Sub
            Private Sub New(first As EnumD, last As EnumD)
                Me.First = first
                Me.Last = last
            End Sub
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Return New Object() {First, Last}
            End Function
        End Class
        Private Class ValueObjectEntity : Inherits Entity
            Private _id As PrimitiveInt
            Public Overloads Property Id As PrimitiveInt
                Get
                    Return _id
                End Get
                Set(value As PrimitiveInt)
                    _id = value
                    If value IsNot Nothing Then
                        MyBase.Id = value
                    End If
                End Set
            End Property
            Public Property Name As PrimitiveString
            Public Sub New()
                Me.New(Nothing)
            End Sub
            Public Sub New(Id As PrimitiveInt)
                Me.Id = Id
            End Sub
        End Class
#End Region

        Private marker As VoPropertyMarker

        <SetUp()> Public Sub SetUp()
            marker = New VoPropertyMarker
        End Sub

        Public Class GetPropertyInfoTest : Inherits VoPropertyMarkerTest

            <Test()> Public Sub Boolean型も切り分け出来る事_但し二項目まで()
                Dim vo As New BooleanVo
                marker.MarkVo(vo)
                Assert.AreSame(vo.GetType.GetProperty("BoolA"), marker.GetPropertyInfo(vo.BoolA))
                Assert.AreSame(vo.GetType.GetProperty("BoolB"), marker.GetPropertyInfo(vo.BoolB))
            End Sub

            <Test()> Public Sub Boolean型が三項目以上の場合_Boolean型のプロパティを指定すると例外になる()
                Dim vo As New BooleanVo2
                marker.MarkVo(vo)
                Try
                    marker.GetPropertyInfo(vo.BoolA)
                    Assert.Fail()
                Catch ex As NotSupportedException
                    Assert.That(ex.Message, [Is].EqualTo("Boolean型が3項目以上の場合は、ラムダ式を利用してください"))
                End Try
            End Sub

            <Test()> Public Sub Boolean型が三項目以上でもBoolean型のプロパティを指定しなければ例外にはならない()
                Dim vo As New BooleanVo3
                marker.MarkVo(vo)
                Assert.That(marker.GetPropertyInfo(vo.Name), [Is].SameAs(vo.GetType.GetProperty("Name")))
            End Sub

            <Test()> Public Sub Boolean型_三項目以上でも_ラムダ式を利用すれば解決する()
                Dim vo As New BooleanVo2
                marker.MarkVo(vo)
                'AssertThatのActualとExpectが逆なので、次回修正時に正しい順序にしてほしいです
                With vo.GetType
                    Assert.That(.GetProperty("BoolA"), [Is].EqualTo(marker.GetPropertyInfo(Of Boolean?)(Function() vo.BoolA)))
                    Assert.That(.GetProperty("BoolB"), [Is].EqualTo(marker.GetPropertyInfo(Of Boolean?)(Function() vo.BoolB)))
                    Assert.That(.GetProperty("BoolC"), [Is].EqualTo(marker.GetPropertyInfo(Of Boolean?)(Function() vo.BoolC)))
                End With
            End Sub

            <Test()> Public Sub 対応したプロパティ型()
                Dim vo As New AllVo
                marker.MarkVo(vo)
                With vo.GetType
                    Assert.AreSame(.GetProperty("ABoolean"), marker.GetPropertyInfo(vo.ABoolean))
                    Assert.AreSame(.GetProperty("ADateTime"), marker.GetPropertyInfo(vo.ADateTime))
                    Assert.AreSame(.GetProperty("ADecimal"), marker.GetPropertyInfo(vo.ADecimal))
                    Assert.AreSame(.GetProperty("ASingle"), marker.GetPropertyInfo(vo.ASingle))
                    Assert.AreSame(.GetProperty("ADouble"), marker.GetPropertyInfo(vo.ADouble))
                    Assert.AreSame(.GetProperty("AInt"), marker.GetPropertyInfo(vo.AInt))
                    Assert.AreSame(.GetProperty("ALong"), marker.GetPropertyInfo(vo.ALong))
                    Assert.AreSame(.GetProperty("AString"), marker.GetPropertyInfo(vo.AString))
                    Assert.AreSame(.GetProperty("AObject"), marker.GetPropertyInfo(vo.AObject))
                    Assert.AreSame(.GetProperty("AByte"), marker.GetPropertyInfo(vo.AByte))
                    Assert.AreSame(.GetProperty("ABooleanArray"), marker.GetPropertyInfo(vo.ABooleanArray))
                    Assert.AreSame(.GetProperty("ADateTimeArray"), marker.GetPropertyInfo(vo.ADateTimeArray))
                    Assert.AreSame(.GetProperty("ADecimalArray"), marker.GetPropertyInfo(vo.ADecimalArray))
                    Assert.AreSame(.GetProperty("ASingleArray"), marker.GetPropertyInfo(vo.ASingleArray))
                    Assert.AreSame(.GetProperty("ADoubleArray"), marker.GetPropertyInfo(vo.ADoubleArray))
                    Assert.AreSame(.GetProperty("AIntArray"), marker.GetPropertyInfo(vo.AIntArray))
                    Assert.AreSame(.GetProperty("ALongArray"), marker.GetPropertyInfo(vo.ALongArray))
                    Assert.AreSame(.GetProperty("AStringArray"), marker.GetPropertyInfo(vo.AStringArray))
                    Assert.AreSame(.GetProperty("AObjectArray"), marker.GetPropertyInfo(vo.AObjectArray))
                    Assert.AreSame(.GetProperty("AByteArray"), marker.GetPropertyInfo(vo.AByteArray))
                    Assert.AreSame(.GetProperty("ABooleanList"), marker.GetPropertyInfo(vo.ABooleanList))
                    Assert.AreSame(.GetProperty("ADateTimeList"), marker.GetPropertyInfo(vo.ADateTimeList))
                    Assert.AreSame(.GetProperty("ADecimalList"), marker.GetPropertyInfo(vo.ADecimalList))
                    Assert.AreSame(.GetProperty("ASingleList"), marker.GetPropertyInfo(vo.ASingleList))
                    Assert.AreSame(.GetProperty("ADoubleList"), marker.GetPropertyInfo(vo.ADoubleList))
                    Assert.AreSame(.GetProperty("AIntList"), marker.GetPropertyInfo(vo.AIntList))
                    Assert.AreSame(.GetProperty("ALongList"), marker.GetPropertyInfo(vo.ALongList))
                    Assert.AreSame(.GetProperty("AStringList"), marker.GetPropertyInfo(vo.AStringList))
                    Assert.AreSame(.GetProperty("AObjectList"), marker.GetPropertyInfo(vo.AObjectList))
                    Assert.AreSame(.GetProperty("AArrayList"), marker.GetPropertyInfo(vo.AArrayList))
                    Assert.AreSame(.GetProperty("ADictionary"), marker.GetPropertyInfo(vo.ADictionary))
                    Assert.AreSame(.GetProperty("ACollection"), marker.GetPropertyInfo(vo.ACollection))
                End With
            End Sub

            <Test()> Public Sub 複数のEnum型でも切り分けできる()
                Dim vo As New EnumVo
                marker.MarkVo(vo)
                With vo.GetType
                    'AssertThatのActualとExpectが逆なので、次回修正時に正しい順序にしてほしいです
                    Assert.That(.GetProperty("A"), [Is].SameAs(marker.GetPropertyInfo(vo.A)))
                    Assert.That(.GetProperty("B"), [Is].SameAs(marker.GetPropertyInfo(vo.B)))
                    Assert.That(.GetProperty("C"), [Is].SameAs(marker.GetPropertyInfo(vo.C)))
                End With
            End Sub

            <Test()> Public Sub Readonlyプロパティはマークできないけど_Writeonlyはできる()
                Dim vo As New ReadWriteOnly
                marker.MarkVo(vo)
                With vo.GetType
                    Assert.That(.GetProperty("IdReadonly"), [Is].Not.Null, "プロパティはあるけど")
                    Assert.That(vo.IdReadonly.HasValue, [Is].False, "Readonlyはマーキングできない")
                    'AssertThatのActualとExpectが逆なので、次回修正時に正しい順序にしてほしいです
                    Assert.That(.GetProperty("NameWriteonly"), [Is].SameAs(marker.GetPropertyInfo(vo.GetName)), "Writeonlyプロパティはマーキングできる")
                    Assert.That(.GetProperty("Value"), [Is].SameAs(marker.GetPropertyInfo(vo.Value)))
                End With
            End Sub

            <Test()> Public Sub Voの中にVoがあっても切り分けできる()
                Dim vo As New FirstVo
                marker.MarkVo(vo)
                With vo
                    'AssertThatのActualとExpectが逆なので、次回修正時に正しい順序にしてほしいです
                    Assert.That(.GetType.GetProperty("Id"), [Is].SameAs(marker.GetPropertyInfo(.Id)))
                    Assert.That(.GetType.GetProperty("Name"), [Is].SameAs(marker.GetPropertyInfo(.Name)))
                    Assert.That(.GetType.GetProperty("SecondVo"), [Is].SameAs(marker.GetPropertyInfo(.SecondVo)))
                    Assert.That(.SecondVo.GetType.GetProperty("Id"), [Is].SameAs(marker.GetPropertyInfo(.SecondVo.Id)))
                    Assert.That(.SecondVo.GetType.GetProperty("LastVo"), [Is].SameAs(marker.GetPropertyInfo(.SecondVo.LastVo)))
                    Assert.That(.SecondVo.LastVo.GetType.GetProperty("Name"), [Is].SameAs(marker.GetPropertyInfo(.SecondVo.LastVo.Name)))
                End With
            End Sub

        End Class

        Public Class ContainsTest : Inherits VoPropertyMarkerTest

            <Test()> Public Sub Boolean型も切り分け出来る事_但し二項目まで()
                Dim vo As New BooleanVo
                marker.MarkVo(vo)
                Assert.That(marker.Contains(vo.BoolA), [Is].True)
                Assert.That(marker.Contains(vo.BoolB), [Is].True)
            End Sub

            <Test()> Public Sub Boolean型_三項目以上だと_その項目値はnullになり_それ以外のプロパティはマーキングできる()
                Dim vo As New BooleanVo3
                marker.MarkVo(vo)
                Assert.That(marker.Contains(vo.BoolA), [Is].True)
                Assert.That(marker.Contains(vo.BoolB), [Is].True)
                Assert.That(vo.BoolC, [Is].Null, "三項目目以降はnull")
                Assert.That(marker.Contains(vo.Name), [Is].True)
            End Sub

            <Test()> Public Sub Boolean型_三項目以上でも_ラムダ式を利用すれば解決する()
                Dim vo As New BooleanVo2
                marker.MarkVo(vo)
                Assert.That(marker.Contains(Of Boolean?)(Function() vo.BoolA), [Is].True)
                Assert.That(marker.Contains(Of Boolean?)(Function() vo.BoolB), [Is].True)
                Assert.That(marker.Contains(Of Boolean?)(Function() vo.BoolC), [Is].True)
            End Sub

            <Test()> Public Sub 対応したプロパティ型()
                Dim vo As New AllVo
                marker.MarkVo(vo)
                Assert.That(marker.Contains(vo.ABoolean), [Is].True)
                Assert.That(marker.Contains(vo.ADateTime), [Is].True)
                Assert.That(marker.Contains(vo.ADecimal), [Is].True)
                Assert.That(marker.Contains(vo.ASingle), [Is].True)
                Assert.That(marker.Contains(vo.ADouble), [Is].True)
                Assert.That(marker.Contains(vo.AInt), [Is].True)
                Assert.That(marker.Contains(vo.ALong), [Is].True)
                Assert.That(marker.Contains(vo.AString), [Is].True)
                Assert.That(marker.Contains(vo.AObject), [Is].True)
                Assert.That(marker.Contains(vo.AByte), [Is].True)
                Assert.That(marker.Contains(vo.ABooleanArray), [Is].True)
                Assert.That(marker.Contains(vo.ADateTimeArray), [Is].True)
                Assert.That(marker.Contains(vo.ADecimalArray), [Is].True)
                Assert.That(marker.Contains(vo.ASingleArray), [Is].True)
                Assert.That(marker.Contains(vo.ADoubleArray), [Is].True)
                Assert.That(marker.Contains(vo.AIntArray), [Is].True)
                Assert.That(marker.Contains(vo.ALongArray), [Is].True)
                Assert.That(marker.Contains(vo.AStringArray), [Is].True)
                Assert.That(marker.Contains(vo.AObjectArray), [Is].True)
                Assert.That(marker.Contains(vo.AByteArray), [Is].True)
                Assert.That(marker.Contains(vo.ABooleanList), [Is].True)
                Assert.That(marker.Contains(vo.ADateTimeList), [Is].True)
                Assert.That(marker.Contains(vo.ADecimalList), [Is].True)
                Assert.That(marker.Contains(vo.ASingleList), [Is].True)
                Assert.That(marker.Contains(vo.ADoubleList), [Is].True)
                Assert.That(marker.Contains(vo.AIntList), [Is].True)
                Assert.That(marker.Contains(vo.ALongList), [Is].True)
                Assert.That(marker.Contains(vo.AStringList), [Is].True)
                Assert.That(marker.Contains(vo.AObjectList), [Is].True)
                Assert.That(marker.Contains(vo.AArrayList), [Is].True)
                Assert.That(marker.Contains(vo.ADictionary), [Is].True)
                Assert.That(marker.Contains(vo.ACollection), [Is].True)
            End Sub

            <Test()> Public Sub 複数のEnum型でも切り分けできる()
                Dim vo As New EnumVo
                marker.MarkVo(vo)
                Assert.That(marker.Contains(vo.A), [Is].True)
                Assert.That(marker.Contains(vo.B), [Is].True)
                Assert.That(marker.Contains(vo.C), [Is].True)
            End Sub

            <Test()> Public Sub Voの中にVoがあっても切り分けできる()
                Dim vo As New FirstVo
                marker.MarkVo(vo)
                Assert.That(marker.Contains(vo.Id), [Is].True)
                Assert.That(marker.Contains(vo.Name), [Is].True)
                Assert.That(marker.Contains(vo.SecondVo), [Is].True)
                Assert.That(marker.Contains(vo.SecondVo.Id), [Is].True)
                Assert.That(marker.Contains(vo.SecondVo.LastVo), [Is].True)
                Assert.That(marker.Contains(vo.SecondVo.LastVo.Name), [Is].True)
            End Sub

        End Class

        Public Class ExceptionTest : Inherits VoPropertyMarkerTest

            <Test()> Public Sub 複数のVoにマーキングできない_例外になる()
                Dim vo1 As New Hoge
                Dim vo2 As New Fuga
                marker.MarkVo(vo1)
                Try
                    marker.MarkVo(vo2)
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("このインスタンスはマーク済. #Clear()するか別のインスタンスでマーキングして"))
                End Try
            End Sub

            <Test()> Public Sub 無限ループに陥るVoは_例外になる()
                Dim vo As New InfiniteLoopVo
                Try
                    marker.MarkVo(vo)
                    Assert.Fail()
                Catch expected As InvalidOperationException
                    Assert.That(expected.Message, [Is].EqualTo("Voの階層が深すぎます。無限ループかも"))
                End Try
            End Sub

        End Class

        Public Class CreateMarkedVoTest : Inherits VoPropertyMarkerTest

            <Test()> Public Sub IntのコンストラクタのValueObjectを作成できる()
                Dim actual As PrimitiveInt = marker.CreateMarkedVo(Of PrimitiveInt)()
                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.Value, [Is].EqualTo(0))
            End Sub

            <Test()> Public Sub NullIntのコンストラクタのValueObjectを作成できる()
                Dim actual As PrimitiveNullInt = marker.CreateMarkedVo(Of PrimitiveNullInt)()
                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.Value, [Is].EqualTo(0))
            End Sub

            <Test()> Public Sub StringのコンストラクタのValueObjectを作成できる()
                Dim actual As PrimitiveString = marker.CreateMarkedVo(Of PrimitiveString)()
                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.Value, [Is].EqualTo("0"))
            End Sub

            <Test()> Public Sub 集約だけどPrimitiveのコンストラクタのValueObjectを作成できる_引数なしコンストラクタのみならそれを使う()
                Dim actual As EmptyAggregate = marker.CreateMarkedVo(Of EmptyAggregate)()
                Assert.That(actual, [Is].Not.Null)
                Assert.That(actual.Id.Value, [Is].EqualTo(9))
            End Sub

        End Class

        Public Class GetValue_マークした値でフィールド値を取得できる : Inherits VoPropertyMarkerTest

            <Test()> Public Sub IntのValueObjectのField値を取得できる()
                Dim markedVo As ValueObjectAggregate = marker.CreateMarkedVo(Of ValueObjectAggregate)()
                Assert.That(marker.GetValue(markedVo.Id.Value, New ValueObjectAggregate(New PrimitiveInt(3), Nothing, Nothing)), [Is].EqualTo(3))
            End Sub

            <Test()> Public Sub IntのValueObject値を取得できる()
                Dim markedVo As ValueObjectAggregate = marker.CreateMarkedVo(Of ValueObjectAggregate)()
                Assert.That(marker.GetValue(markedVo.Id, New ValueObjectAggregate(New PrimitiveInt(3), Nothing, Nothing)), [Is].EqualTo(New PrimitiveInt(3)))
            End Sub

            <Test()> Public Sub StringのValueObjectのField値を取得できる()
                Dim markedVo As ValueObjectAggregate = marker.CreateMarkedVo(Of ValueObjectAggregate)()
                Assert.That(marker.GetValue(markedVo.Name.Value, New ValueObjectAggregate(Nothing, New PrimitiveString("abc"), Nothing)), [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub StringのValueObject値を取得できる()
                Dim markedVo As ValueObjectAggregate = marker.CreateMarkedVo(Of ValueObjectAggregate)()
                Assert.That(marker.GetValue(markedVo.Name, New ValueObjectAggregate(Nothing, New PrimitiveString("abc"), Nothing)), [Is].EqualTo(New PrimitiveString("abc")))
            End Sub

            <Test()> Public Sub StringのValueObjectのField値を取得できる_Primitiveコンストラクタで内部でValueObject化()
                Dim markedVo As PrimitiveAggregate = marker.CreateMarkedVo(Of PrimitiveAggregate)()
                Assert.That(marker.GetValue(markedVo.Name.Value, New PrimitiveAggregate(Nothing, "abc", Nothing)), [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub StringのValueObject値を取得できる_Primitiveコンストラクタで内部でValueObject化()
                Dim markedVo As PrimitiveAggregate = marker.CreateMarkedVo(Of PrimitiveAggregate)()
                Assert.That(marker.GetValue(markedVo.Name, New PrimitiveAggregate(Nothing, "abc", Nothing)), [Is].EqualTo(New PrimitiveString("abc")))
            End Sub

            <Test()> Public Sub 静的変数で自分自身を内包してても_値を取得できる()
                Dim markedVo As ValueObjectIncludedMySelf = marker.CreateMarkedVo(Of ValueObjectIncludedMySelf)()
                Assert.That(marker.GetValue(markedVo.Name, New ValueObjectIncludedMySelf("abc")), [Is].EqualTo(New PrimitiveString("abc")))
                Assert.That(ValueObjectIncludedMySelf.MY_SELF.Name.Value, [Is].Null, "消さないでね")
            End Sub

            <Test()> Public Sub ReadonlyなProperty値を取得できる()
                Dim markedVo As ValueObjectIncludedROProperty = marker.CreateMarkedVo(Of ValueObjectIncludedROProperty)()
                Dim arg As New ValueObjectIncludedROProperty("abc")
                Assert.That(marker.GetValue(markedVo.Name, arg), [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub 集約だけどPrimitiveのコンストラクタのValueObjectで取得できる()
                Dim marked As PrimitiveAggregate = marker.CreateMarkedVo(Of PrimitiveAggregate)()
                Dim arg As New PrimitiveAggregate(12, "3", 45)
                Assert.That(marker.GetValue(marked.Id, arg), [Is].EqualTo(New PrimitiveInt(12)))
                Assert.That(marker.GetValue(marked.Name, arg), [Is].EqualTo(New PrimitiveString("3")))
                Assert.That(marker.GetValue(marked.NullId, arg), [Is].EqualTo(New PrimitiveNullInt(45)))
            End Sub

            <Test()> Public Sub 集約だけどPrimitiveのコンストラクタのValueObjectで取得できる_引数なしコンストラクタを無視できる()
                Dim marked As PrimitiveAndEmptyAggregate = marker.CreateMarkedVo(Of PrimitiveAndEmptyAggregate)()
                Dim arg As New PrimitiveAndEmptyAggregate(12, "3", 45)
                Assert.That(marker.GetValue(marked.Id, arg), [Is].EqualTo(New PrimitiveInt(12)))
                Assert.That(marker.GetValue(marked.Name, arg), [Is].EqualTo(New PrimitiveString("3")))
                Assert.That(marker.GetValue(marked.NullId, arg), [Is].EqualTo(New PrimitiveNullInt(45)))
            End Sub

            <Test()> Public Sub 集約で_ValueObjectのコンストラクタで取得できる()
                Dim marked As ValueObjectAggregate = marker.CreateMarkedVo(Of ValueObjectAggregate)()
                Dim arg As New ValueObjectAggregate(New PrimitiveInt(12), New PrimitiveString("3"), New PrimitiveNullInt(45))
                Assert.That(marker.GetValue(marked.Id, arg), [Is].EqualTo(New PrimitiveInt(12)))
                Assert.That(marker.GetValue(marked.Name, arg), [Is].EqualTo(New PrimitiveString("3")))
                Assert.That(marker.GetValue(marked.NullId, arg), [Is].EqualTo(New PrimitiveNullInt(45)))
            End Sub

            <Test()> Public Sub 集約で_ValueObjectのコンストラクタで取得できる_PrimitiveのコンストラクタがあってもベストマッチはValueObject引数()
                Dim marked As ValueObjectAndPrimitiveAggregate = marker.CreateMarkedVo(Of ValueObjectAndPrimitiveAggregate)()
                Dim arg As New ValueObjectAndPrimitiveAggregate(12, "3", 45)
                Assert.That(marker.GetValue(marked.Id, arg), [Is].EqualTo(New PrimitiveInt(12)))
                Assert.That(marker.GetValue(marked.Name, arg), [Is].EqualTo(New PrimitiveString("3")))
                Assert.That(marker.GetValue(marked.NullId, arg), [Is].EqualTo(New PrimitiveNullInt(45)))
            End Sub

        End Class

        Public Class GetValue_マークした値でプロパティ値を取得できる : Inherits VoPropertyMarkerTest

            <Test()> Public Sub ネストしてる値でも取得できる_2階層()
                Dim markedVo As FirstVo = marker.CreateMarkedVo(Of FirstVo)()
                Assert.That(marker.GetValue(markedVo.SecondVo.Id, New FirstVo With {.SecondVo = New SecondVo With {.Id = 5}}), [Is].EqualTo(5))
            End Sub

            <Test()> Public Sub ネストしてる値でも取得できる_3階層()
                Dim markedVo As FirstVo = marker.CreateMarkedVo(Of FirstVo)()
                Assert.That(marker.GetValue(markedVo.SecondVo.LastVo.Name, New FirstVo With {.SecondVo = New SecondVo With {.LastVo = New LastVo With {.Name = "abc"}}}), [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub インデクサ付きプロパティは無視して取得できる()
                Dim markedVo As VoWithIndexer = marker.CreateMarkedVo(Of VoWithIndexer)()
                Assert.That(marker.GetValue(markedVo.Name, New VoWithIndexer With {.Name = "abc"}), [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub ValueObjectを内包するVoで_ValueObject値を取得できる()
                Dim markedVo As VoIncludedValueObject = marker.CreateMarkedVo(Of VoIncludedValueObject)()
                Assert.That(marker.GetValue(markedVo.Info.Name.Value, New VoIncludedValueObject With {.Info = New PrimitiveAggregate(Nothing, "qw", Nothing)}), [Is].EqualTo("qw"))
            End Sub

            <Test()> Public Sub ICollection型をマークできて_さらに設定したList値を取得できる()
                Dim markedVo As VoIncludeCollection = marker.CreateMarkedVo(Of VoIncludeCollection)()
                Dim vos As List(Of SecondVo) = ({New SecondVo}).ToList
                Assert.That(marker.GetValue(markedVo.Details, New VoIncludeCollection With {.Details = vos}), [Is].SameAs(vos))
            End Sub

            <Test()> Public Sub ICollection型をマークできて_さらに設定した配列値を取得できる()
                Dim markedVo As VoIncludeCollection = marker.CreateMarkedVo(Of VoIncludeCollection)()
                Dim vos As SecondVo() = {New SecondVo}
                Assert.That(marker.GetValue(markedVo.Details, New VoIncludeCollection With {.Details = vos}), [Is].SameAs(vos))
            End Sub

            <Test()> Public Sub コンストラクタで初期化しているCollectionObjectで_値を取得できる()
                Dim markedVo As VoIncludedValueObjectCollection = marker.CreateMarkedVo(Of VoIncludedValueObjectCollection)()
                Dim arg As New VoIncludedValueObjectCollection(New ValueObjectAggregateCollection)
                Assert.That(marker.GetValue(markedVo.Details, arg), [Is].SameAs(arg.Details))
            End Sub

            <Test()> Public Sub VoのFieldにもマークできる()
                Dim markedVo As VoFieldOnlyIncludedPrimitive = marker.CreateMarkedVo(Of VoFieldOnlyIncludedPrimitive)()
                Dim arg As New VoFieldOnlyIncludedPrimitive With {.Name = New PrimitiveString("abc")}
                Assert.That(marker.GetValue(markedVo.Name.Value, arg), [Is].EqualTo("abc"))
            End Sub

            <TestCase(True, False)>
            <TestCase(False, True)>
            <TestCase(True, True)>
            <TestCase(False, False)>
            Public Sub 適切なコンストラクタがPrivateでも_メンバー変数値を一意に初期化できる(first As Boolean, last As Boolean)
                Dim markedVo As MainConstructorPrivate = marker.CreateMarkedVo(Of MainConstructorPrivate)()
                Dim arg As New MainConstructorPrivate(first, last)
                Assert.That(marker.GetValue(markedVo.First, arg), [Is].EqualTo(If(first, EnumD.Yes, EnumD.No)))
                Assert.That(marker.GetValue(markedVo.Last, arg), [Is].EqualTo(If(last, EnumD.Yes, EnumD.No)))
            End Sub

            <Test()> Public Sub コンストラクタで初期化するValueObject形式なんだけど_別途プロパティ値もあるEntityも初期化できる()
                Dim marked As ValueObjectEntity = marker.CreateMarkedVo(Of ValueObjectEntity)()
                Dim vo As New ValueObjectEntity(New PrimitiveInt(3)) With {.Name = New PrimitiveString("4")}
                Assert.That(marker.GetValue(marked.Id, vo), [Is].EqualTo(New PrimitiveInt(3)))
                Assert.That(marker.GetValue(marked.Name, vo), [Is].EqualTo(New PrimitiveString("4")))
            End Sub

            Private Class DateDate
                Public Property Date1 As DateTime
                Public Property Date2 As DateTime
            End Class

            <Test()> Public Sub DateTimeが二つあっても_ぞれぞれ別々に取得できる()
                Dim marked As DateDate = marker.CreateMarkedVo(Of DateDate)()
                Dim vo As New DateDate With {.Date1 = CDate("2012/3/4"), .Date2 = CDate("2013/4/5")}
                Assert.That(marker.GetValue(marked.Date1, vo), [Is].EqualTo(CDate("2012/3/4")))
                Assert.That(marker.GetValue(marked.Date2, vo), [Is].EqualTo(CDate("2013/4/5")))
            End Sub

        End Class

        Public Class マーキングのテスト : Inherits VoPropertyMarkerTest

            Private Class ROFieldVO : Inherits ValueObject
                Public ReadOnly Name As PrimitiveString
                Public ReadOnly Address As PrimitiveString
                Public Sub New()
                    ' このコンストラクタが利用されていないことを確認している. 消さないこと
                End Sub
                Public Sub New(name As PrimitiveString, address As PrimitiveString)
                    Me.Name = name
                    Me.Address = address
                End Sub
                Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                    Return New Object() {Name, Address}
                End Function
            End Class
            Private Class EntityIncludedROFieldVO : Inherits Entity
                Public Property SubId As PrimitiveInt
                Public Property Info As ROFieldVO
                Public Sub New()
                    ' このコンストラクタが利用されていないことを確認している. 消さないこと
                End Sub
                Public Sub New(subId As PrimitiveInt, info As ROFieldVO)
                    Me.SubId = subId
                    Me.Info = info
                End Sub
            End Class
            <Test()> Public Sub 読取専用フィールドをマーキングできる()
                Dim marked As EntityIncludedROFieldVO = marker.CreateMarkedVo(Of EntityIncludedROFieldVO)()
                Dim vo As New EntityIncludedROFieldVO(New PrimitiveInt(5), New ROFieldVO(name:=New PrimitiveString("a"), address:=New PrimitiveString("b")))
                Assert.That(marker.GetValue(marked.SubId, vo), [Is].EqualTo(New PrimitiveInt(5)))
                Assert.That(marker.GetValue(marked.Info.Name, vo), [Is].EqualTo(New PrimitiveString("a")))
                Assert.That(marker.GetValue(marked.Info.Address, vo), [Is].EqualTo(New PrimitiveString("b")))
            End Sub

            Private Class ROPropertyVO : Inherits ValueObject
                Private _Name As PrimitiveString
                Public ReadOnly Property Name As PrimitiveString
                    Get
                        Return _Name
                    End Get
                End Property
                Private _Address As PrimitiveString
                Public ReadOnly Property Address As PrimitiveString
                    Get
                        Return _Address
                    End Get
                End Property
                Public Sub New()
                    ' このコンストラクタが利用されていないことを確認している. 消さないこと
                End Sub
                Public Sub New(name As PrimitiveString, address As PrimitiveString)
                    Me._Name = name
                    Me._Address = address
                End Sub
                Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                    Return New Object() {Name, Address}
                End Function
            End Class
            Private Class EntityIncludedROPropertyVO : Inherits Entity
                Public Property SubId As PrimitiveInt
                Public Property Info As ROPropertyVO
                Public Sub New()
                    ' このコンストラクタが利用されていないことを確認している. 消さないこと
                End Sub
                Public Sub New(subId As PrimitiveInt, info As ROPropertyVO)
                    Me.SubId = subId
                    Me.Info = info
                End Sub
            End Class
            <Test()> Public Sub 読取専用プロパティをマーキングできる()
                Dim marked As EntityIncludedROPropertyVO = marker.CreateMarkedVo(Of EntityIncludedROPropertyVO)()
                Dim vo As New EntityIncludedROPropertyVO(New PrimitiveInt(5), New ROPropertyVO(name:=New PrimitiveString("a"), address:=New PrimitiveString("b")))
                Assert.That(marker.GetValue(marked.SubId, vo), [Is].EqualTo(New PrimitiveInt(5)))
                Assert.That(marker.GetValue(marked.Info.Name, vo), [Is].EqualTo(New PrimitiveString("a")))
                Assert.That(marker.GetValue(marked.Info.Address, vo), [Is].EqualTo(New PrimitiveString("b")))
            End Sub

            Private Class InvalidVO : Inherits ValueObject
                Public Property Name As PrimitiveString
                Public Property Address As PrimitiveString
                Public Sub New()
                    ' このコンストラクタが利用されていないことを確認している. 消さないこと
                End Sub
                Public Sub New(name As PrimitiveString, address As PrimitiveString)
                    Me.Name = name
                    Me.Address = address
                End Sub
                Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                    Return New Object() {Name, Address}
                End Function
            End Class
            Private Class EntityIncludedInvalidVO : Inherits Entity
                Public Property SubId As PrimitiveInt
                Public Property Info As InvalidVO
                Public Sub New()
                    ' このコンストラクタが利用されていないことを確認している. 消さないこと
                End Sub
                Public Sub New(subId As PrimitiveInt, info As InvalidVO)
                    Me.SubId = subId
                    Me.Info = info
                End Sub
            End Class
            <Test()> Public Sub 読取専用じゃなくてもプロパティをマーキングできる()
                Dim marked As EntityIncludedInvalidVO = marker.CreateMarkedVo(Of EntityIncludedInvalidVO)()
                Dim vo As New EntityIncludedInvalidVO(New PrimitiveInt(5), New InvalidVO(name:=New PrimitiveString("a"), address:=New PrimitiveString("b")))
                Assert.That(marker.GetValue(marked.SubId, vo), [Is].EqualTo(New PrimitiveInt(5)))
                Assert.That(marker.GetValue(marked.Info.Name, vo), [Is].EqualTo(New PrimitiveString("a")))
                Assert.That(marker.GetValue(marked.Info.Address, vo), [Is].EqualTo(New PrimitiveString("b")))
            End Sub

        End Class

        Public Class GetValue_PrimitiveValueObjectでTest : Inherits VoPropertyMarkerTest

            Private Class Sheet : Inherits PrimitiveValueObject(Of String)
                Public Sub New(ByVal value As String)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class Weapon : Inherits PrimitiveValueObject(Of String)
                Public Count As Integer ' 消さないで（これが出力されないことを確認する）
                Public Sub New(ByVal value As String)
                    MyBase.New(value)
                End Sub
            End Class
            Private Class SheetEntity
                Public Name As String
                Public Title As Sheet
                Public Addr As Weapon
            End Class

            <Test()> Public Sub マーキングして取得できる()
                Dim marked As SheetEntity = marker.CreateMarkedVo(Of SheetEntity)()
                Dim vo As New SheetEntity With {.Name = "aaa", .Title = New Sheet("bbb"), .Addr = New Weapon("ccc") With {.Count = 123}}
                Assert.That(marker.GetValue(marked.Name, vo), [Is].EqualTo("aaa"))
                Assert.That(marker.GetValue(marked.Title, vo), [Is].EqualTo(New Sheet("bbb")))
                Assert.That(marker.GetValue(marked.Addr, vo), [Is].EqualTo(New Weapon("ccc")))
                Assert.That(marker.GetValue(marked.Addr.Count, vo), [Is].EqualTo(123))
            End Sub

        End Class

        Public Class 非PublicコンストラクタなVoを出力Test_Debug用なんだからPrivateだろうがなんだろうが出力させたい : Inherits VoPropertyMarkerTest

            Private Class PublicVo
                Public Property Name As String
                Public Property Id As Integer
            End Class
            Private Class PrivateVo : Inherits PublicVo
                Public Shared ReadOnly Instance As New PrivateVo
                Private Sub New()
                End Sub
            End Class
            Private Class FriendVo : Inherits PublicVo
                Friend Sub New()
                End Sub
            End Class

            <Test()> Public Sub Friendコンストラクタでも_マーキングして取得できる()
                Dim marked As FriendVo = marker.CreateMarkedVo(Of FriendVo)()
                Dim vo As New FriendVo With {.Id = 123, .Name = "Ron"}
                Assert.That(marker.GetValue(marked.Id, vo), [Is].EqualTo(123))
                Assert.That(marker.GetValue(marked.Name, vo), [Is].EqualTo("Ron"))
            End Sub

            <Test()> Public Sub Privateコンストラクタでも_マーキングして取得できる()
                Dim marked As PrivateVo = marker.CreateMarkedVo(Of PrivateVo)()
                PrivateVo.Instance.Id = 7
                PrivateVo.Instance.Name = "Alice"
                Assert.That(marker.GetValue(marked.Id, PrivateVo.Instance), [Is].EqualTo(7))
                Assert.That(marker.GetValue(marked.Name, PrivateVo.Instance), [Is].EqualTo("Alice"))
            End Sub

        End Class

    End Class
End Namespace