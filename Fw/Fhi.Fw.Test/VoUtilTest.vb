Imports System
Imports Fhi.Fw.TestUtil.DebugString
Imports Fhi.Fw.Db
Imports NUnit.Framework
Imports Fhi.Fw.Domain

<TestFixture()> Public MustInherit Class VoUtilTest

#Region "テストに使用するVO"
    Protected Class Hoge
        Private _Id As Nullable(Of Integer)
        Private _Name As String
        Private _Fx As Nullable(Of Decimal)
        Private _Adate As DateTime?

        Public Sub New()
        End Sub
        Public Sub New(id As Integer?, name As String)
            Me.Id = id
            Me.Name = name
        End Sub

        Public Property Id() As Nullable(Of Integer)
            Get
                Return _Id
            End Get
            Set(ByVal value As Nullable(Of Integer))
                _Id = value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                _Name = value
            End Set
        End Property
        Public Property Fx() As Nullable(Of Decimal)
            Get
                Return _Fx
            End Get
            Set(ByVal value As Nullable(Of Decimal))
                _Fx = value
            End Set
        End Property
        Public Property Adate() As DateTime?
            Get
                Return _Adate
            End Get
            Set(ByVal value As DateTime?)
                _Adate = value
            End Set
        End Property
    End Class

    Protected Class Fuga
        Private _Id As String
        Private _Name As String
        Private _Fx As String
        Private _Adate As String
        Public Property Id() As String
            Get
                Return _Id
            End Get
            Set(ByVal value As String)
                _Id = value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                _Name = value
            End Set
        End Property
        Public Property Fx() As String
            Get
                Return _Fx
            End Get
            Set(ByVal value As String)
                _Fx = value
            End Set
        End Property
        Public Property Adate() As String
            Get
                Return _Adate
            End Get
            Set(ByVal value As String)
                _Adate = value
            End Set
        End Property
    End Class
    Private Class StrCollection : Inherits CollectionObject(Of String)

        Public Sub New()
        End Sub

        Public Sub New(ByVal src As CollectionObject(Of String))
            MyBase.New(src)
        End Sub

        Public Sub New(ByVal initialList As IEnumerable(Of String))
            MyBase.New(initialList)
        End Sub
    End Class

    Private Class FugaVo
        Private _id As Integer
        Private _hoge As Hoge
        Private _strArray As String()
        Private _strList As List(Of String)
        Public Property StrCollection As StrCollection
        Private _hogeArray As Hoge()
        Private _hogeList As List(Of Hoge)

        Public Property Id() As Integer
            Get
                Return _id
            End Get
            Set(ByVal value As Integer)
                _id = value
            End Set
        End Property

        Public Property Hoge() As Hoge
            Get
                Return _hoge
            End Get
            Set(ByVal value As Hoge)
                _hoge = value
            End Set
        End Property

        Public Property StrArray() As String()
            Get
                Return _strArray
            End Get
            Set(ByVal value As String())
                _strArray = value
            End Set
        End Property

        Public Property StrList() As List(Of String)
            Get
                Return _strList
            End Get
            Set(ByVal value As List(Of String))
                _strList = value
            End Set
        End Property

        Public Property HogeArray() As Hoge()
            Get
                Return _hogeArray
            End Get
            Set(ByVal value As Hoge())
                _hogeArray = value
            End Set
        End Property

        Public Property HogeList() As List(Of Hoge)
            Get
                Return _hogeList
            End Get
            Set(ByVal value As List(Of Hoge))
                _hogeList = value
            End Set
        End Property
    End Class

    Protected Class BooleanVo
        Private _boolA As Boolean?
        Private _boolB As Boolean?

        Public Property BoolA() As Boolean?
            Get
                Return _boolA
            End Get
            Set(ByVal value As Boolean?)
                _boolA = value
            End Set
        End Property

        Public Property BoolB() As Boolean?
            Get
                Return _boolB
            End Get
            Set(ByVal value As Boolean?)
                _boolB = value
            End Set
        End Property
    End Class

    Protected Class AllVo
        Private _aInt As Integer?
        Private _aLong As Long?
        Private _aDouble As Double?
        Private _aDecimal As Decimal?
        Private _aDateTime As DateTime?
        Private _aBoolean As Boolean?
        Private _aString As String
        Private _aObject As Object

        Private _aIntArray As Integer?()
        Private _aLongArray As Long?()
        Private _aDoubleArray As Double?()
        Private _aDecimalArray As Decimal?()
        Private _aDateTimeArray As DateTime?()
        Private _aBooleanArray As Boolean?()
        Private _aStringArray As String()
        Private _aObjectArray As Object()

        Public Property AInt() As Integer?
            Get
                Return _aInt
            End Get
            Set(ByVal value As Integer?)
                _aInt = value
            End Set
        End Property

        Public Property ALong() As Long?
            Get
                Return _aLong
            End Get
            Set(ByVal value As Long?)
                _aLong = value
            End Set
        End Property

        Public Property ADouble() As Double?
            Get
                Return _aDouble
            End Get
            Set(ByVal value As Double?)
                _aDouble = value
            End Set
        End Property

        Public Property ADecimal() As Decimal?
            Get
                Return _aDecimal
            End Get
            Set(ByVal value As Decimal?)
                _aDecimal = value
            End Set
        End Property

        Public Property ADateTime() As Date?
            Get
                Return _aDateTime
            End Get
            Set(ByVal value As Date?)
                _aDateTime = value
            End Set
        End Property

        Public Property ABoolean() As Boolean?
            Get
                Return _aBoolean
            End Get
            Set(ByVal value As Boolean?)
                _aBoolean = value
            End Set
        End Property

        Public Property AString() As String
            Get
                Return _aString
            End Get
            Set(ByVal value As String)
                _aString = value
            End Set
        End Property

        Public Property AObject() As Object
            Get
                Return _aObject
            End Get
            Set(ByVal value As Object)
                _aObject = value
            End Set
        End Property

        Public Property AIntArray() As Integer?()
            Get
                Return _aIntArray
            End Get
            Set(ByVal value As Integer?())
                _aIntArray = value
            End Set
        End Property

        Public Property ALongArray() As Long?()
            Get
                Return _aLongArray
            End Get
            Set(ByVal value As Long?())
                _aLongArray = value
            End Set
        End Property

        Public Property ADoubleArray() As Double?()
            Get
                Return _aDoubleArray
            End Get
            Set(ByVal value As Double?())
                _aDoubleArray = value
            End Set
        End Property

        Public Property ADecimalArray() As Decimal?()
            Get
                Return _aDecimalArray
            End Get
            Set(ByVal value As Decimal?())
                _aDecimalArray = value
            End Set
        End Property

        Public Property ADateTimeArray() As Date?()
            Get
                Return _aDateTimeArray
            End Get
            Set(ByVal value As Date?())
                _aDateTimeArray = value
            End Set
        End Property

        Public Property ABooleanArray() As Boolean?()
            Get
                Return _aBooleanArray
            End Get
            Set(ByVal value As Boolean?())
                _aBooleanArray = value
            End Set
        End Property

        Public Property AStringArray() As String()
            Get
                Return _aStringArray
            End Get
            Set(ByVal value As String())
                _aStringArray = value
            End Set
        End Property

        Public Property AObjectArray() As Object()
            Get
                Return _aObjectArray
            End Get
            Set(ByVal value As Object())
                _aObjectArray = value
            End Set
        End Property
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
        Private _First As EnumA

        Public Property First() As EnumA
            Get
                Return _First
            End Get
            Set(ByVal value As EnumA)
                _First = value
            End Set
        End Property
    End Class
    Protected Class CEnumB
        Private _First As EnumB

        Public Property First() As EnumB
            Get
                Return _First
            End Get
            Set(ByVal value As EnumB)
                _First = value
            End Set
        End Property
    End Class
    Private Class UserA
        Private _A As SuperA()

        Public Property A() As SuperA()
            Get
                Return _A
            End Get
            Set(ByVal value As SuperA())
                _A = value
            End Set
        End Property
    End Class

    Private Class SuperA
        Public SuperA As String
    End Class
    Private Class SubA : Inherits SuperA
        Public SubA As String
    End Class

    Private Class TestingIdAttribute : Inherits Attribute
    End Class
    Private Class TestingNameAttribute : Inherits Attribute
    End Class
    Private Class AttrVo
        <TestingId()>
        Public Property Id As Integer
        <TestingName()>
        Public Property Name As String
    End Class
    Private Class Has3BooleanVo
        Public Property Id As Integer?
        Public Property Name As String
        Public Property IsA As Boolean
        Public Property IsB As Boolean
        Public Property IsC As Boolean
    End Class
    Private Class IntPVO : Inherits PrimitiveValueObject(Of Integer)
        Public Sub New(value As Integer)
            MyBase.New(value)
        End Sub
        Public Function AddOne() As IntPVO
            Return New IntPVO(Value + 1)
        End Function
    End Class
#End Region

    Protected Function NewHoge(ByVal id As Integer?, ByVal name As String, ByVal fx As Decimal?, ByVal aDate As DateTime?) As Hoge
        Dim vo As New Hoge
        vo.Id = id
        vo.Name = name
        vo.Fx = fx
        vo.Adate = aDate
        Return vo
    End Function
    Private Function NewFuga(ByVal id As String, ByVal name As String, ByVal fx As String, ByVal aDate As String) As Fuga
        Dim vo As New Fuga
        vo.Id = id
        vo.Name = name
        vo.Fx = fx
        vo.Adate = aDate
        Return vo
    End Function

    Public Class VoUTil基本テスト : Inherits VoUtilTest

        <Test()> Public Sub Learning_Propertyの型()
            Assert.AreEqual(GetType(Nullable(Of Integer)), GetType(Hoge).GetProperty("Id").PropertyType)
        End Sub

        <Test()> Public Sub Learning_PropertyがNullableかを判断()
            Dim propertyType As Type = GetType(Hoge).GetProperty("Id").PropertyType

            Assert.IsTrue(propertyType.Name.StartsWith("Nullable`"))
            Assert.AreEqual(1, propertyType.GetGenericArguments.Length)
            Assert.AreEqual(GetType(Integer), propertyType.GetGenericArguments(0))
        End Sub

    End Class

    Public Class CopyPropertiesTest : Inherits VoUtilTest

        Private Overloads Function ToString(ParamArray records As Hoge()) As String
            Return New DebugStringMaker(Of Hoge)(Function(define, vo) define.Bind(vo.Id, vo.Name, vo.Fx, vo.Adate)).MakeString(records)
        End Function

        Private Overloads Function ToString(ParamArray records As Fuga()) As String
            Return New DebugStringMaker(Of Fuga)(Function(define, vo) define.Bind(vo.Id, vo.Name, vo.Fx, vo.Adate)).MakeString(records)
        End Function

        <Test()> Public Sub CopyProperties_同じ型同士_同名プロパティが同じ型()
            Dim src As New Hoge With {.Id = 123, .Name = "hogeName", .Fx = 3.141592653589@, .Adate = CDate("2010/01/02 12:34:56")}
            Dim dest As New Hoge
            VoUtil.CopyProperties(src, dest)

            Assert.That(ToString(dest), [Is].EqualTo(
                "Id  Name       Fx      Adate                " & vbCrLf &
                "123 'hogeName' 3.141… '2010/01/02 12:34:56'"))
        End Sub

        <Test()> Public Sub CopyProperties_String型へ()
            Dim src As New Hoge With {.Id = 123, .Name = "hogeName", .Fx = 3.141592653589@, .Adate = CDate("2010/01/02 12:34:56")}
            Dim dest As New Fuga
            VoUtil.CopyProperties(src, dest)

            Assert.That(ToString(dest), [Is].EqualTo(
                "Id    Name       Fx               Adate                " & vbCrLf &
                "'123' 'hogeName' '3.141592653589' '2010/01/02 12:34:56'"))
        End Sub

        <Test()> Public Sub CopyProperties_String型からそれぞれの型へ()
            Dim src As New Fuga With {.Id = "123", .Name = "hogeName", .Fx = "3.141592653589", .Adate = "2010/01/02 12:34:56"}
            Dim dest As New Hoge
            VoUtil.CopyProperties(src, dest)

            Assert.That(ToString(dest), [Is].EqualTo(
                "Id  Name       Fx      Adate                " & vbCrLf &
                "123 'hogeName' 3.141… '2010/01/02 12:34:56'"))
        End Sub

        <Test()> Public Sub CopyProperties_同名だけど違うEnum型はエラーとする()
            Dim src As New CEnumA
            src.First = EnumA.A_FIRST

            Dim dest As New CEnumB
            Try
                VoUtil.CopyProperties(src, dest)
                Assert.Fail("違う型なら例外になるべき")

            Catch expect As ArgumentException
                Assert.AreEqual("EnumA型（値=A_FIRST）を EnumB 型へ変換できない", expect.Message)
            End Try
        End Sub

        <Test()> Public Sub CopyPropertiesWithIgnoreProperties_ラムダ式でコピーしたくないプロパティを指定できる()
            Dim src As New Hoge With {.Id = 123, .Name = "hogeName", .Fx = 3.141592653589@, .Adate = CDate("2010/01/02 12:34:56")}
            Dim dest As New Hoge

            Dim configure As VoUtil.PropertySpecifierConfigure(Of Hoge) = Function(define, vo) define.AppendProperties(vo.Adate)
            VoUtil.CopyPropertiesWithIgnoreProperties(Of Hoge)(src, dest, configure)

            Assert.That(ToString(dest), [Is].EqualTo(
                "Id  Name       Fx      Adate" & vbCrLf &
                "123 'hogeName' 3.141… null "))
        End Sub

        <Test()> Public Sub CopyPropertiesWithIgnoreProperties_ラムダ式でコピーしたくないプロパティを指定できる_型違いでもOK()
            Dim src As New Hoge With {.Id = 123, .Name = "hogeName", .Fx = 3.141592653589@, .Adate = CDate("2010/01/02 12:34:56")}
            Dim dest As New Fuga

            Dim configure As VoUtil.PropertySpecifierConfigure(Of Hoge) = Function(define, vo) define.AppendProperties(vo.Adate).AppendProperties(vo.Fx)
            VoUtil.CopyPropertiesWithIgnoreProperties(Of Hoge)(src, dest, configure)

            Assert.That(ToString(dest), [Is].EqualTo(
                "Id    Name       Fx   Adate" & vbCrLf &
                "'123' 'hogeName' null null "))
        End Sub

    End Class

    Public Class IsEqualsTest : Inherits VoUtilTest

        <Test()> Public Sub IsEquals_型違いはfalse()
            Assert.IsFalse(VoUtil.IsEquals(New Hoge, New Fuga))
        End Sub

        <Test()> Public Sub IsEquals_同じ型で違う値ならfalse()
            Assert.IsFalse(VoUtil.IsEquals(NewFuga("a", "bb", "ccc", "d"), NewFuga("a", "bb", "333", "d")))
            Assert.IsFalse(VoUtil.IsEquals(NewFuga("a", "bb", "ccc", "d"), NewFuga("a", "bb", Nothing, "d")))
            Assert.IsFalse(VoUtil.IsEquals(NewFuga("a", "bb", Nothing, "d"), NewFuga("a", "bb", "333", "d")))
        End Sub

        <Test()> Public Sub IsEquals_同じ型で同じ値ならtrue()
            Assert.IsTrue(VoUtil.IsEquals(NewFuga("a", "bb", "ccc", "d"), NewFuga("a", "bb", "ccc", "d")))
        End Sub

        <Test()> Public Sub IsEquals_同じ型で同じ値ならtrue_Nullable()
            Assert.IsTrue(VoUtil.IsEquals(NewHoge(12, "bb", 3.14D, CDate("11:22:33")), NewHoge(12, "bb", 3.14D, CDate("11:22:33"))))
        End Sub

        <Test()> Public Sub IsEquals_コレクションも比較_true()
            Assert.IsTrue(VoUtil.IsEquals(New String() {"a", "b"}, New String() {"a", "b"}), "配列の比較")
            Assert.IsTrue(VoUtil.IsEquals(EzUtil.NewList("c", "d", "e"), EzUtil.NewList("c", "d", "e")), "Listの比較")
        End Sub

        <Test()> Public Sub IsEquals_コレクションも比較_false()
            Assert.IsFalse(VoUtil.IsEquals(New String() {"a", "b"}, New String() {"a", "c"}), "配列の比較")
            Assert.IsFalse(VoUtil.IsEquals(EzUtil.NewList("c", "d", "e"), EzUtil.NewList("d", "e")), "Listの比較")
            Assert.IsFalse(VoUtil.IsEquals(EzUtil.NewList("c", "d", "e"), New String() {"a", "b"}), "配列・Listの比較")
        End Sub

        <Test()> Public Sub IsEquals_Voの中のコレクションの中身も比較_true()
            Assert.IsTrue(VoUtil.IsEquals(New FugaVo With {.StrArray = New String() {"a", "b"}}, New FugaVo With {.StrArray = New String() {"a", "b"}}), "配列の比較")
            Assert.IsTrue(VoUtil.IsEquals(New FugaVo With {.StrList = EzUtil.NewList("c", "d", "e")}, New FugaVo With {.StrList = EzUtil.NewList("c", "d", "e")}), "Listの比較")
        End Sub

        <Test()> Public Sub IsEquals_Voの中のコレクションの中身も比較_false()
            Assert.IsFalse(VoUtil.IsEquals(New FugaVo With {.StrArray = New String() {"a", "b"}}, New FugaVo With {.StrArray = New String() {"a", "c"}}), "配列の比較")
            Assert.IsFalse(VoUtil.IsEquals(New FugaVo With {.StrList = EzUtil.NewList("c", "d", "e")}, New FugaVo With {.StrList = EzUtil.NewList("d", "e")}), "Listの比較")
            Assert.IsFalse(VoUtil.IsEquals(New FugaVo With {.StrList = EzUtil.NewList("c", "d", "e")}, New FugaVo With {.StrArray = New String() {"a", "b"}}), "配列・Listの比較")
        End Sub

        <Test()> Public Sub IDプロパティが違うのでfalseになる()
            Assert.That(VoUtil.IsEquals(New AttrVo With {.Id = 19, .Name = "A"},
                                        New AttrVo With {.Id = 20, .Name = "A"}), [Is].False)
        End Sub

        <Test()> Public Sub IDプロパティが違うのでfalseになる_がID属性を無視すれば同値になる()
            Assert.That(VoUtil.IsEquals(New AttrVo With {.Id = 19, .Name = "A"},
                                        New AttrVo With {.Id = 20, .Name = "A"}, GetType(TestingIdAttribute)), [Is].True)
        End Sub

        <Test()> Public Sub Nameプロパティが違うのでfalseになる()
            Assert.That(VoUtil.IsEquals(New AttrVo With {.Id = 5, .Name = "A"},
                                        New AttrVo With {.Id = 5, .Name = "B"}), [Is].False)
        End Sub

        <Test()> Public Sub Nameプロパティが違うのでfalseになる_がName属性を無視すれば同値になる()
            Assert.That(VoUtil.IsEquals(New AttrVo With {.Id = 5, .Name = "A"},
                                        New AttrVo With {.Id = 5, .Name = "B"}, GetType(TestingNameAttribute)), [Is].True)
        End Sub

    End Class

    Public Class GetValueTest : Inherits VoUtilTest

        <Test()> Public Sub GetValue_dot経由でプロパティ値のプロパティ値を取得()
            Dim vo As New FugaVo
            vo.Hoge = New Hoge
            vo.Hoge.Name = "a"
            Assert.AreEqual("a", VoUtil.GetValue(vo, "Hoge.Name"))
        End Sub

        <Test()> Public Sub GetValue_括弧で配列の要素を取得()
            Dim vo As New FugaVo
            vo.StrArray = New String() {"one", "two"}
            Assert.AreEqual("one", VoUtil.GetValue(vo, "StrArray(0)"))
            Assert.AreEqual("two", VoUtil.GetValue(vo, "StrArray(1)"))
        End Sub

        <Test()> Public Sub GetValue_配列値の要素数はLengthで参照()
            Dim vo As New FugaVo
            vo.StrArray = New String() {"one", "two"}
            Assert.AreEqual(2, VoUtil.GetValue(vo, "StrArray.Length"))
        End Sub

        <Test()> Public Sub GetValue_括弧で配列の要素を取得してdot経由で更にそのプロパティ値()
            Dim vo As New FugaVo
            vo.HogeArray = New Hoge() {New Hoge}
            vo.HogeArray(0).Name = "fugahoge"
            Assert.AreEqual("fugahoge", VoUtil.GetValue(vo, "HogeArray(0).Name"))
        End Sub

        <Test()> Public Sub GetValue_括弧でListの要素を取得()
            Dim vo As New FugaVo
            vo.StrList = New List(Of String)(New String() {"x", "y"})
            Assert.AreEqual("x", VoUtil.GetValue(vo, "StrList(0)"))
            Assert.AreEqual("y", VoUtil.GetValue(vo, "StrList(1)"))
        End Sub

        <Test()> Public Sub GetValue_List値の要素数はCountで参照()
            Dim vo As New FugaVo
            vo.StrList = New List(Of String)(New String() {"x", "y", "z"})
            Assert.AreEqual(3, VoUtil.GetValue(vo, "StrList.Count"))
        End Sub

        <Test()> Public Sub GetValue_括弧でListの要素を取得してdot経由で更にそのプロパティ値()
            Dim vo As New FugaVo
            vo.HogeList = New List(Of Hoge)(New Hoge() {New Hoge})
            vo.HogeList(0).Name = "fugahoge"
            Assert.AreEqual("fugahoge", VoUtil.GetValue(vo, "HogeList(0).Name"))
        End Sub

        <Test()> Public Sub GetValue_配列値の場合はValue括弧で値を参照()
            Dim param As String() = New String() {"a", "b", "c"}
            Assert.AreEqual("b", VoUtil.GetValue(param, "Value(1)"))
        End Sub

        <Test()> Public Sub GetValue_配列Vo値の場合はValue括弧とdot経由で値を参照()
            Dim param As Hoge() = New Hoge() {New Hoge, New Hoge}
            param(0).Name = "x"
            param(1).Name = "y"
            Assert.AreEqual("y", VoUtil.GetValue(param, "Value(1).Name"))
        End Sub

        <Test()> Public Sub GetValue_CollectionObject値の場合はValue括弧で値を参照できる()
            Dim vo As New FugaVo
            vo.StrCollection = New StrCollection({"one", "two"})
            Assert.That(VoUtil.GetValue(vo, "StrCollection(0)"), [Is].EqualTo("one"))
            Assert.That(VoUtil.GetValue(vo, "StrCollection(1)"), [Is].EqualTo("two"))
        End Sub

    End Class

    Public Class NewElementInstanceFromCollectionType_コレクション型から要素のインスタンスを作成する : Inherits VoUtilTest

        <Test()> Public Sub ListOfVoからならVo作成()
            Dim actual As Object = VoUtil.NewElementInstanceFromCollectionType(GetType(List(Of Hoge)))
            Assert.IsNotNull(actual)
            Assert.AreEqual(GetType(Hoge), actual.GetType)
        End Sub

        <Test()> Public Sub ArrayOfVoからならVo作成()
            Dim actual As Object = VoUtil.NewElementInstanceFromCollectionType(GetType(Hoge).MakeArrayType)
            Assert.IsNotNull(actual)
            Assert.AreEqual(GetType(Hoge), actual.GetType)
        End Sub

        <Test()> Public Sub ListOfStringからならNull_Stringは作成しない()
            Dim actual As Object = VoUtil.NewElementInstanceFromCollectionType(GetType(List(Of String)))
            Assert.IsNull(actual, "Stringは作成しない")
        End Sub

        <Test()> Public Sub ListOfIntからなら初期値0()
            Dim actual As Object = VoUtil.NewElementInstanceFromCollectionType(GetType(List(Of Integer)))
            Assert.AreEqual(GetType(Integer), actual.GetType)
            Assert.AreEqual(0, actual)
        End Sub

        <Test()> Public Sub ListOfNullableIntからならNull()
            Dim actual As Object = VoUtil.NewElementInstanceFromCollectionType(GetType(List(Of Integer?)))
            Assert.IsNull(actual, "Nullable型も作成しない")
        End Sub

    End Class

    Public Class NewInstanceTest : Inherits VoUtilTest

        <Test()> Public Sub シャローコピーしたインスタンスを作成する()
            Dim obj As New Object
            Dim vo As New AllVo With {.AInt = 1, .ALong = 2L, .ADouble = 3.0R, .ADecimal = 4D, .ADateTime = #5/5/2005#, .ABoolean = True, .AString = "7", .AObject = obj}
            Dim actual As AllVo = VoUtil.NewInstance(vo)
            Assert.That(actual.AInt, [Is].EqualTo(1))
            Assert.That(actual.ALong, [Is].EqualTo(2L))
            Assert.That(actual.ADouble, [Is].EqualTo(3.0R))
            Assert.That(actual.ADecimal, [Is].EqualTo(4D))
            Assert.That(actual.ADateTime, [Is].EqualTo(#5/5/2005#))
            Assert.That(actual.ABoolean, [Is].EqualTo(True))
            Assert.That(actual.AString, [Is].EqualTo("7"))
            Assert.That(actual.AObject, [Is].SameAs(obj), "シャローコピーだからプロパティのインスタンスは同じ")
            Assert.That(actual, [Is].Not.SameAs(vo), "親のインスタンスは別物")
        End Sub

        <Test()> Public Sub 配列を作成できる()
            Dim propertiesObj As String() = {"aaa", "123"}
            Dim actuals As String() = VoUtil.NewInstance(Of String())(propertiesObj)
            Assert.That(actuals, [Is].EquivalentTo(New String() {"aaa", "123"}))
            Assert.That(actuals, [Is].Not.SameAs(propertiesObj), "インスタンスは別物")
        End Sub

        <Test()> Public Sub 配列を作成できる_引数が無ければ_長さ0で作成できる()
            Dim actuals As String() = VoUtil.NewInstance(Of String())()
            Assert.That(actuals, [Is].EquivalentTo(New String() {}))
        End Sub

        <Test()> Public Sub nullableな_Integer配列も作成できる()
            Dim vo As New AllVo With {.AIntArray = New Integer?() {1, Nothing, 5}}
            Dim actual As AllVo = VoUtil.NewInstance(vo)
            Assert.That(actual.AIntArray, [Is].EquivalentTo(New Integer?() {1, Nothing, 5}))
        End Sub

        <Test()> Public Sub nullableな_Decimal配列も作成できる()
            Dim vo As New AllVo With {.ADecimalArray = New Decimal?() {10.5D, 20D, Nothing, -3.14159D}}
            Dim actual As AllVo = VoUtil.NewInstance(vo)
            Assert.That(actual.ADecimalArray, [Is].EquivalentTo(New Decimal?() {10.5D, 20D, Nothing, -3.14159D}))
        End Sub

        <Test()> Public Sub nullableな_Datetime配列も作成できる()
            Dim vo As New AllVo With {.ADateTimeArray = New DateTime?() {#1/1/2012#, Nothing}}
            Dim actual As AllVo = VoUtil.NewInstance(vo)
            Assert.That(actual.ADateTimeArray, [Is].EquivalentTo(New DateTime?() {#1/1/2012#, Nothing}))
        End Sub

        <Test()> Public Sub nullableな_Boolean配列も作成できる()
            Dim vo As New AllVo With {.ABooleanArray = New Boolean?() {True, Nothing, False}}
            Dim actual As AllVo = VoUtil.NewInstance(vo)
            Assert.That(actual.ABooleanArray, [Is].EquivalentTo(New Boolean?() {True, Nothing, False}))
        End Sub

        <Test()> Public Sub 長さ0の配列も作成できる()
            Dim vo As New AllVo With {.AStringArray = New String() {}}
            Dim actual As AllVo = VoUtil.NewInstance(vo)
            Assert.That(actual.AStringArray, [Is].Not.Null.And.Empty)
        End Sub

        <Test()> Public Sub 配列要素のサブクラスは_シャローコピーする()
            Dim vo As New UserA With {.A = New SubA() {New SubA(), New SubA() With {.SuperA = "SuperA", .SubA = "SubA"}}}
            Dim actual As UserA = VoUtil.NewInstance(vo)
            Assert.That(actual.A(1).SuperA, [Is].EqualTo("SuperA"))
            Assert.That(actual.A, [Is].AssignableFrom(GetType(SubA())))
            Assert.That(DirectCast(actual.A(1), SubA).SubA, [Is].EqualTo("SubA"))
        End Sub

    End Class

    Public Class NewInstanceByCollectionTypeTest : Inherits VoUtilTest
        <Test()> Public Sub NewInstanceByCollectionType_List型のVo()
            Dim actual As Object = VoUtil.NewCollectionInstance(GetType(List(Of Hoge)), 3)
            Assert.AreEqual(GetType(List(Of Hoge)), actual.GetType, "List型のVo")

            Dim actuals As List(Of Hoge) = DirectCast(actual, List(Of Hoge))
            Assert.AreEqual(3, actuals.Count)
            Assert.IsNotNull(actuals(0))
            Assert.IsNotNull(actuals(1))
            Assert.IsNotNull(actuals(2))
        End Sub

        <Test()> Public Sub NewInstanceByCollectionType_配列型のVo()
            Dim actual As Object = VoUtil.NewCollectionInstance(GetType(Hoge).MakeArrayType, 2)
            Assert.AreEqual(GetType(Hoge).MakeArrayType, actual.GetType, "配列型のVo")

            Dim actuals As Hoge() = DirectCast(actual, Hoge())
            Assert.AreEqual(2, actuals.Length)
            Assert.IsNotNull(actuals(0))
            Assert.IsNotNull(actuals(1))
        End Sub

        <Test()> Public Sub NewInstanceByCollectionType_List型のInt()
            Dim actual As Object = VoUtil.NewCollectionInstance(GetType(List(Of Integer)), 2)
            Assert.AreEqual(GetType(List(Of Integer)), actual.GetType, "List型のInt")

            Dim actuals As List(Of Integer) = DirectCast(actual, List(Of Integer))
            Assert.AreEqual(2, actuals.Count)
            Assert.AreEqual(0, actuals(0))
            Assert.AreEqual(0, actuals(1))
        End Sub

        <Test()> Public Sub NewInstanceByCollectionType_配列型のNullableInt()
            Dim actual As Object = VoUtil.NewCollectionInstance(GetType(Integer?).MakeArrayType, 3)
            Assert.AreEqual(GetType(Integer?).MakeArrayType, actual.GetType, "配列型のNullableInt")

            Dim actuals As Integer?() = DirectCast(actual, Integer?())
            Assert.AreEqual(3, actuals.Length)
            Assert.AreEqual(Nothing, actuals(0))
            Assert.AreEqual(Nothing, actuals(1))
            Assert.AreEqual(Nothing, actuals(2))
        End Sub

        <Test()> Public Sub NewInstanceByCollectionType_List型のString()
            Dim actual As Object = VoUtil.NewCollectionInstance(GetType(List(Of String)), 2)
            Assert.AreEqual(GetType(List(Of String)), actual.GetType, "List型のstring")

            Dim actuals As List(Of String) = DirectCast(actual, List(Of String))
            Assert.AreEqual(2, actuals.Count)
            Assert.AreEqual(Nothing, actuals(0))
            Assert.AreEqual(Nothing, actuals(1))
        End Sub

        <Test()> Public Sub NewInstanceByCollectionType_List型のString_初期値を渡す()
            Dim actual As Object = VoUtil.NewCollectionInstanceWithInitialElement(GetType(List(Of String)), 2, New String() {"A"})
            Assert.AreEqual(GetType(List(Of String)), actual.GetType, "List型のstring")

            Dim actuals As List(Of String) = DirectCast(actual, List(Of String))
            Assert.AreEqual(2, actuals.Count)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual(Nothing, actuals(1))
        End Sub
    End Class

    Public Class NewArrayInstanceTest : Inherits VoUtilTest
        <Test()> Public Sub NewArrayInstance_長さゼロの配列を作成できる()
            Assert.AreEqual(0, VoUtil.NewArrayInstance(Of Integer)(0).Length)
            Assert.AreEqual(0, VoUtil.NewArrayInstance(Of String)(0).Length)
            Assert.AreEqual(0, VoUtil.NewArrayInstance(Of Object)(0).Length)
        End Sub

        <Test()> Public Sub NewArrayInstance_初期値は以下の通り()
            Dim actualsInt As Integer() = VoUtil.NewArrayInstance(Of Integer)(2)
            Assert.AreEqual(0, actualsInt(1), "Premitiveの初期値は0")

            Dim actualsStr As String() = VoUtil.NewArrayInstance(Of String)(2)
            Assert.IsNull(actualsStr(0), "文字列の初期値はNull")
        End Sub
    End Class

    Public Class ResolveValueTest : Inherits VoUtilTest
        <Test()> Public Sub ResolveValue_DateTime()
            Assert.AreEqual(CDate("2011/02/03 11:22:33"), VoUtil.ResolveValue(CDate("2011/02/03 11:22:33"), GetType(DateTime)), "日時")
            Assert.AreEqual(CDate("2011/02/03 11:22:33"), VoUtil.ResolveValue("2011/02/03 11:22:33", GetType(DateTime)), "日時")
            Assert.AreEqual(CDate("2011/02/03 11:22:33.444"), VoUtil.ResolveValue("2011/02/03 11:22:33.444", GetType(DateTime)), "ミリ秒含む日時")
            Assert.AreEqual(CDate("11:22:33"), VoUtil.ResolveValue("11:22:33", GetType(DateTime)), "時刻のみ")
            Assert.AreEqual(CDate("2011/02/03"), VoUtil.ResolveValue("2011/02/03", GetType(DateTime)), "日付のみ")
            Assert.IsNull(VoUtil.ResolveValue("aaaaaaaaa", GetType(DateTime)), "不正な値はNothing")
        End Sub

        <Test()> Public Sub ResolveValue_Decimal()
            Assert.AreEqual(3.14159D, VoUtil.ResolveValue(3.14159D, GetType(Decimal)))
            Assert.AreEqual(3.14159D, VoUtil.ResolveValue("3.14159", GetType(Decimal)))
            Assert.AreEqual(12345D, VoUtil.ResolveValue("12345", GetType(Decimal)))
            Assert.IsNull(VoUtil.ResolveValue("aaaaaaaa", GetType(Decimal)), "不正な値はNothing")
        End Sub

        <Test()> Public Sub ResolveValue_Integer()
            Assert.AreEqual(12345, VoUtil.ResolveValue(12345, GetType(Integer)))
            Assert.AreEqual(12345, VoUtil.ResolveValue("12345", GetType(Integer)))
            Assert.IsNull(VoUtil.ResolveValue("3.14", GetType(Integer)), "不正な値(少数値)はNothing")
            Assert.IsNull(VoUtil.ResolveValue("9999999999", GetType(Integer)), "不正な値(Integer超値)はNothing")
            Assert.IsNull(VoUtil.ResolveValue("aaaaaaaa", GetType(Integer)), "不正な値(文字列)はNothing")
        End Sub

        <Test()> Public Sub ResolveValue_Long()
            Assert.AreEqual(12345L, VoUtil.ResolveValue(12345L, GetType(Long)))
            Assert.AreEqual(12345L, VoUtil.ResolveValue("12345", GetType(Long)))
            Assert.AreEqual(9999999999L, VoUtil.ResolveValue("9999999999", GetType(Long)))
            Assert.IsNull(VoUtil.ResolveValue("3.14", GetType(Long)), "不正な値(少数値)はNothing")
            Assert.IsNull(VoUtil.ResolveValue("aaaaaaaa", GetType(Long)), "不正な値(文字列)はNothing")
        End Sub

        <Test()> Public Sub ResolveValue_Double()
            Assert.AreEqual(12345.6R, VoUtil.ResolveValue(12345.6R, GetType(Double)))
            Assert.AreEqual(12345.7R, VoUtil.ResolveValue("12345.7", GetType(Double)))
            Assert.AreEqual(9999999999.0R, VoUtil.ResolveValue("9999999999", GetType(Double)))
            Assert.AreEqual(3.14R, VoUtil.ResolveValue("3.14", GetType(Double)))
            Assert.AreEqual(3.14R, VoUtil.ResolveValue(3.14D, GetType(Double)), "Double値も")
            Assert.IsNull(VoUtil.ResolveValue("aaaaaaaa", GetType(Double)), "不正な値(文字列)はNothing")
        End Sub

        <Test()> Public Sub ResolveValue_Single()
            Assert.AreEqual(12345.6F, VoUtil.ResolveValue(12345.6F, GetType(Single)))
            Assert.AreEqual(12345.7F, VoUtil.ResolveValue("12345.7", GetType(Single)))
            Assert.AreEqual(3.14F, VoUtil.ResolveValue("3.14", GetType(Single)))
            Assert.AreEqual(3.14F, VoUtil.ResolveValue(3.14D, GetType(Single)), "Double値も")
            Assert.AreEqual(1.0E+10F, VoUtil.ResolveValue("9999999999", GetType(Single)), "値は正確じゃないけど変換される")
            Assert.IsNull(VoUtil.ResolveValue("aaaaaaaa", GetType(Single)), "不正な値(文字列)はNothing")
        End Sub

        <Test()> Public Sub ResolveValue_String()
            Assert.AreEqual("12345", VoUtil.ResolveValue(12345, GetType(String)))
            Assert.AreEqual("3.14", VoUtil.ResolveValue(3.14R, GetType(String)))
            Assert.AreEqual("3.141592653589", VoUtil.ResolveValue(3.141592653589D, GetType(String)), "Double値も")
            Assert.AreEqual("2012/01/02 3:04:05", VoUtil.ResolveValue(CDate("2012/01/02 03:04:05"), GetType(String)), "3時の頭0が除外される点に注意")
            Assert.AreEqual("2012/01/02 3:04:05", VoUtil.ResolveValue(CDate("2012/01/02 03:04:05.678"), GetType(String)), "ミリ秒は出力されない")
            Assert.IsNull(VoUtil.ResolveValue(Nothing, GetType(String)), "NothingはNothing")
        End Sub

        <Test()> Public Sub ResolveValue_Enum()
            Assert.AreEqual(SampleEnum.A, VoUtil.ResolveValue(SampleEnum.A, GetType(SampleEnum)))
            Assert.AreEqual(SampleEnum.B, VoUtil.ResolveValue("B", GetType(SampleEnum)))
            Assert.AreEqual(1, CInt(VoUtil.ResolveValue(SampleEnum.A, GetType(SampleEnum))))
            Assert.AreEqual(2, CInt(VoUtil.ResolveValue("B", GetType(SampleEnum))))
        End Sub

        <Test()> Public Sub ResolveValue_Byte()
            Assert.AreEqual(1, VoUtil.ResolveValue(1, GetType(Byte)))
            Assert.AreEqual(255, VoUtil.ResolveValue("255", GetType(Byte)))
            Assert.IsNull(VoUtil.ResolveValue("256", GetType(Byte)), "限界")
            Assert.IsNull(VoUtil.ResolveValue("aaaaaaaa", GetType(Byte)), "不正な値(文字列)はNothing")
        End Sub

        <Test()> Public Sub 同じ配列型にするなら_インスタンスもそのまま_Int()
            Dim values As Integer() = {1, 2}
            Dim actual As Object = VoUtil.ResolveValue(values, GetType(Integer()))
            Assert.That(actual, [Is].EquivalentTo(New Long() {1, 2}))
            Assert.That(actual, [Is].AssignableTo(GetType(Integer())))
            Assert.That(actual, [Is].SameAs(values), "インスタンスも同じ")
        End Sub

        <Test()> Public Sub 同じ配列型にするなら_インスタンスもそのまま_String()
            Dim values As String() = {"aaa", "123"}
            Dim actual As Object = VoUtil.ResolveValue(values, GetType(String()))
            Assert.That(actual, [Is].EquivalentTo(New String() {"aaa", "123"}))
            Assert.That(actual, [Is].AssignableTo(GetType(String())))
            Assert.That(actual, [Is].SameAs(values), "インスタンスも同じ")
        End Sub

        <Test()> Public Sub 同じ配列型にするなら_インスタンスもそのまま_Byte()
            Dim values As Byte() = {1, 2}
            Dim actual As Object = VoUtil.ResolveValue(values, GetType(Byte()))
            Assert.That(actual, [Is].EquivalentTo(New Long() {1, 2}))
            Assert.That(actual, [Is].AssignableTo(GetType(Byte())))
            Assert.That(actual, [Is].SameAs(values), "インスタンスも同じ")
        End Sub

        <Test()> Public Sub 配列型が違うなら_インスタンスは変わる_Int配列をLong配列にできる()
            Dim values As Integer() = {1, 2}
            Dim actual As Object = VoUtil.ResolveValue(values, GetType(Long()))
            Assert.That(actual, [Is].EquivalentTo(New Long() {1, 2}))
            Assert.That(actual.GetType, [Is].SameAs(GetType(Long())))
            Assert.That(actual, [Is].Not.SameAs(values), "型違いの配列だからインスタンスも違う")
        End Sub

        <Test()> Public Sub 配列型が違うなら_インスタンスは変わる_NullableDecimal配列をNullableDouble配列にできる()
            Dim values As Decimal?() = {3.14D, -123D}
            Dim actual As Object = VoUtil.ResolveValue(values, GetType(Double?()))
            Assert.That(actual, [Is].EquivalentTo(New Double?() {3.14R, -123.0R}))
            Assert.That(actual.GetType, [Is].SameAs(GetType(Double?())))
            Assert.That(actual, [Is].Not.SameAs(values), "型違いの配列だからインスタンスも違う")
        End Sub

        <Test()> Public Sub 配列型が違うなら_インスタンスは変わる_Byte配列をLong配列にできる()
            Dim values As Byte() = {1, 2}
            Dim actual As Object = VoUtil.ResolveValue(values, GetType(Long()))
            Assert.That(actual, [Is].EquivalentTo(New Long() {1, 2}))
            Assert.That(actual.GetType, [Is].SameAs(GetType(Long())))
            Assert.That(actual, [Is].Not.SameAs(values), "型違いの配列だからインスタンスも違う")
        End Sub

    End Class

    Public Class InvokeExpressionTest_Lambdaで指定した式を評価する : Inherits VoUtilTest

        <Test()> Public Sub プロパティInt型を評価できる()
            Dim vo As New Hoge
            Assert.That(VoUtil.InvokeExpressionBy(New Hoge With {.Id = 123}, Function() vo.Id), [Is].EqualTo(123))
        End Sub

        <Test()> Public Sub プロパティString型を評価できる()
            Dim vo As New Hoge
            Assert.That(VoUtil.InvokeExpressionBy(New Hoge With {.Name = "ABC"}, Function() vo.Name), [Is].EqualTo("ABC"))
        End Sub

        <Test()> Public Sub プロパティObject型のプロパティを評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo With {.Hoge = New Hoge With {.Name = "Z"}}, _
                                                  Function() vo.Hoge.Name), [Is].EqualTo("Z"))
        End Sub

        <Test()> Public Sub プロパティString配列型添え字を評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo With {.StrArray = New String() {"i0", "i1"}}, _
                                                  Function() vo.StrArray(1)), [Is].EqualTo("i1"))
        End Sub

        <Test()> Public Sub プロパティString配列のプロパティを評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo With {.StrArray = New String() {"i0", "i1"}}, _
                                                  Function() vo.StrArray.Length), [Is].EqualTo(2))
        End Sub

        <Test()> Public Sub プロパティString配列引数のメソッドを評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo With {.StrArray = New String() {"i0", "i1"}}, _
                                                  Function() Join(vo.StrArray, ",")), [Is].EqualTo("i0,i1"))
        End Sub

        <Test()> Public Sub プロパティStringList型添え字を評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo With {.StrList = (New String() {"i0", "i1"}).ToList}, _
                                                  Function() vo.StrList(1)), [Is].EqualTo("i1"))
        End Sub

        <Test()> Public Sub プロパティStringList型のプロパティを評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo With {.StrList = (New String() {"i0", "i1"}).ToList}, _
                                                  Function() vo.StrList.Count), [Is].EqualTo(2))
        End Sub

        <Test()> Public Sub プロパティStringList引数のメソッドを評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo With {.StrList = (New String() {"i0", "i1"}).ToList}, _
                                                  Function() Join(vo.StrList.ToArray, ",")), [Is].EqualTo("i0,i1"))
        End Sub

        <Test()> Public Sub プロパティString型引数_型変換CIntを評価できる()
            Dim vo As New Fuga
            Assert.That(VoUtil.InvokeExpressionBy(New Fuga With {.Id = "12"}, _
                                                  Function() CInt(vo.Id)), [Is].EqualTo(12))
        End Sub

        <Test()> Public Sub プロパティObject配列型添え字のプロパティを評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo With {.HogeArray = New Hoge() {New Hoge With {.Name = "n0"}, _
                                                                                            New Hoge With {.Name = "n1"}}}, _
                                                  Function() vo.HogeArray(1).Name), [Is].EqualTo("n1"))
        End Sub

        Private Shared ReadOnly SharedReadonlyPiyo As New Hoge With {.Name = "SharedReadonly"}
        <Test()> Public Sub 第一引数のインスタンスを利用しない_Shared変数を参照できる()
            Dim actual As Object = VoUtil.InvokeExpressionBy(New Fuga, Function() SharedReadonlyPiyo.Name)
            Assert.That(actual, [Is].EqualTo("SharedReadonly"))
        End Sub

        <Test()> Public Sub 第一引数のインスタンスを利用しない_New演算子を使用できる()
            Dim actual As Object = VoUtil.InvokeExpressionBy(New Fuga, Function() New Hoge)
            Assert.That(actual, [Is].TypeOf(Of Hoge))
        End Sub

        <Test()> Public Sub 第一引数のインスタンスを利用しない_New演算子を使用できる_コンストラクタ引数に定数()
            Dim actual As Object = VoUtil.InvokeExpressionBy(New Fuga, Function() New Hoge(2, "f"))
            Assert.That(actual, [Is].TypeOf(Of Hoge))
            Assert.That(DirectCast(actual, Hoge).Name, [Is].EqualTo("f"))
        End Sub

        <Test()> Public Sub 第一引数のインスタンスを利用しない_New演算子を使用できる_With句に定数()
            Dim actual As Object = VoUtil.InvokeExpressionBy(New Fuga, Function() New Hoge With {.Name = "g"})
            Assert.That(actual, [Is].TypeOf(Of Hoge))
            Assert.That(DirectCast(actual, Hoge).Name, [Is].EqualTo("g"))
        End Sub

        <Test()> Public Sub New演算子_コンストラクタで_代入できる()
            Dim vo As New Fuga
            Dim actual As Object = VoUtil.InvokeExpressionBy(New Fuga With {.Id = "3", .Name = "jan"},
                                                             Function() New Hoge(CInt(vo.Id), vo.Name))
            Assert.That(actual, [Is].TypeOf(Of Hoge))
            Assert.That(DirectCast(actual, Hoge).Id, [Is].EqualTo(3))
            Assert.That(DirectCast(actual, Hoge).Name, [Is].EqualTo("jan"))
        End Sub

        <Test()> Public Sub New演算子_With句で_代入できる()
            Dim vo As New Fuga
            Dim actual As Object = VoUtil.InvokeExpressionBy(New Fuga With {.Name = "abc"},
                                                             Function() New Hoge With {.Name = vo.Name})
            Assert.That(actual, [Is].TypeOf(Of Hoge))
            Assert.That(DirectCast(actual, Hoge).Name, [Is].EqualTo("abc"))
        End Sub

        <Test()> Public Sub is演算子を評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo, _
                                                  Function() If(vo.StrArray Is Nothing, "foo", vo.StrArray(1))), [Is].EqualTo("foo"))
        End Sub

        <Test()> Public Sub isNot演算子を評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo, _
                                                  Function() If(vo.StrArray IsNot Nothing, vo.StrArray(1), "bar")), [Is].EqualTo("bar"))
        End Sub

        <Test()> Public Sub LT演算子を評価できる()
            Dim vo As New Hoge
            Assert.That(VoUtil.InvokeExpressionBy(New Hoge With {.Id = 0}, _
                                                  Function() If(vo.Id < 0, "true", "false")), [Is].EqualTo("false"))
        End Sub

        <Test()> Public Sub LE演算子を評価できる()
            Dim vo As New Hoge
            Assert.That(VoUtil.InvokeExpressionBy(New Hoge With {.Id = 0}, _
                                                  Function() If(vo.Id <= 0, "true", "false")), [Is].EqualTo("true"))
        End Sub

        <Test()> Public Sub GT演算子を評価できる()
            Dim vo As New Hoge
            Assert.That(VoUtil.InvokeExpressionBy(New Hoge With {.Id = 0}, _
                                                  Function() If(vo.Id > 0, "true", "false")), [Is].EqualTo("false"))
        End Sub

        <Test()> Public Sub GE演算子を評価できる()
            Dim vo As New Hoge
            Assert.That(VoUtil.InvokeExpressionBy(New Hoge With {.Id = 0}, _
                                                  Function() If(vo.Id >= 0, "true", "false")), [Is].EqualTo("true"))
        End Sub

        <Test()> Public Sub iif関数を評価できる()
            Dim vo As New FugaVo
            Assert.That(VoUtil.InvokeExpressionBy(New FugaVo, _
                                                  Function() IIf(vo.StrArray Is Nothing, "true", "false")), [Is].EqualTo("true"))
        End Sub

        <Test()> Public Sub 内部のLambda指定は現在未対応()
            Dim vo As New FugaVo
            Try
                VoUtil.InvokeExpressionBy(New FugaVo With {.HogeArray = New Hoge() {New Hoge With {.Name = "n0"}, _
                                                                                    New Hoge With {.Name = "n1"}}}, _
                                          Function() Join(vo.HogeArray.Select(Function(vo2) vo2.Name).ToArray, ","))
            Catch ex As NotSupportedException
                Assert.That(ex.Message, [Is].EqualTo("内部のLambdaは未対応"))
            End Try
        End Sub

        <Test()> Public Sub 値の差替_Integer値_乗算()
            Dim value As Integer
            Assert.That(VoUtil.InvokeExpressionBy(3, Function() value * 2), [Is].EqualTo(6))
        End Sub

        <Test()> Public Sub 値の差替_Integer値_加算()
            Dim value As Integer
            Assert.That(VoUtil.InvokeExpressionBy(3, Function() value + 2), [Is].EqualTo(5))
        End Sub

        <Test()> Public Sub 値の差替_Integer値_減算()
            Dim value As Integer
            Assert.That(VoUtil.InvokeExpressionBy(3, Function() value - 2), [Is].EqualTo(1))
        End Sub

        <Test()> Public Sub 値の差替_Integer値_除算は未対応()
            Try
                Dim value As Integer
                Dim dummy = VoUtil.InvokeExpressionBy(12, Function() value / 4)
                Assert.Fail()
            Catch expected As NotSupportedException
                Assert.That(expected.Message, [Is].EqualTo("除数(right=4)がDoubleになるので未対応"))
            End Try
        End Sub

        <Test()> Public Sub 値の差替_Long値()
            Dim value As Long
            Assert.That(VoUtil.InvokeExpressionBy(3L, Function() value), [Is].EqualTo(3))
        End Sub

        <Test()> Public Sub 値の差替_Long値_四則演算はInt型だけ()
            Try
                Dim value As Long
                Dim dummy = VoUtil.InvokeExpressionBy(CLng(Integer.MaxValue), Function() value + 1L)
                Assert.Fail()
            Catch expected As NotSupportedException
                Assert.That(expected.Message, [Is].EqualTo("四則演算はInt型同士だけ可能. 必要なら拡張すべし (left, right) = (Int64, Int64)"))
            End Try
        End Sub

        <Test()> Public Sub 値の差替_DateTime値()
            Dim value As DateTime
            Assert.That(VoUtil.InvokeExpressionBy(CDate("2020/10/12"), Function() value.AddDays(3)), [Is].EqualTo(CDate("2020/10/15")))
        End Sub

        <Test()> Public Sub 値の差替_PVO値()
            Dim value As IntPVO = Nothing
            Assert.That(VoUtil.InvokeExpressionBy(New IntPVO(4), Function() value.AddOne()), [Is].EqualTo(New IntPVO(5)))
        End Sub

        <Test()> Public Sub 値の差替_Integer値_で配列の値を取得できる()
            Dim idx As Integer
            Dim anArray As String() = {"a", "b", "c"}
            Assert.That(VoUtil.InvokeExpressionBy(2, Function() anArray(idx)), [Is].EqualTo("c"))
        End Sub

    End Class

    Public Class GetPropertyNamesTest : Inherits VoUtilTest

        Private Class ParentClass
            Protected Property Protect As String
            Public Property ParentPublic As String
        End Class

        Private Class ChildClass : Inherits ParentClass
            Private Property Hidden As String
            Public Property ChildPublic As String
        End Class

        <Test()>
        Public Sub Nothingを渡すと例外を返す()
            Try
                VoUtil.GetPropertyNames(Nothing)
                Assert.Fail()
            Catch expected As ArgumentNullException
                Assert.That(expected.Message, [Is].EqualTo("値を Null にすることはできません。" & vbCrLf & "パラメーター名:type").Or.StringContaining("Value cannot be null."))
            End Try
        End Sub

        <Test()>
        Public Sub 型を指定すればNothingでも_プロパティ名の配列を返す()
            Assert.That(VoUtil.GetPropertyNames(Of SampleVo)(vo:=Nothing), [Is].EqualTo({"HogeId", "HogeName", "HogeDate", "HogeDecimal", "HogeEnum"}))
        End Sub

        <Test()>
        Public Sub 型が明示されていればNothingでも_プロパティ名の配列を返す()
            Dim vo As SampleVo = Nothing
            Assert.That(VoUtil.GetPropertyNames(vo), [Is].EqualTo({"HogeId", "HogeName", "HogeDate", "HogeDecimal", "HogeEnum"}))
        End Sub

        <Test()>
        Public Sub 型を指定すれば_プロパティ名の配列を返す()
            Assert.That(VoUtil.GetPropertyNames(Of SampleVo), [Is].EqualTo({"HogeId", "HogeName", "HogeDate", "HogeDecimal", "HogeEnum"}))
        End Sub

        <Test()>
        Public Sub 型を渡せば_プロパティ名の配列を返す()
            Assert.That(VoUtil.GetPropertyNames(GetType(SampleVo)), [Is].EqualTo({"HogeId", "HogeName", "HogeDate", "HogeDecimal", "HogeEnum"}))
        End Sub

        <Test()>
        Public Sub Voを渡せば_プロパティ名の配列を返す()
            Assert.That(VoUtil.GetPropertyNames(New SampleVo), [Is].EqualTo({"HogeId", "HogeName", "HogeDate", "HogeDecimal", "HogeEnum"}))
        End Sub

        <Test()>
        Public Sub 外部参照可能なProperty名のみ_取得できる()
            Assert.That(VoUtil.GetPropertyNames(New ChildClass), [Is].EquivalentTo({"ParentPublic", "ChildPublic"}))
        End Sub

        <Test()>
        Public Sub Delegateを渡す事で_抽出条件を指定して_プロパティ名一覧を取得できる()
            Assert.That(VoUtil.GetPropertyNames(New SampleVo, Function(name) StringUtil.GetLengthByte(name) = 8), [Is].EquivalentTo({"HogeName", "HogeDate", "HogeEnum"}))
        End Sub

        <Test()>
        Public Sub プロパティを指定するとそれらのプロパティ名一覧を取得できる()
            Assert.That(VoUtil.GetSpecifyPropertyNames(Of SampleVo)(Function(define, vo)
                                                                        define.AppendProperties(vo.HogeEnum)
                                                                        define.AppendProperties(vo.HogeDecimal, vo.HogeDate)
                                                                        Return define
                                                                    End Function), [Is].EquivalentTo({"HogeEnum", "HogeDecimal", "HogeDate"}))
        End Sub

        <Test()>
        Public Sub ラムダ式で指定した順番でプロパティ名一覧を取得できる()
            Dim actual As String() = VoUtil.GetSpecifyPropertyNames(Of SampleVo)(Function(define, vo)
                                                                                     define.AppendProperties(vo.HogeEnum, vo.HogeDate, vo.HogeId)
                                                                                     Return define
                                                                                 End Function)
            Assert.That(actual, [Is].EqualTo(New String() {"HogeEnum", "HogeDate", "HogeId"}))
        End Sub

        <Test()>
        Public Sub Boolean三項目以上のVoでも_AppendPropertyWithFuncを使えばプロパティ名を取得できる()
            Dim actual As String() = VoUtil.GetSpecifyPropertyNames(Of Has3BooleanVo)(Function(define, vo)
                                                                                          define.AppendPropertyWithFunc(Function() vo.IsB).AppendPropertyWithFunc(Function() vo.IsA).AppendPropertyWithFunc(Function() vo.IsC)
                                                                                          Return define
                                                                                      End Function)
            Assert.That(actual, [Is].EqualTo(New String() {"IsB", "IsA", "IsC"}))
        End Sub

    End Class

    Public Class PopulateValueTest : Inherits VoUtilTest
        Private Class PopulateTestVo
            Public PublicVal As String
            Public Property PublicProp As String
            Private privateVal As String
            Private Property PrivateProp As String
        End Class

        Private Overloads Function ToString(ParamArray vos As PopulateTestVo()) As String
            Dim maker As New DebugStringMaker(Of PopulateTestVo)(Function(define, vo)
                                                                     define.Bind(vo.PublicProp)
                                                                     define.Bind(vo.PublicVal)
                                                                     Return define
                                                                 End Function)
            Return maker.MakeString(vos)
        End Function

        <Test()>
        Public Sub プロパティ名を指定すれば値をセットできる()
            Dim vo As New PopulateTestVo

            VoUtil.SetValue(vo, "PublicProp", "val")

            Assert.That(ToString(vo), [Is].EqualTo(
                        "PublicProp PublicVal" & vbCrLf &
                        "'val'      null     "))

        End Sub

        <Test()>
        Public Sub プロパティ名を指定すればNull値でもセットできる()
            Dim vo As New PopulateTestVo With {.PublicProp = "hogehoge"}

            VoUtil.SetValue(vo, "PublicProp", Nothing)

            Assert.That(ToString(vo), [Is].EqualTo(
                        "PublicProp PublicVal" & vbCrLf &
                        "null       null     "))

        End Sub

        <Test()>
        Public Sub 存在しないプロパティ名を指定された場合_例外を返す()
            Dim vo As New PopulateTestVo
            Try
                VoUtil.SetValue(vo, "HogeProp", "val")
                Assert.Fail()
            Catch ex As NotSupportedException
                Assert.That(ex.Message, [Is].EqualTo("プロパティ名:HogePropは型:PopulateTestVoでサポートされていません"))
            End Try
        End Sub

        <Test()>
        Public Sub 変数名での指定はサポート外なので_例外を返す()
            Dim vo As New PopulateTestVo

            Try
                VoUtil.SetValue(vo, "PublicVal", "val")
                Assert.Fail()
            Catch ex As NotSupportedException
                Assert.That(ex.Message, [Is].EqualTo("プロパティ名:PublicValは型:PopulateTestVoでサポートされていません"))
            End Try
        End Sub

        <Test()>
        Public Sub 参照できないスコープのプロパティを指定された場合_例外を返す()
            Dim vo As New PopulateTestVo

            Try
                VoUtil.SetValue(vo, "PrivateProp", "val")
                Assert.Fail()
            Catch ex As NotSupportedException
                Assert.That(ex.Message, [Is].EqualTo("プロパティ名:PrivatePropは型:PopulateTestVoでサポートされていません"))
            End Try
        End Sub

        <Test()>
        Public Sub 引数Voのインスタンスがなければ_例外を返す()
            Dim vo As PopulateTestVo = Nothing

            Try
                VoUtil.SetValue(Of PopulateTestVo)(vo, "PublicProp", "val")
                Assert.Fail()
            Catch ex As ArgumentNullException
                Assert.That(ex.Message, [Is].EqualTo("値を Null にすることはできません。" & vbCrLf & "パラメーター名:vo").Or.StringContaining("Value cannot be null."))
            End Try
        End Sub

    End Class

    Public Class HasPropertyTest : Inherits VoUtilTest

        <TestCase("Id")>
        <TestCase("Name")>
        <TestCase("Fx")>
        Public Sub 存在するプロパティ名ならTrue(propertyName As String)
            Assert.True(VoUtil.HasProperty(New Hoge, propertyName))
            Assert.True(VoUtil.HasProperty(Of Hoge)(propertyName))
        End Sub

        <TestCase("id")>
        <TestCase("NAME")>
        <TestCase("fX")>
        Public Sub 大文字小文字不一致ならFalse(propertyName As String)
            Assert.False(VoUtil.HasProperty(New Hoge, propertyName))
            Assert.False(VoUtil.HasProperty(Of Hoge)(propertyName))
        End Sub

        <TestCase("No")>
        <TestCase("Namae")>
        Public Sub 存在しないプロパティ名ならFalse(propertyName As String)
            Assert.False(VoUtil.HasProperty(New Hoge, propertyName))
            Assert.False(VoUtil.HasProperty(Of Hoge)(propertyName))
        End Sub

        <Test()> Public Sub VOがNULLなら例外を投げる()
            Dim vo As Hoge = Nothing
            Try
                Assert.False(VoUtil.HasProperty(vo, "AAAA"))
                Assert.Fail()
            Catch ex As ArgumentNullException
                Assert.That(ex.Message, [Is].EqualTo("値を Null にすることはできません。" & vbCrLf & "パラメーター名:vo").Or.StringContaining("Value cannot be null."))
            End Try
        End Sub

    End Class

    Public Class GetPropertyNameTest : Inherits VoUtilTest

        <Test()>
        Public Sub Boolean三項目以上持つVoでも_Boolean以外のプロパティを指定するなら_プロパティ名取得できる()
            Assert.That(VoUtil.GetPropertyName(Of Has3BooleanVo)(Function(vo) vo.Id), [Is].EqualTo("Id"))
        End Sub

        <Test()>
        Public Sub ラムダ式を返すラムダ式を引数に指定すると_プロパティ名を取得できる()
            Assert.That(VoUtil.GetPropertyName(Of Hoge)(Function(vo) vo.Id), [Is].EqualTo("Id"))
            Assert.That(VoUtil.GetPropertyName(Of Hoge)(Function(vo) vo.Name), [Is].EqualTo("Name"))
        End Sub

        <Test()>
        Public Sub Boolean三項目以上持つVoでもプロパティ名取得できる()
            Assert.That(VoUtil.GetPropertyName(Of Has3BooleanVo)(Function(vo) vo.Id), [Is].EqualTo("Id"))
            Assert.That(VoUtil.GetPropertyName(Of Has3BooleanVo)(Function(vo) vo.IsA), [Is].EqualTo("IsA"))
            Assert.That(VoUtil.GetPropertyName(Of Has3BooleanVo)(Function(vo) vo.IsC), [Is].EqualTo("IsC"))
        End Sub

    End Class

End Class
