Imports NUnit.Framework
Imports System.Text
Imports Fhi.Fw.Domain

Public MustInherit Class TypeUtilTest

#Region "テストに使用するVO"
    Protected Class Hoge
        Private _Id As Nullable(Of Integer)
        Private _Name As String
        Private _Fx As Nullable(Of Decimal)
        Private _Adate As DateTime?

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
#End Region
#Region "Nested classes..."
    Private Class MyEnumerable : Implements IEnumerable
        Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Throw New NotImplementedException()
        End Function
    End Class
#End Region

    <Test()> Public Sub IsTypeAnonymous_匿名型のインスタンスでtrue()
        Dim val = New With {.Id = 123, .Name = "hoge"}
        Assert.AreEqual(True, TypeUtil.IsTypeAnonymous(val))
    End Sub

    <Test()> Public Sub IsTypeAnonymous_匿名型でtrue()
        Dim val = New With {.Id = 123, .Name = "hoge"}
        Assert.AreEqual(True, TypeUtil.IsTypeAnonymous(val.GetType))
    End Sub

    <Test()> Public Sub IsTypeAnonymous_ユーザー定義型でfalse()
        Dim val As New Hoge With {.Id = 123, .Name = "hoge"}
        Assert.AreEqual(False, TypeUtil.IsTypeAnonymous(val.GetType))
    End Sub

    <Test()> Public Sub IsTypeArray_配列ならtrue()
        Dim value As String() = {}
        Assert.AreEqual(True, TypeUtil.IsTypeArray(value.GetType), "インスタンスから指定")
        Assert.AreEqual(True, TypeUtil.IsTypeArray(GetType(Integer())), "型を直接指定")
    End Sub

    <Test()> Public Sub IsTypeArray_ICollectionはfalse()
        Dim value As New List(Of String)
        Assert.AreEqual(False, TypeUtil.IsTypeArray(value.GetType), "インスタンスから指定")
        Assert.AreEqual(False, TypeUtil.IsTypeArray(GetType(List(Of Integer))), "型を直接指定")
        Assert.AreEqual(False, TypeUtil.IsTypeArray(GetType(ArrayList)), "ジェネリックじゃないList")
        Assert.AreEqual(False, TypeUtil.IsTypeArray(GetType(System.Array)), "Array型もコレクション扱い")
    End Sub

    <Test()> Public Sub IsTypeICollection_ICollectionでも配列でもtrue()
        Dim value As New List(Of String)
        Assert.AreEqual(True, TypeUtil.IsTypeCollection(value.GetType), "インスタンスから指定")
        Assert.AreEqual(True, TypeUtil.IsTypeCollection(GetType(List(Of Integer))), "型を直接指定")
        Assert.AreEqual(True, TypeUtil.IsTypeCollection(GetType(ICollection(Of Integer))), "型を直接指定")
        Assert.AreEqual(True, TypeUtil.IsTypeCollection(GetType(ArrayList)), "ジェネリックじゃないList")
        Assert.AreEqual(True, TypeUtil.IsTypeCollection(GetType(System.Array)), "Array型もコレクション扱い")
        Dim value2 As String() = {}
        Assert.AreEqual(True, TypeUtil.IsTypeCollection(value2.GetType), "インスタンスから指定")
        Assert.AreEqual(True, TypeUtil.IsTypeCollection(GetType(Integer())), "型を直接指定")
    End Sub

    <Test()> Public Sub DetectElementType_ListOf型の要素を返す()
        Dim listOfStr As New List(Of String)
        Assert.AreEqual(GetType(String), TypeUtil.DetectElementType(listOfStr.GetType))
        Dim listOfInt As New List(Of Integer)
        Assert.AreEqual(GetType(Integer), TypeUtil.DetectElementType(listOfInt.GetType))
        Dim listOfNullableInt As New List(Of Integer?)
        Assert.AreEqual(GetType(Integer?), TypeUtil.DetectElementType(listOfNullableInt.GetType))
    End Sub

    <Test()> Public Sub DetectElementType_配列型の要素を返す()
        Dim arrayOfStr As String() = {}
        Assert.AreEqual(GetType(String), TypeUtil.DetectElementType(arrayOfStr.GetType))
        Dim arrayOfInt As Integer() = {}
        Assert.AreEqual(GetType(Integer), TypeUtil.DetectElementType(arrayOfInt.GetType))
        Dim arrayOfNullableInt As Integer?() = {}
        Assert.AreEqual(GetType(Integer?), TypeUtil.DetectElementType(arrayOfNullableInt.GetType))
    End Sub

    Public Class IsTypeImmutable_不変オブジェクトかを返す_Nullable型も不変オブジェクト扱い : Inherits TypeUtilTest
        Private Class SubValueObject : Inherits ValueObject
            Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
                Throw New NotImplementedException()
            End Function
        End Class
        <Test()> Public Sub 不変オブジェクトならtrue()
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(Integer)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(Long)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(String)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(DateTime)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(Decimal)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(Single)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(Double)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(Short)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(Char)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(Boolean)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(Integer?)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(DateTime?)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(Decimal?)))
        End Sub

        <Test()> Public Sub ValueObjectのサブクラスも_不変オブジェクトとして_true()
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(ValueObject)))
            Assert.IsTrue(TypeUtil.IsTypeImmutable(GetType(SubValueObject)))
        End Sub

        <Test()> Public Sub 不変オブジェクトでなければfalse()
            Assert.IsFalse(TypeUtil.IsTypeImmutable(GetType(Object)))
            Assert.IsFalse(TypeUtil.IsTypeImmutable(GetType(StringBuilder)))
            Assert.IsFalse(TypeUtil.IsTypeImmutable(GetType(VoUtil)))
        End Sub

        <TestCase(GetType(ValueObject), True)>
        <TestCase(GetType(SubValueObject), True)>
        <TestCase(GetType(StringBuilder), False)>
        <TestCase(GetType(Integer), False)>
        Public Sub IsTypeValueObjectOrSubClass_ValueObjectのサブクラスかを判定できる(aType As Type, expected As Boolean)
            Assert.That(TypeUtil.IsTypeValueObjectOrSubClass(aType), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class IsArrayOrCollection_配列かコレクションならtrue : Inherits TypeUtilTest
        <Test()> Public Sub 一連のコレクション型はtrue()
            Dim arrayOfStr As String() = {}
            Assert.IsTrue(TypeUtil.IsArrayOrCollection(arrayOfStr), "文字列型配列はtrue")
            Dim arrayOfInt As Integer() = {}
            Assert.IsTrue(TypeUtil.IsArrayOrCollection(arrayOfInt), "int型配列はtrue")
            Dim arrayOfNullableInt As Long?() = {}
            Assert.IsTrue(TypeUtil.IsArrayOrCollection(arrayOfNullableInt), "Loong?型配列はtrue")
            Dim listOfString As New List(Of String)
            Assert.IsTrue(TypeUtil.IsArrayOrCollection(listOfString), "文字列型Listはtrue")
            Dim listOfInt As New List(Of Integer)
            Assert.IsTrue(TypeUtil.IsArrayOrCollection(listOfInt), "int型Listはtrue")
            Dim listOfNullableInt As New List(Of Integer?)
            Assert.IsTrue(TypeUtil.IsArrayOrCollection(listOfNullableInt), "int?型Listはtrue")
        End Sub

        <Test(), Sequential()> Public Sub コレクション型以外はfalse( _
                <Values("str", 123)> ByVal val As Object, _
                <Values("文字列型", "int型")> ByVal msg As String)

            Assert.IsFalse(TypeUtil.IsArrayOrCollection(val), msg)
        End Sub

        <Test()> Public Sub IEnumerableはfalse()
            Assert.IsFalse(TypeUtil.IsArrayOrCollection(New MyEnumerable), "IEnumerableはfalse")
        End Sub
    End Class

    Public Class IsTypeCollection_コレクション型ならtrue : Inherits TypeUtilTest
        <Test(), Sequential()> Public Sub 一連のコレクション型はtrue( _
                <Values(GetType(String()), GetType(Integer()), GetType(Long?()), GetType(List(Of String)), GetType(List(Of DateTime)), GetType(List(Of Decimal?)))> ByVal aType As Type, _
                <Values("文字列配列", "int型配列", "Long?型配列", "文字列List", "DateTime型List", "Decimal?型List")> ByVal msg As String)

            Assert.IsTrue(TypeUtil.IsTypeCollection(aType), msg)
        End Sub

        <Test(), Sequential()> Public Sub コレクション型以外はfalse( _
                <Values(GetType(String), GetType(Integer), GetType(Long?), GetType(MyEnumerable))> ByVal aType As Type, _
                <Values("文字列型", "int型", "Long?型", "IEnumerableはfalse")> ByVal msg As String)

            Assert.IsFalse(TypeUtil.IsTypeCollection(aType), msg)
        End Sub
    End Class

    Public Class IsTypeGenericOrSubClass_Genericな継承先かを判定できる : Inherits TypeUtilTest
        Private Class TestingGeneric(Of T)
        End Class
        Private Class SubClass : Inherits TestingGeneric(Of String)
        End Class

        <TestCase(GetType(SubClass), GetType(TestingGeneric(Of )), True, "ワイルドカード風の指定でtrue")>
        <TestCase(GetType(SubClass), GetType(TestingGeneric(Of String)), True, "SubClassは継承してるのでtrue")>
        <TestCase(GetType(SubClass), GetType(TestingGeneric(Of Integer)), False, "Integerじゃないのでfalse")>
        <TestCase(GetType(TestingGeneric(Of Integer)), GetType(TestingGeneric(Of Integer)), True, "同じ型はtrue")>
        <TestCase(GetType(TestingGeneric(Of String)), GetType(TestingGeneric(Of Integer)), False, "")>
        <TestCase(GetType(TestingGeneric(Of String)), GetType(TestingGeneric(Of )), True, "")>
        <TestCase(GetType(SubClass), GetType(String), False, "無関係だからfalse")>
        Public Sub 判定できる(aType As Type, aGenericType As Type, expected As Boolean, comment As String)
            Assert.That(TypeUtil.IsTypeGenericOrSubClass(aType, aGenericType), [Is].EqualTo(expected), comment)
        End Sub

    End Class

End Class
