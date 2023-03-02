Imports NUnit.Framework

Namespace Domain
    Public MustInherit Class PrimitiveValueObjectTest

#Region "Testing classes..."
        Private Class StringAttr : Inherits PrimitiveValueObject(Of String)
            Public Sub New(ByVal value As String)
                MyBase.New(value)
            End Sub
            Public Overloads ReadOnly Property InternalValue As String
                Get
                    Return MyBase.Value
                End Get
            End Property
        End Class
        Private Class IntAttr : Inherits PrimitiveValueObject(Of Integer)
            Public Sub New(ByVal value As Integer)
                MyBase.New(value)
            End Sub
            Public Overloads ReadOnly Property InternalValue As Integer
                Get
                    Return MyBase.Value
                End Get
            End Property
        End Class
        Private Class NullableIntAttr : Inherits PrimitiveValueObject(Of Integer?)
            Public Sub New(ByVal value As Integer?)
                MyBase.New(value)
            End Sub
            Public Overloads ReadOnly Property InternalValue As Integer?
                Get
                    Return MyBase.Value
                End Get
            End Property
        End Class
#End Region

        <SetUp()> Public Overridable Sub SetUp()

        End Sub

        Public Class String型でTest : Inherits PrimitiveValueObjectTest

            <Test()> Public Sub nullだったら_中身はnullになる()
                Dim actual As New StringAttr(Nothing)
                Assert.That(actual.InternalValue, [Is].Null)
                Assert.That(actual.ToString, [Is].Null)
            End Sub

            <TestCase("")>
            <TestCase("  ")>
            <TestCase("abc")>
            <TestCase("123")>
            <TestCase("日本語")>
            Public Sub null以外なら_そのままStringで保持できる(value As String)
                Dim actual As New StringAttr(value)
                Assert.That(actual.InternalValue, [Is].EqualTo(value))
                Assert.That(actual.ToString, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub 標準でソートできる()
                Dim list As New List(Of StringAttr)({New StringAttr("c"), New StringAttr("aa"), New StringAttr("bbb")})
                list.Sort()
                Assert.That(list(0), [Is].EqualTo(New StringAttr("aa")))
                Assert.That(list(1), [Is].EqualTo(New StringAttr("bbb")))
                Assert.That(list(2), [Is].EqualTo(New StringAttr("c")))
            End Sub

            <Test()> Public Sub 標準でソートできる_Null含む()
                Dim list As New List(Of StringAttr)({New StringAttr("c"), New StringAttr("aa"), New StringAttr(Nothing)})
                list.Sort()
                Assert.That(list(0), [Is].EqualTo(New StringAttr(Nothing)))
                Assert.That(list(1), [Is].EqualTo(New StringAttr("aa")))
                Assert.That(list(2), [Is].EqualTo(New StringAttr("c")))
            End Sub

        End Class

        Public Class Int型でTest : Inherits PrimitiveValueObjectTest

            <TestCase(3, "3")>
            <TestCase(100, "100")>
            Public Sub Int値を保持できる(value As Integer, strValue As String)
                Dim actual As New IntAttr(value)
                Assert.That(actual.InternalValue, [Is].EqualTo(value))
                Assert.That(actual.ToString, [Is].EqualTo(strValue))
            End Sub

            <Test()> Public Sub 標準でソートできる()
                Dim list As New List(Of IntAttr)({New IntAttr(100), New IntAttr(5), New IntAttr(30)})
                list.Sort()
                Assert.That(list(0), [Is].EqualTo(New IntAttr(5)))
                Assert.That(list(1), [Is].EqualTo(New IntAttr(30)))
                Assert.That(list(2), [Is].EqualTo(New IntAttr(100)))
            End Sub

        End Class

        Public Class NullableInt型でTest : Inherits PrimitiveValueObjectTest

            <TestCase(Nothing)>
            Public Sub nullだったら_中身はnullになる(value As Integer?)
                Dim actual As New NullableIntAttr(value)
                Assert.That(actual.InternalValue, [Is].Null)
                Assert.That(actual.ToString, [Is].Null)
            End Sub

            <TestCase(3)>
            <TestCase(100)>
            Public Sub Int値を保持できる(value As Integer?)
                Dim actual As New NullableIntAttr(value)
                Assert.That(actual.InternalValue, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub 標準でソートできる()
                Dim list As New List(Of NullableIntAttr)({New NullableIntAttr(100), New NullableIntAttr(5), New NullableIntAttr(30)})
                list.Sort()
                Assert.That(list(0), [Is].EqualTo(New NullableIntAttr(5)))
                Assert.That(list(1), [Is].EqualTo(New NullableIntAttr(30)))
                Assert.That(list(2), [Is].EqualTo(New NullableIntAttr(100)))
            End Sub

            <Test()> Public Sub 標準でソートできる_Null含む()
                Dim list As New List(Of NullableIntAttr)({New NullableIntAttr(100), New NullableIntAttr(-5), New NullableIntAttr(Nothing)})
                list.Sort()
                Assert.That(list(0), [Is].EqualTo(New NullableIntAttr(Nothing)))
                Assert.That(list(1), [Is].EqualTo(New NullableIntAttr(-5)))
                Assert.That(list(2), [Is].EqualTo(New NullableIntAttr(100)))
            End Sub

        End Class

        Public Class IsEmptyTest : Inherits PrimitiveValueObjectTest

            Private Class Attr(Of T) : Inherits PrimitiveValueObject(Of T)
                Public Sub New(ByVal value As T)
                    MyBase.New(value)
                End Sub
            End Class

            <Test()> Public Sub NullableIntで_NothingならTrue()
                Assert.That((New NullableIntAttr(Nothing)).IsEmpty, [Is].True)
                Assert.That((New NullableIntAttr(Nothing)).IsNotEmpty, [Is].False)
            End Sub

            <TestCase(0)>
            <TestCase(1)>
            <TestCase(234)>
            Public Sub NullableIntで_値があれば_偽になる(arg As Integer?)
                Assert.That((New NullableIntAttr(arg)).IsEmpty, [Is].False)
                Assert.That((New NullableIntAttr(arg)).IsNotEmpty, [Is].True)
            End Sub

            <Test()> Public Sub Intで_NothingならFalse()
                Assert.That((New IntAttr(Nothing)).IsEmpty, [Is].False)
                Assert.That((New IntAttr(Nothing)).IsNotEmpty, [Is].True)
            End Sub

            <TestCase(Nothing)>
            <TestCase(0)>
            <TestCase(1)>
            Public Sub Intで_値があれば_偽になる(arg As Integer)
                Assert.That((New IntAttr(arg)).IsEmpty, [Is].False)
                Assert.That((New IntAttr(arg)).IsNotEmpty, [Is].True)
            End Sub

            <Test()> Public Sub NullableLongで_NothingならTrue()
                Assert.That((New Attr(Of Long?)(Nothing)).IsEmpty, [Is].True)
                Assert.That((New Attr(Of Long?)(Nothing)).IsNotEmpty, [Is].False)
            End Sub

            <TestCase(0L)>
            <TestCase(1L)>
            <TestCase(234L)>
            Public Sub NullableLongで_値があれば_偽になる(arg As Long?)
                Assert.That((New Attr(Of Long?)(arg)).IsEmpty, [Is].False)
                Assert.That((New Attr(Of Long?)(arg)).IsNotEmpty, [Is].True)
            End Sub

            <Test()> Public Sub Longで_NothingならFalse()
                Assert.That((New Attr(Of Long)(Nothing)).IsEmpty, [Is].False)
                Assert.That((New Attr(Of Long)(Nothing)).IsNotEmpty, [Is].True)
            End Sub

            <TestCase(Nothing)>
            <TestCase(0L)>
            <TestCase(1L)>
            Public Sub Longで_値があれば_偽になる(arg As Long)
                Assert.That((New Attr(Of Long)(arg)).IsEmpty, [Is].False)
                Assert.That((New Attr(Of Long)(arg)).IsNotEmpty, [Is].True)
            End Sub

            <Test()> Public Sub NullableDecimalで_NothingならTrue()
                Assert.That((New Attr(Of Decimal?)(Nothing)).IsEmpty, [Is].True)
                Assert.That((New Attr(Of Decimal?)(Nothing)).IsNotEmpty, [Is].False)
            End Sub

            <TestCase("0")>
            <TestCase("1")>
            <TestCase("3.14")>
            Public Sub NullableDecimalで_値があれば_偽になる(decimalString As String)
                Dim value As Decimal? = CDec(decimalString)
                Assert.That((New Attr(Of Decimal?)(value)).IsEmpty, [Is].False)
                Assert.That((New Attr(Of Decimal?)(value)).IsNotEmpty, [Is].True)
            End Sub

            <Test()> Public Sub Decimalで_NothingならFalse()
                Assert.That((New Attr(Of Decimal)(Nothing)).IsEmpty, [Is].False)
                Assert.That((New Attr(Of Decimal)(Nothing)).IsNotEmpty, [Is].True)
            End Sub

            <TestCase(Nothing)>
            <TestCase("0")>
            <TestCase("1")>
            Public Sub Decimalで_値があれば_偽になる(decimalString As String)
                Dim value As Decimal = If(IsNumeric(decimalString), CDec(decimalString), Nothing)
                Assert.That((New Attr(Of Decimal)(value)).IsEmpty, [Is].False)
                Assert.That((New Attr(Of Decimal)(value)).IsNotEmpty, [Is].True)
            End Sub

            <Test()> Public Sub NullableDateTimeで_NothingならTrue()
                Assert.That((New Attr(Of DateTime?)(Nothing)).IsEmpty, [Is].True)
                Assert.That((New Attr(Of DateTime?)(Nothing)).IsNotEmpty, [Is].False)
            End Sub

            <TestCase("1940/1/1")>
            <TestCase("1970/1/1")>
            <TestCase("2020/2/2")>
            Public Sub NullableDateTimeで_値があれば_偽になる(dateString As String)
                Dim value As DateTime? = CDate(dateString)
                Assert.That((New Attr(Of DateTime?)(value)).IsEmpty, [Is].False)
                Assert.That((New Attr(Of DateTime?)(value)).IsNotEmpty, [Is].True)
            End Sub

            <Test()> Public Sub DateTimeで_NothingならFalse()
                Assert.That((New Attr(Of DateTime)(Nothing)).IsEmpty, [Is].False)
                Assert.That((New Attr(Of DateTime)(Nothing)).IsNotEmpty, [Is].True)
            End Sub

            <TestCase(Nothing)>
            <TestCase("1970/1/1")>
            <TestCase("2020/2/2")>
            Public Sub DateTimeで_値があれば_偽になる(dateString As String)
                Dim value As Date = If(IsDate(dateString), CDate(dateString), Nothing)
                Assert.That((New Attr(Of DateTime)(value)).IsEmpty, [Is].False)
                Assert.That((New Attr(Of DateTime)(value)).IsNotEmpty, [Is].True)
            End Sub

            <TestCase("")>
            <TestCase(Nothing)>
            Public Sub 値がNullなら_真になる_Str(val As String)
                Assert.That((New StringAttr(val)).IsEmpty, [Is].True)
                Assert.That((New StringAttr(val)).IsNotEmpty, [Is].False)
            End Sub

            <TestCase("a")>
            <TestCase("漢字")>
            Public Sub 値があれば_偽になる_Str(val As String)
                Assert.That((New StringAttr(val)).IsEmpty, [Is].False)
                Assert.That((New StringAttr(val)).IsNotEmpty, [Is].True)
            End Sub

            <Test()>
            Public Sub Nullable_HasValueのようにPVOインターフェース自身がNothingでも_IsEmptyは使える()
                Dim pvoStr As PrimitiveValueObject = Nothing
                Assert.That(pvoStr.IsEmpty, [Is].True)
                Assert.That(pvoStr.IsNotEmpty, [Is].False)
            End Sub

            <Test()>
            Public Sub Nullable_HasValueのようにPVO実装インスタンスがNothingでも_IsEmptyは使える()
                Dim pvoStr As StringAttr = Nothing
                Assert.That(pvoStr.IsEmpty, [Is].True)
                Assert.That(pvoStr.IsNotEmpty, [Is].False)
            End Sub

        End Class

        Public Class HasValueTest : Inherits PrimitiveValueObjectTest

            Private Class Attr(Of T) : Inherits PrimitiveValueObject(Of T)
                Public Sub New(ByVal value As T)
                    MyBase.New(value)
                End Sub
            End Class

            <Test()>
            Public Sub PVO自身がNothingなら_HasValueはfalse_PVO自身がNothingでも_HasValueは使える()
                Dim pvoStr As Attr(Of String) = Nothing
                Dim pvoNullableInt As Attr(Of Integer?) = Nothing
                Dim pvoInt As Attr(Of Integer) = Nothing
                Dim pvo As PrimitiveValueObject = Nothing
                Assert.That(pvoStr.HasValue, [Is].False)
                Assert.That(pvoNullableInt.HasValue, [Is].False)
                Assert.That(pvoInt.HasValue, [Is].False)
                Assert.That(pvo.HasValue, [Is].False)
            End Sub

            <Test()>
            Public Sub PVOが持つ値がNothingなら_HasValueはfalse()
                Dim pvoStr As New Attr(Of String)(Nothing)
                Dim pvoNullableInt As New Attr(Of Integer?)(Nothing)
                Dim pvoInt As New Attr(Of Integer)(Nothing)
                Assert.That(pvoStr.HasValue, [Is].False)
                Assert.That(pvoNullableInt.HasValue, [Is].False)
                Assert.That(pvoInt.HasValue, [Is].True, "Int型はnothingにならないからtrue")
            End Sub

            <Test()>
            Public Sub PVOが持つ値がNothing以外なら_HasValueはtrue()
                Dim pvoStr As New Attr(Of String)("")
                Dim pvoNullableInt As New Attr(Of Integer?)(0)
                Dim pvoInt As New Attr(Of Integer)(0)
                Assert.That(pvoStr.HasValue, [Is].True)
                Assert.That(pvoNullableInt.HasValue, [Is].True)
                Assert.That(pvoInt.HasValue, [Is].True)
            End Sub

        End Class

    End Class
End Namespace
