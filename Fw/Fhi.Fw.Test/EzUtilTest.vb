Imports System.Collections.Generic
Imports NUnit.Framework
Imports Fhi.Fw.Domain

<TestFixture()> Public MustInherit Class EzUtilTest

#Region "テスト用PVO"
    Private Class IntPvo : Inherits PrimitiveValueObject(Of Integer?)
        Public Sub New(ByVal value As Integer?)
            MyBase.New(value)
        End Sub
    End Class

    Private Class IntPvo2 : Inherits PrimitiveValueObject(Of Integer?)
        Public Sub New(value As Integer?)
            MyBase.New(value)
        End Sub
    End Class

    Private Class StrPvo : Inherits PrimitiveValueObject(Of String)
        Public Sub New(ByVal value As String)
            MyBase.New(value)
        End Sub
    End Class
    Private Class StrPvo2 : Inherits PrimitiveValueObject(Of String)
        Public Sub New(value As String)
            MyBase.New(value)
        End Sub
    End Class
#End Region

    Public Class 通常 : Inherits EzUtilTest

        <Test()> Public Sub Increment()
            Dim i As Integer = 100
            Assert.AreEqual(100, EzUtil.Increment(i))
            Assert.AreEqual(101, EzUtil.Increment(i))
        End Sub

        <Test()> Public Sub TestConvIndexToAlphabet()
            Assert.AreEqual("A", EzUtil.ConvIndexToAlphabet(0))
            Assert.AreEqual("Z", EzUtil.ConvIndexToAlphabet(25))
            Assert.AreEqual("AA", EzUtil.ConvIndexToAlphabet(26 * 1 + 0))
            Assert.AreEqual("ZZ", EzUtil.ConvIndexToAlphabet(26 * 26 + 25))
            Assert.AreEqual("AAA", EzUtil.ConvIndexToAlphabet(26 * 26 * 1 + 26 * 1 + 0))
            Assert.AreEqual("EG", EzUtil.ConvIndexToAlphabet(26 * 5 + 6))
        End Sub
    End Class

    Public Class IsEqualTest : Inherits EzUtilTest

        <Test()> Public Sub String値がNotNull()
            Assert.IsTrue(EzUtil.IsEqualIfNull("x", "x"))
            Assert.IsFalse(EzUtil.IsEqualIfNull("a", "b"))
        End Sub

        <Test()> Public Sub 片方だけ値がnull()
            Assert.IsFalse(EzUtil.IsEqualIfNull("b", Nothing))
            Assert.IsFalse(EzUtil.IsEqualIfNull(Nothing, "a"))
        End Sub

        <Test()> Public Sub 双方とも値がnull()
            Assert.IsTrue(EzUtil.IsEqualIfNull(Nothing, Nothing))
        End Sub

        <Test> Public Sub PVOとPrimitive型で比較したら_PVOの中身が同じでもFalse()
            Assert.IsFalse(EzUtil.IsEqualIfNull("abc", New StrPvo("abc")))
            Assert.IsFalse(EzUtil.IsEqualIfNull(123, New IntPvo(123)))
        End Sub
    End Class

    Public Class IsNotEqualIfNullTest : Inherits EzUtilTest

        <Test> Public Sub String値がNotNull()
            Assert.IsFalse(EzUtil.IsNotEqualIfNull("x", "x"))
            Assert.IsTrue(EzUtil.IsNotEqualIfNull("a", "b"))
        End Sub

        <Test> Public Sub 片方だけ値がnull()
            Assert.IsTrue(EzUtil.IsNotEqualIfNull("b", Nothing))
            Assert.IsTrue(EzUtil.IsNotEqualIfNull(Nothing, "a"))
        End Sub

        <Test> Public Sub 双方とも値がnull()
            Assert.IsFalse(EzUtil.IsNotEqualIfNull(Nothing, Nothing))
        End Sub

        <Test> Public Sub PVOとPrimitive型で比較したら_True()
            Assert.IsTrue(EzUtil.IsNotEqualIfNull("abc", New StrPvo("abc")))
            Assert.IsTrue(EzUtil.IsNotEqualIfNull(123, New IntPvo(123)))
        End Sub
    End Class

    Public Class IsEqualIgnorePvoIfNullTest : Inherits EzUtilTest

        <Test> Public Sub Primitive型同士で比較_同じならTrue()
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(Nothing, Nothing))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull("abc", "abc"))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(123, 123))
        End Sub

        <Test> Public Sub Primitive型同士で比較_異なればFalse()
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull("abc", Nothing))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(Nothing, "xyz"))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull("abc", "xyz"))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(123, 789))
        End Sub

        <Test> Public Sub PVO同士で比較できる_同じならTrue()
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo(Nothing), New StrPvo(Nothing)))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo("abc")))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(New IntPvo(123), New IntPvo(123)))
        End Sub

        <Test> Public Sub PVO同士で比較できる_異なればFalse()
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo(Nothing), New StrPvo("xyz")))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo(Nothing)))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo("xyz")))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New IntPvo(123), New IntPvo(789)))
        End Sub

        <Test> Public Sub Primitive型とPVOの比較_PVO内のValueが一致したらTrue()
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo(Nothing), Nothing))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo("abc"), "abc"))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(Nothing, New StrPvo(Nothing)))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull("xyz", New StrPvo("xyz")))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(New IntPvo(123), 123))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(789, New IntPvo(789)))
        End Sub

        <Test> Public Sub Primitive型とPVOの比較_PVO内のValueが不一致ならFalse()
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo("abc"), Nothing))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo(Nothing), "xyz"))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo("abc"), "xyz"))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(Nothing, New StrPvo("xyz")))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull("abc", New StrPvo(Nothing)))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull("abc", New StrPvo("xyz")))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New IntPvo(123), 789))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(123, New IntPvo(789)))
        End Sub

        <Test> Public Sub 異なるPVO同士の比較_PVO内のValueが同じならTrue()
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo(Nothing), New StrPvo2(Nothing)))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo2("abc")))
            Assert.IsTrue(EzUtil.IsEqualIgnorePvoIfNull(New IntPvo(123), New IntPvo2(123)))
        End Sub

        <Test> Public Sub 異なるPVO同士の比較_PVO内のValueが異なればFalse()
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo(Nothing), New StrPvo2("xyz")))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo2(Nothing)))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo2("xyz")))
            Assert.IsFalse(EzUtil.IsEqualIgnorePvoIfNull(New IntPvo(123), New IntPvo2(789)))
        End Sub
    End Class

    Public Class IsNotEqualIgnorePvoIfNullTest : Inherits EzUtilTest

        <Test> Public Sub Primitive型同士で比較_同じならFalse()
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(Nothing, Nothing))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull("abc", "abc"))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(123, 123))
        End Sub

        <Test> Public Sub Primitive型同士で比較_異なればTrue()
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull("abc", Nothing))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(Nothing, "xyz"))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull("abc", "xyz"))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(123, 789))
        End Sub

        <Test> Public Sub PVO同士で比較できる_同じならFalse()
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo(Nothing), New StrPvo(Nothing)))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo("abc")))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(New IntPvo(123), New IntPvo(123)))
        End Sub

        <Test> Public Sub PVO同士で比較できる_異なればTrue()
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo(Nothing), New StrPvo("xyz")))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo(Nothing)))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo("xyz")))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New IntPvo(123), New IntPvo(789)))
        End Sub

        <Test> Public Sub Primitive型とPVOの比較_PVO内のValueが一致したらFalse()
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo(Nothing), Nothing))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo("abc"), "abc"))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(Nothing, New StrPvo(Nothing)))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull("xyz", New StrPvo("xyz")))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(New IntPvo(123), 123))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(789, New IntPvo(789)))
        End Sub

        <Test> Public Sub Primitive型とPVOの比較_PVO内のValueが不一致ならTrue()
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo("abc"), Nothing))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo(Nothing), "xyz"))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo("abc"), "xyz"))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(Nothing, New StrPvo("xyz")))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull("abc", New StrPvo(Nothing)))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull("abc", New StrPvo("xyz")))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New IntPvo(123), 789))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(123, New IntPvo(789)))
        End Sub

        <Test> Public Sub 異なるPVO同士の比較_PVO内のValueが同じならFalse()
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo(Nothing), New StrPvo2(Nothing)))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo2("abc")))
            Assert.IsFalse(EzUtil.IsNotEqualIgnorePvoIfNull(New IntPvo(123), New IntPvo2(123)))
        End Sub

        <Test> Public Sub 異なるPVO同士の比較_PVO内のValueが異なればTrue()
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo(Nothing), New StrPvo2("xyz")))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo2(Nothing)))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New StrPvo("abc"), New StrPvo2("xyz")))
            Assert.IsTrue(EzUtil.IsNotEqualIgnorePvoIfNull(New IntPvo(123), New IntPvo2(789)))
        End Sub
    End Class

    Public Class ConvAlphabetToIndexTest : Inherits EzUtilTest

        <Test()> Public Sub アルファベットindexを数値indexにする()
            Assert.AreEqual(0, EzUtil.ConvAlphabetToIndex("A"))
            Assert.AreEqual(25, EzUtil.ConvAlphabetToIndex("Z"))
            Assert.AreEqual(26, EzUtil.ConvAlphabetToIndex("AA"))
            Assert.AreEqual(26 * 26 + 25, EzUtil.ConvAlphabetToIndex("ZZ"))
            Assert.AreEqual(26 * 26 * 1 + 26 * 1 + 0, EzUtil.ConvAlphabetToIndex("AAA"))
            Assert.AreEqual(26 * 26 * 26 + 26 * 26 + 25, EzUtil.ConvAlphabetToIndex("ZZZ"))
        End Sub

        <Test()> Public Sub アルファベットindexを数値indexにする_小文字()
            Assert.AreEqual(0, EzUtil.ConvAlphabetToIndex("a"))
            Assert.AreEqual(25, EzUtil.ConvAlphabetToIndex("z"))
            Assert.AreEqual(26, EzUtil.ConvAlphabetToIndex("aa"))
            Assert.AreEqual(26 * 26 + 25, EzUtil.ConvAlphabetToIndex("zz"))
        End Sub
    End Class

    Public Class NvlTest : Inherits EzUtilTest

        <Test()> Public Sub Integer_値が入っていれば_値をそのまま返す()
            Dim actual As Integer = EzUtil.Nvl(123, 0)
            Assert.That(actual, [Is].EqualTo(123))
        End Sub

        <Test()> Public Sub Nullable_Integer_値が入っていれば_値をそのまま返す()
            Dim value As Integer? = 123
            Dim actual As Integer = EzUtil.Nvl(value, 0)
            Assert.That(actual, [Is].EqualTo(123))
        End Sub

        <Test()> Public Sub Integer_値がNULLなら_NULLだった場合の値を返す()
            Dim actual As Integer = EzUtil.Nvl(Nothing, 0)
            Assert.That(actual, [Is].EqualTo(0))
        End Sub

        <Test()> Public Sub Decimal_値が入っていれば_値をそのまま返す()
            Dim actual As Decimal = EzUtil.Nvl(123.45D, 0.5D)
            Assert.That(actual, [Is].EqualTo(123.45D))
        End Sub

        <Test()> Public Sub Nullable_Decimal_値が入っていれば_値をそのまま返す()
            Dim value As Decimal? = 123.45D
            Dim actual As Decimal = EzUtil.Nvl(123.45D, 0.5D)
            Assert.That(actual, [Is].EqualTo(123.45D))
        End Sub

        <Test()> Public Sub Decimal_値がNULLなら_NULLだった場合の値を返す()
            Dim actual As Decimal = EzUtil.Nvl(Nothing, 0.5D)
            Assert.That(actual, [Is].EqualTo(0.5D))
        End Sub

        <Test()> Public Sub PrimitiveValueObject_値が入っていれば_値をそのまま返す()
            Dim pvo As New StrPvo("abc")
            Dim actual As StrPvo = EzUtil.Nvl(pvo, New StrPvo("NullVal"))
            Assert.That(actual, [Is].EqualTo(New StrPvo("abc")))
        End Sub

        <Test()> Public Sub PrimitiveValueObject_値がNULLなら_NULLだった場合の値を返す()
            Dim pvo As New StrPvo(Nothing)
            Dim actual As StrPvo = EzUtil.Nvl(pvo, New StrPvo("NullVal"))
            Assert.That(actual, [Is].EqualTo(New StrPvo("NullVal")))
        End Sub

    End Class

    Public Class IsBooleanValueTest : Inherits EzUtilTest

        <Test()> Public Sub TrueやFalseの文字列やboolean値そのものならtrue()
            Assert.IsTrue(EzUtil.IsBooleanValue("true"))
            Assert.IsTrue(EzUtil.IsBooleanValue("trUE"))
            Assert.IsTrue(EzUtil.IsBooleanValue("TRUE"))
            Assert.IsTrue(EzUtil.IsBooleanValue("false"))
            Assert.IsTrue(EzUtil.IsBooleanValue("faLSE"))
            Assert.IsTrue(EzUtil.IsBooleanValue("FALSE"))

            Assert.IsTrue(EzUtil.IsBooleanValue(True))
            Assert.IsTrue(EzUtil.IsBooleanValue(False))
        End Sub

        <Test()> Public Sub TrueやFalse以外ならfalse()
            Assert.IsFalse(EzUtil.IsBooleanValue(Nothing))
            Assert.IsFalse(EzUtil.IsBooleanValue(0))
            Assert.IsFalse(EzUtil.IsBooleanValue(1))
            Assert.IsFalse(EzUtil.IsBooleanValue("0"))
            Assert.IsFalse(EzUtil.IsBooleanValue("1"))
        End Sub
    End Class

    Public Class JoinAsStringTest : Inherits EzUtilTest

        <Test()> Public Sub 文字列のリスト連結()
            Dim tmp As New List(Of String)(New String() {"A", "B", "C", "D", "E", "F"})
            Dim act As String = EzUtil.JoinAsString(", ", tmp)
            Assert.AreEqual("A, B, C, D, E, F", act)
        End Sub

        <Test()> Public Sub 文字列のリスト連結2()
            Dim tmp1 As New List(Of String)(New String() {"A", "B", "C", "D", "E"})
            Dim tmp2 As New List(Of String)(New String() {"a", "b", "c", "d", "e"})
            Dim act As String = EzUtil.JoinAsString(", ", tmp1, tmp2)
            Assert.AreEqual("A, B, C, D, E, a, b, c, d, e", act)
        End Sub

        <Test()> Public Sub 数値のリスト連結()
            Dim tmp As New List(Of Integer)(New Integer() {1, 2, 3, 4, 5, 6})
            Dim act As String = EzUtil.JoinAsString(", ", tmp)
            Assert.AreEqual("1, 2, 3, 4, 5, 6", act)
        End Sub

        <Test()> Public Sub 長さゼロならnullを返す()
            Dim tmp As New List(Of String)(New String() {})
            Assert.IsNull(EzUtil.JoinAsString(",", tmp))
            Assert.IsNull(Join(New String() {}, ","), "Joinもnull")
        End Sub
    End Class

    Public Class MakeKey : Inherits EzUtilTest

        <Test()> Public Sub 一意のキーを作成する()
            Assert.AreEqual("1", EzUtil.MakeKey("1"))
            Assert.AreEqual("A;'"":B", EzUtil.MakeKey("A", "B"))
            Assert.AreEqual("xx;'"":yy;'"":zz", EzUtil.MakeKey("xx", "yy", "zz"))
        End Sub

        <Test()> Public Sub 引数なしならnull()
            Assert.IsNull(EzUtil.MakeKey())
        End Sub
    End Class

    Public Class IsEmptyTest : Inherits EzUtilTest
        <TestCase(Nothing)>
        <TestCase(0)>
        Public Sub IsEmpty_Integerがnullか0なら真になる(arg As Integer?)
            Assert.AreEqual(True, EzUtil.IsEmpty(arg))
        End Sub
        <TestCase(1)>
        Public Sub IsEmpty_Integerが有効値なら偽になる(arg As Integer?)
            Assert.AreEqual(False, EzUtil.IsEmpty(arg))
        End Sub
        <TestCase(Nothing)>
        <TestCase(0L)>
        Public Sub IsEmpty_Longがnullか0なら真になる(arg As Long?)
            Assert.AreEqual(True, EzUtil.IsEmpty(arg))
        End Sub
        <TestCase(1L)>
        Public Sub IsEmpty_Longが有効値なら偽になる(arg As Long?)
            Assert.AreEqual(False, EzUtil.IsEmpty(arg))
        End Sub
        <TestCase(Nothing)>
        <TestCase(0.0F)>
        Public Sub IsEmpty_Singleがnullか0なら真になる(arg As Single?)
            Assert.AreEqual(True, EzUtil.IsEmpty(arg))
        End Sub
        <TestCase(1.1F)>
        Public Sub IsEmpty_Singleが有効値なら偽になる(arg As Single?)
            Assert.AreEqual(False, EzUtil.IsEmpty(arg))
        End Sub
        <TestCase(Nothing)>
        <TestCase(0.0R)>
        Public Sub IsEmpty_Doubleがnullか0なら真になる(arg As Double?)
            Assert.AreEqual(True, EzUtil.IsEmpty(arg))
        End Sub
        <TestCase(1.23R)>
        Public Sub IsEmpty_Doubleが有効値なら偽になる(arg As Double?)
            Assert.AreEqual(False, EzUtil.IsEmpty(arg))
        End Sub
        <TestCase(Nothing)>
        <TestCase("0")>
        Public Sub IsEmpty_Decimalがnullか0なら真になる(arg As String)
            Dim d As Decimal? = Nothing
            If IsNumeric(arg) Then
                d = CDec(arg)
            End If
            Assert.AreEqual(True, EzUtil.IsEmpty(d))
        End Sub
        <TestCase("1.23")>
        Public Sub IsEmpty_Decimalが有効値なら偽になる(arg As String)
            Assert.AreEqual(False, EzUtil.IsEmpty(CDec(arg)))
        End Sub
        <Test()> Public Sub IsEmpty_日付がnullなら真になる()
            Dim da As DateTime? = Nothing
            Assert.AreEqual(True, EzUtil.IsEmpty(da))
        End Sub
        <Test()> Public Sub IsEmpty_日付が0なら真になる()
            Dim da As DateTime? = DateTime.MinValue
            Assert.AreEqual(True, EzUtil.IsEmpty(da))
        End Sub
        <Test()> Public Sub IsEmpty_日付が有効値なら偽になる()
            Dim da As DateTime? = DateTime.Now
            Assert.AreEqual(False, EzUtil.IsEmpty(da))
        End Sub
    End Class

    Public Class IsEmpty_PrimitiveValueObject_Test : Inherits EzUtilTest

        <Test()> Public Sub 自身がNullなら_真になる()
            Dim arg As IntPvo = Nothing
            Assert.That(EzUtil.IsEmpty(arg), [Is].True)
            Assert.That(EzUtil.IsNotEmpty(arg), [Is].False)
        End Sub

        <Test()> Public Sub NullableIntのPVOで_値がNullなら_真になる()
            Dim arg As Integer? = Nothing
            Assert.That(EzUtil.IsEmpty(New IntPvo(arg)), [Is].True)
            Assert.That(EzUtil.IsNotEmpty(New IntPvo(arg)), [Is].False)
        End Sub

        <TestCase(0)>
        <TestCase(1)>
        <TestCase(234)>
        Public Sub NullableIntのPVOで_値があれば_偽になる(arg As Integer?)
            Assert.That(EzUtil.IsEmpty(New IntPvo(arg)), [Is].False)
            Assert.That(EzUtil.IsNotEmpty(New IntPvo(arg)), [Is].True)
        End Sub

        <TestCase("")>
        <TestCase(Nothing)>
        Public Sub 値がNullか空文字なら_真になる(arg As String)
            Assert.That(EzUtil.IsEmpty(New StrPvo(arg)), [Is].True)
            Assert.That(EzUtil.IsNotEmpty(New StrPvo(arg)), [Is].False)
        End Sub

    End Class

    Public Class IsEqualIfNull_ : Inherits EzUtilTest

        <Test()> Public Sub 型が同じで_値有有で_一致するなら_true()
            Dim s1 As String = "x"
            Dim s2 As String = "x"
            Assert.IsTrue(EzUtil.IsEqualIfNull(s1, s2), "String型")

            Dim i1 As Integer? = 11
            Dim i2 As Integer? = 11
            Assert.IsTrue(EzUtil.IsEqualIfNull(i1, i2), "Integer?型")

            Dim d1 As DateTime? = CDate("2012/03/04 05:06:07")
            Dim d2 As DateTime? = CDate("2012/03/04 05:06:07")
            Assert.IsTrue(EzUtil.IsEqualIfNull(d1, d2))
        End Sub

        <Test()> Public Sub 型が同じで_値有有で_不一致なら_false()
            Dim s1 As String = "x"
            Dim s2 As String = "y"
            Assert.IsFalse(EzUtil.IsEqualIfNull(s1, s2), "String型")

            Dim i1 As Integer? = 11
            Dim i2 As Integer? = 22
            Assert.IsFalse(EzUtil.IsEqualIfNull(i1, i2), "Integer?型")

            Dim d1 As DateTime? = CDate("2012/03/04 05:06:07")
            Dim d2 As DateTime? = CDate("2012/12/12 12:12:12")
            Assert.IsFalse(EzUtil.IsEqualIfNull(d1, d2), "DateTime?型")
        End Sub

        <Test()> Public Sub 型が同じで_値有無なら_false()
            Dim s1 As String = "x"
            Dim s2 As String = Nothing
            Assert.IsFalse(EzUtil.IsEqualIfNull(s1, s2), "String型")

            Dim i1 As Integer? = 11
            Dim i2 As Integer? = Nothing
            Assert.IsFalse(EzUtil.IsEqualIfNull(i1, i2), "Integer?型")

            Dim d1 As DateTime? = CDate("2012/03/04 05:06:07")
            Dim d2 As DateTime? = Nothing
            Assert.IsFalse(EzUtil.IsEqualIfNull(d1, d2), "DateTime?型")
        End Sub

        <Test()> Public Sub 型が同じで_値無無なら_false()
            Dim s1 As String = Nothing
            Dim s2 As String = Nothing
            Assert.IsTrue(EzUtil.IsEqualIfNull(s1, s2), "String型")

            Dim i1 As Integer? = Nothing
            Dim i2 As Integer? = Nothing
            Assert.IsTrue(EzUtil.IsEqualIfNull(i1, i2), "Integer?型")

            Dim d1 As DateTime? = Nothing
            Dim d2 As DateTime? = Nothing
            Assert.IsTrue(EzUtil.IsEqualIfNull(d1, d2), "DateTime?型")
        End Sub

    End Class

    Public Class AssertParameterIsNotEmpty_空文字や長さゼロのCollectionなら例外 : Inherits EzUtilTest

        <Test()> Public Sub String型_空文字だから例外()
            Dim value As String = ""
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.Fail("空文字が例外にならず!!")
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub String型_null値だから例外()
            Dim value As String = Nothing
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.Fail("null値が例外にならず!!")
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub String型_値アリだからok()
            Dim value As String = "a"
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.IsTrue(True)
            Catch expected As ArgumentException
                Assert.Fail("値アリが例外!!")
            End Try
        End Sub

        <Test()> Public Sub String型配列_長さゼロだから例外()
            Dim value As String() = {}
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.Fail("長さゼロ配列が例外にならず!!")
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub String型配列_null値だから例外()
            Dim value As String() = Nothing
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.Fail("null値が例外にならず!!")
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub String型配列_値アリだからok()
            Dim value As String() = {"a"}
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.IsTrue(True)
            Catch expected As ArgumentException
                Assert.Fail("値アリが例外!!")
            End Try
        End Sub

        <Test()> Public Sub String型List_長さゼロだから例外()
            Dim value As New List(Of String)
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.Fail("長さゼロListが例外にならず!!")
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub String型List_null値だから例外()
            Dim value As List(Of String) = Nothing
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.Fail("null値が例外にならず!!")
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub String型List_値アリだからok()
            Dim value As New List(Of String)(New String() {"a"})
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.IsTrue(True)
            Catch expected As ArgumentException
                Assert.Fail("値アリが例外!!")
            End Try
        End Sub

        Private Class TestingCollectionObject : Inherits CollectionObject(Of String)
            Public Sub New()
            End Sub
            Public Sub New(ByVal src As CollectionObject(Of String))
                MyBase.New(src)
            End Sub
            Public Sub New(ByVal initialList As IEnumerable(Of String))
                MyBase.New(initialList)
            End Sub
        End Class

        <Test()> Public Sub CollectionObject_長さゼロだから例外()
            Dim value As New TestingCollectionObject
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.Fail("sizeゼロが例外にならず!!")
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub CollectionObject_値アリだからok()
            Dim value As New TestingCollectionObject({"a", "xyz"})
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.IsTrue(True)
            Catch expected As ArgumentException
                Assert.Fail("値アリが例外!!")
            End Try
        End Sub

        <Test()> Public Sub PrimitiveString_長さゼロだから例外()
            Dim value As New PrimitiveString("")
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.Fail("長さゼロが例外にならず!!")
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub PrimitiveString_長さアリだからok()
            Dim value As New PrimitiveString("abc")
            Try
                EzUtil.AssertParameterIsNotEmpty(value, "value")
                Assert.IsTrue(True)
            Catch expected As ArgumentException
                Assert.Fail("値アリが例外!!")
            End Try
        End Sub

    End Class

    Private Class PrimitiveString : Inherits PrimitiveValueObject(Of String)
        Public Sub New(ByVal value As String)
            MyBase.New(value)
        End Sub
    End Class
    Private Class PrimitiveStringB : Inherits PrimitiveValueObject(Of String)
        Public Sub New(ByVal value As String)
            MyBase.New(value)
        End Sub
    End Class
    Private Class ValueComparable : Inherits ValueObject : Implements IComparable
        Public Property A() As PrimitiveString
        Public Property B() As PrimitiveStringB

        Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
            Return New Object() {A, B}
        End Function

        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            Dim other As ValueComparable = DirectCast(obj, ValueComparable)
            Return EzUtil.CombineCompare(Function() EzUtil.CompareNullsLast(A, other.A),
                                         Function() EzUtil.CompareNullsLast(B, other.B))
        End Function
    End Class

    Public Class CompareNullsLast_Null値を最後に並べるComparer値を返す : Inherits EzUtilTest

        <Test()> Public Sub Null同士なら一致_0()
            Assert.AreEqual(0, EzUtil.CompareNullsLast(Nothing, Nothing))
        End Sub

        <Test()> Public Sub どちらかがNullなら_Null側が最後にソートされる値を返す()
            Assert.AreEqual(1, EzUtil.CompareNullsLast(Nothing, 33), "> 0 だから1 - Null値のほうが大きい判定→最後にソート")
            Assert.AreEqual(-1, EzUtil.CompareNullsLast(33, Nothing), "< 0 だから-1 - Null値のほうが大きい判定→最後にソート")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート()
            Assert.AreEqual(-1, EzUtil.CompareNullsLast(33, 222), "< 0 だから-1")
            Assert.AreEqual(1, EzUtil.CompareNullsLast(222, 33), "> 0 だから1")
            Assert.AreEqual(0, EzUtil.CompareNullsLast(44, 44), "= 0")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート_Object型()
            Assert.AreEqual(-1, EzUtil.CompareNullsLast(DirectCast(33, Object), 222), "< 0 だから-1")
            Dim a As Integer? = 222
            Assert.AreEqual(1, EzUtil.CompareNullsLast(DirectCast(a, Object), 33), "> 0 だから1")
        End Sub

        <Test()> Public Sub 型が違うならエラー()
            Try
                Dim dummy As Integer = EzUtil.CompareNullsLast("44", 44)
                Assert.Fail()
            Catch ex As ArgumentException
                Assert.That(ex.Message, [Is].EqualTo("型が合っていない. x As String, y As Int32"))
            End Try
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート_String型とDate型()
            Assert.AreEqual(-1, EzUtil.CompareNullsLast("abc", "xyz"), "< 0 だから-1")
            Dim a As Integer? = 222
            Assert.AreEqual(1, EzUtil.CompareNullsLast("z", "aaa"), "> 0 だから1")
            Assert.AreEqual(0, EzUtil.CompareNullsLast(CDate("2016/02/02"), CDate("2016/02/02")), "= 0")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート_Boolean型()
            Assert.That(EzUtil.CompareNullsLast(False, False), [Is].EqualTo(0), "同じ型で同じ値だから0")
            Assert.That(EzUtil.CompareNullsLast(True, False), [Is].EqualTo(-1), "同じ型だけど値が違うので<>0")
        End Sub

        <Test()> Public Sub PrimitiveValueObjectをそのまま比較できる()
            Assert.That(EzUtil.CompareNullsLast(New PrimitiveString("ABC"), New PrimitiveString("ABC")), [Is].EqualTo(0), "同じ型で同じ値だから0")
            Assert.That(EzUtil.CompareNullsLast(New PrimitiveString("ABC"), New PrimitiveString("XYZ")), [Is].EqualTo(-1), "同じ型だけど値が違うので<>0")
        End Sub

        <Test()> Public Sub PrimitiveValueObjectでも_型が違うならエラー()
            Try
                Dim dummy As Integer = EzUtil.CompareNullsLast(New PrimitiveString("ABC"), New PrimitiveStringB("ABC"))
                Assert.Fail()
            Catch ex As ArgumentException
                Assert.That(ex.Message, [Is].EqualTo("型が合っていない. x As PrimitiveString, y As PrimitiveStringB"))
            End Try
        End Sub

        <TestCase("A", "B", "A", "B", 0, "同じ")>
        <TestCase("A1", "B1", "A2", "B2", -1, "< 0")>
        <TestCase("A2", "B2", "A1", "B1", 1, "> 0")>
        Public Sub IComparableを実装してれば_比較できる(xA As String, xB As String, yA As String, yB As String, expected As Integer, message As String)
            Dim x As New ValueComparable With {.A = New PrimitiveString(xA), .B = New PrimitiveStringB(xB)}
            Dim y As New ValueComparable With {.A = New PrimitiveString(yA), .B = New PrimitiveStringB(yB)}
            Assert.That(EzUtil.CompareNullsLast(x, y), [Is].EqualTo(expected), message)
        End Sub

        <Test()> Public Sub 実習_Null値は最後にソートする()
            Dim aList As New List(Of Integer?)
            aList.AddRange(New Integer?() {10, 3, Nothing, 20})
            aList.Sort(Function(x As Integer?, y As Integer?) EzUtil.CompareNullsLast(x, y))

            Assert.AreEqual(4, aList.Count)
            Assert.AreEqual(3, aList(0))
            Assert.AreEqual(10, aList(1))
            Assert.AreEqual(20, aList(2))
            Assert.IsNull(aList(3))
        End Sub
    End Class

    Public Class CompareDescNullsLast_Null値を最後に並べる降順Comparer値を返す : Inherits EzUtilTest

        <Test()> Public Sub Null同士なら一致_0()
            Assert.AreEqual(0, EzUtil.CompareDescNullsLast(Nothing, Nothing))
        End Sub

        <Test()> Public Sub どちらかがNullなら_Null側が最後にソートされる値を返す()
            Assert.AreEqual(1, EzUtil.CompareDescNullsLast(Nothing, 33), "> 0 だから1 - Null値のほうが大きい判定→最後にソート")
            Assert.AreEqual(-1, EzUtil.CompareDescNullsLast(33, Nothing), "< 0 だから-1 - Null値のほうが大きい判定→最後にソート")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート()
            Assert.AreEqual(1, EzUtil.CompareDescNullsLast(33, 222), "< 0 だけど降順Sortしたいから反対の1")
            Assert.AreEqual(-1, EzUtil.CompareDescNullsLast(222, 33), "> 0 だけど降順Sortしたいから反対の-1")
            Assert.AreEqual(0, EzUtil.CompareDescNullsLast(44, 44), "= 0")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート_Object型()
            Assert.AreEqual(1, EzUtil.CompareDescNullsLast(DirectCast(33, Object), 222), "< 0 だけど降順Sortしたいから反対の1")
            Dim a As Integer? = 222
            Assert.AreEqual(-1, EzUtil.CompareDescNullsLast(DirectCast(a, Object), 33), "> 0 だけど降順Sortしたいから反対の-1")
        End Sub

        <Test()> Public Sub 型が違うならエラー()
            Try
                Dim dummy As Integer = EzUtil.CompareDescNullsLast(44, "44")
                Assert.Fail()
            Catch ex As ArgumentException
                Assert.That(ex.Message, [Is].EqualTo("型が合っていない. x As Int32, y As String"))
            End Try
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート_Boolean型()
            Assert.That(EzUtil.CompareDescNullsLast(False, False), [Is].EqualTo(0), "同じ型で同じ値だから0")
            Assert.That(EzUtil.CompareDescNullsLast(True, False), [Is].EqualTo(1), "同じ型だけど値が違うので<>0")
        End Sub

        <Test()> Public Sub PrimitiveValueObjectをそのまま比較できる2()
            Assert.That(EzUtil.CompareDescNullsLast(New PrimitiveString("ABC"), New PrimitiveString("ABC")), [Is].EqualTo(0), "同じ型で同じ値だから0")
            Assert.That(EzUtil.CompareDescNullsLast(New PrimitiveString("ABC"), New PrimitiveString("XYZ")), [Is].EqualTo(1), "同じ型だけど値が違うので<>0")
        End Sub

        <Test()> Public Sub PrimitiveValueObjectでも_型が違うならエラー()
            Try
                Dim dummy As Integer = EzUtil.CompareDescNullsLast(New PrimitiveString("ABC"), New PrimitiveStringB("ABC"))
                Assert.Fail()
            Catch ex As ArgumentException
                Assert.That(ex.Message, [Is].EqualTo("型が合っていない. x As PrimitiveString, y As PrimitiveStringB"))
            End Try
        End Sub

        <TestCase("A", "B", "A", "B", 0, "同じ")>
        <TestCase("A1", "B1", "A2", "B2", 1, "< 0 だけど降順Sort")>
        <TestCase("A2", "B2", "A1", "B1", -1, "> 0 だけど降順Sort")>
        Public Sub IComparableを実装してれば_比較できる(xA As String, xB As String, yA As String, yB As String, expected As Integer, message As String)
            Dim x As New ValueComparable With {.A = New PrimitiveString(xA), .B = New PrimitiveStringB(xB)}
            Dim y As New ValueComparable With {.A = New PrimitiveString(yA), .B = New PrimitiveStringB(yB)}
            Assert.That(EzUtil.CompareDescNullsLast(x, y), [Is].EqualTo(expected), message)
        End Sub

        <Test()> Public Sub 実習_Null値は最後にソートする()
            Dim aList As New List(Of Integer?)
            aList.AddRange(New Integer?() {10, 3, Nothing, 20})
            aList.Sort(Function(x As Integer?, y As Integer?) EzUtil.CompareDescNullsLast(x, y))

            Assert.AreEqual(4, aList.Count)
            Assert.AreEqual(20, aList(0))
            Assert.AreEqual(10, aList(1))
            Assert.AreEqual(3, aList(2))
            Assert.IsNull(aList(3))
        End Sub
    End Class

    Public Class MaxInArray : Inherits EzUtilTest

        <Test()> Public Sub Integerで最大値取得()
            Assert.AreEqual(10, EzUtil.MaxInArray(5, 2, 10))
        End Sub

        <Test()> Public Sub Decimalで最大値取得()
            Assert.AreEqual(3.3D, EzUtil.MaxInArray(-1.2D, 1.2D, 3.3D))
        End Sub

        <Test()> Public Sub Doubleで最大値取得()
            Assert.AreEqual(3.3R, EzUtil.MaxInArray(-1.2R, 1.2R, 3.3R))
        End Sub
    End Class

    Public Class ParseAsBoolean_Boolean値へ変換 : Inherits EzUtilTest

        <Test()> Public Sub Learning_TryParse_不正値ならResultもValueもfalse()
            Dim value As Boolean
            Dim result As Boolean = Boolean.TryParse("a", value)
            Assert.IsFalse(result, "変換できないからfalse")
            Assert.IsFalse(value)
        End Sub

        <Test()> Public Sub Learning_TryParse_良値ならResult_true()
            Dim value As Boolean
            Dim result As Boolean = Boolean.TryParse("false", value)
            Assert.IsTrue(result, "変換できたからtrue")
            Assert.IsFalse(value, "変換結果はfalse")
        End Sub

        <Test()> Public Sub Learning_TryParse_良値ならResult_true_valueもtrue()
            Dim value As Boolean
            Dim result As Boolean = Boolean.TryParse("true", value)
            Assert.IsTrue(result, "変換できたからtrue")
            Assert.IsTrue(value, "変換結果もtrue")
        End Sub

        <Test()> Public Sub nullなら変換できないから結果もnull()
            Dim result As Boolean? = EzUtil.ParseAsBoolean(Nothing)
            Assert.IsFalse(result.HasValue)
        End Sub

        <Test()> Public Sub 文字列trueなら結果true()
            Dim result As Boolean? = EzUtil.ParseAsBoolean("true")
            Assert.IsTrue(result.HasValue)
            Assert.IsTrue(result.Value)
        End Sub

        <Test()> Public Sub 文字列falseなら結果false()
            Dim result As Boolean? = EzUtil.ParseAsBoolean("false")
            Assert.IsTrue(result.HasValue)
            Assert.IsFalse(result.Value)
        End Sub

        <Test()> Public Sub 文字列不正値なら結果null()
            Dim result As Boolean? = EzUtil.ParseAsBoolean("hoge")
            Assert.IsFalse(result.HasValue)
        End Sub

        <Test()> Public Sub 数値1なら結果true()
            Dim result As Boolean? = EzUtil.ParseAsBoolean(1)
            Assert.IsTrue(result.HasValue)
            Assert.IsTrue(result.Value)
        End Sub

        <Test()> Public Sub 数値0なら結果false()
            Dim result As Boolean? = EzUtil.ParseAsBoolean(0)
            Assert.IsTrue(result.HasValue)
            Assert.IsFalse(result.Value)
        End Sub

    End Class

    Public Class IsEqualIfEpsilon_Single : Inherits EzUtilTest

        <Test()> Public Sub 浮動少数点型の誤差()

            Assert.That(0.1F = 1.0F - 0.9F, [Is].False, "① 値直書きでもfalse")

            Dim a As Single = 0.1F
            Dim b As Single = 1.0F - 0.9F
            Assert.That(a = b, [Is].False, "② 変数に入れて計算してもfalse")

            Assert.That(EzUtil.IsEqualIfEpsilon(a, b), [Is].True, "③ 上記②で不一致でもEpsilon考慮の比較でtrue")

            Assert.That(EzUtil.IsEqualIfEpsilon(1.0F - 0.9F, 0.1F), [Is].True, "④ 念の為、①でもtrue")
        End Sub

        <Test(), Sequential()> Public Sub 同じ値同士だからtrue( _
                <Values(0, 1, 3.14159F)> ByVal a As Single, _
                <Values(0.0F, 1.0F, 3.14159F)> ByVal b As Single)
            Assert.That(EzUtil.IsEqualIfEpsilon(a, b), [Is].True)
        End Sub

        <Test(), Sequential()> Public Sub 違う値同士だからfalse( _
                <Values(0, 1.000002F, 3.14159F)> ByVal a As Single, _
                <Values(0.1F, 1.000001F, 3.1415925F)> ByVal b As Single)
            Assert.That(EzUtil.IsEqualIfEpsilon(a, b), [Is].False)
        End Sub

    End Class

    Public Class IsEqualIfEpsilon_Double : Inherits EzUtilTest
        <Test()> Public Sub 浮動少数点型の誤差()
            Assert.That(0.1R + 0.2R = 0.3R, [Is].False, "① 値直書きでもfalse")

            Dim a As Double = 0.1R
            Dim b As Double = 0.2R
            Assert.That(a + b = 0.3R, [Is].False, "② 変数に入れて計算してもfalse")

            Assert.That(EzUtil.IsEqualIfEpsilon(a + b, 0.3R), [Is].True, "③ 上記②で不一致でもEpsilon考慮の比較でtrue")
            Assert.That(EzUtil.IsEqualIfEpsilon(0.1R + 0.2R, 0.3R), [Is].True, "④ 念の為、①でもtrue")
            Assert.That(EzUtil.IsEqualIfEpsilonStrictly(a + b, 0.3R), [Is].False)
            Assert.That(EzUtil.IsEqualIfEpsilonStrictly(0.1R + 0.2R, 0.3R), [Is].False)
        End Sub

        <Test(), Sequential()> Public Sub 同じ値同士だからtrue( _
                <Values(0, 1, 3.14159R)> ByVal a As Double, _
                <Values(0.0R, 1.0R, 3.14159R)> ByVal b As Double)
            Assert.That(EzUtil.IsEqualIfEpsilon(a, b), [Is].True)
            Assert.That(EzUtil.IsEqualIfEpsilonStrictly(a, b), [Is].True)
        End Sub

        <Test(), Sequential()> Public Sub 違う値同士だからfalse( _
                <Values(0.1 + 0.2, 1.000002, 3.14159R)> ByVal a As Double, _
                <Values(0.31R, 1.000001R, 3.141592R)> ByVal b As Double)
            Assert.That(EzUtil.IsEqualIfEpsilon(a, b), [Is].False)
            Assert.That(EzUtil.IsEqualIfEpsilonStrictly(a, b), [Is].False)
        End Sub
    End Class

    Public Class GetDecimalSize_小数点以下桁数を返す : Inherits EzUtilTest

        <Test(), Sequential()> Public Sub 各値はexpectedの通りになる(<Values(1.0R, 2.0R, 3.14R, 4.0000001R, 0.0R)> ByVal d As Double, _
                                                            <Values(0, 0, 2, 7, 0)> ByVal expected As Integer)
            Assert.That(EzUtil.GetDecimalSize(d), [Is].EqualTo(expected))
        End Sub

        <Test(), Sequential()> Public Sub 各値はexpectedの通りになる(<Values("1", "2", "3.14", "4.0000001", "0.00")> ByVal d As String, _
                                                            <Values(0, 0, 2, 7, 0)> ByVal expected As Integer)
            Assert.That(EzUtil.GetDecimalSize(CDec(d)), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class ConvValueToEnum_ : Inherits EzUtilTest

        Public Enum Hoge
            ONE
            Two
            three
        End Enum

        <Test(), Sequential()> Public Sub Enumの数値を変換できる(<Values(0, 1, 2)> ByVal value As Integer, _
                                                            <Values(Hoge.ONE, Hoge.Two, Hoge.three)> ByVal expected As Hoge)
            Assert.That(EzUtil.ConvValueToEnum(Of Hoge)(value), [Is].EqualTo(expected))
        End Sub

        <Test(), Sequential()> Public Sub Enumの値名を変換できる(<Values("ONE", "Two", "three")> ByVal value As String, _
                                                            <Values(Hoge.ONE, Hoge.Two, Hoge.three)> ByVal expected As Hoge)
            Assert.That(EzUtil.ConvNameToEnum(Of Hoge)(value), [Is].EqualTo(expected))
        End Sub

        <Test(), ExpectedException(ExpectedException:=GetType(ArgumentException)), Sequential()> _
        Public Sub Enumに指定値がなければ例外(<Values(-1, 3, 99)> ByVal value As Integer)
            Const DUMMY As Integer = 99
            Assert.That(EzUtil.ConvValueToEnum(Of Hoge)(value), [Is].EqualTo(DUMMY))
        End Sub

        <Test(), ExpectedException(ExpectedException:=GetType(ArgumentException)), Sequential()> _
        Public Sub Enumに値名がなければ例外_大文字小文字違いも例外(<Values("one", "tWO", "THREE", "FOUR")> ByVal value As String)
            Const DUMMY As Integer = 99
            Assert.That(EzUtil.ConvNameToEnum(Of Hoge)(value), [Is].EqualTo(DUMMY))
        End Sub

    End Class

    Public Class ConvDblToDec_Double値をDecimal値に変換する : Inherits EzUtilTest
        <Test()> Public Sub Decimalに変換できる()
            Assert.That(EzUtil.ConvDblToDec(1.234567890123E-30), [Is].EqualTo(0D))
            Assert.That(EzUtil.ConvDblToDec(1.2345678901234E-25), [Is].EqualTo(0.0000000000000000000000001235D))
            Assert.That(EzUtil.ConvDblToDec(1.2345678901234567), [Is].EqualTo(1.23456789012346D))
            Assert.That(EzUtil.ConvDblToDec(1234567890123.4568), [Is].EqualTo(1234567890123.46D))
            Assert.That(EzUtil.ConvDblToDec(1.2345678901234568E+28), [Is].EqualTo(12345678901234600000000000000D))
        End Sub

        <Test()> Public Sub Decimal値に変換できずオーバーフローする場合_例外がスローされる()
            Try
                Dim dec As Decimal = EzUtil.ConvDblToDec(Double.MaxValue)
                Assert.Fail("例外がスローされなかった")
            Catch ex As OverflowException
                Assert.Pass("OverflowExceptionが投げられた")
            End Try
        End Sub

        <Test()> Public Sub Decimal値に変換できずオーバーフローする場合_例外がスローされる_計算処理結果の場合()
            Dim test As Double = Math.Pow(100000, 10)
            Try
                Dim dec As Decimal = EzUtil.ConvDblToDec(test)
                Assert.Fail("例外がスローされなかった")
            Catch ex As OverflowException
                Assert.Pass("OverflowExceptionが投げられた")
            End Try
        End Sub
    End Class

    Private Class TestingVo
        Public Id As Integer?
        Public Name As String
    End Class

    Public Class JudgeCompare_比較処理を順次実行する : Inherits EzUtilTest

        <Test()> Public Sub Hoge()
            Dim lst As List(Of TestingVo) = EzUtil.NewList(New TestingVo With {.Id = 1, .Name = "ONE"}, _
                                                           New TestingVo With {.Id = 2, .Name = "TWO"})
            lst.Sort(Function(x, y) EzUtil.JudgeCompare(Function() EzUtil.CompareNullsLast(x.Id, y.Id)))
            Assert.That(lst(0).Id, [Is].EqualTo(1))
            Assert.That(lst(1).Id, [Is].EqualTo(2))
        End Sub

        <Test()> Public Sub Hoge2()
            Dim lst As List(Of TestingVo) = EzUtil.NewList(New TestingVo With {.Id = 1, .Name = "ONE"}, _
                                                           New TestingVo With {.Id = 2, .Name = "TWO"})
            lst.Sort(Function(x, y) EzUtil.JudgeCompare(Function() StringUtil.CompareDescNullsLast(x.Name, y.Name)))
            Assert.That(lst(0).Name, [Is].EqualTo("TWO"))
            Assert.That(lst(1).Name, [Is].EqualTo("ONE"))
        End Sub

        <Test()> Public Sub Hoge33()
            Dim lst As List(Of TestingVo) = EzUtil.NewList(New TestingVo With {.Id = 1, .Name = "C"}, _
                                                           New TestingVo With {.Id = 2, .Name = "B"}, _
                                                           New TestingVo With {.Id = 1, .Name = "A"})
            lst.Sort(Function(x, y) EzUtil.JudgeCompare(Function() EzUtil.CompareNullsLast(x.Id, y.Id), _
                                                        Function() StringUtil.CompareNullsLast(x.Name, y.Name)))
            Assert.That(lst(0).Name, [Is].EqualTo("A"))
            Assert.That(lst(1).Name, [Is].EqualTo("C"))
            Assert.That(lst(2).Name, [Is].EqualTo("B"))
        End Sub

    End Class

    Public Class PowerTest_Decimal型でべき乗計算 : Inherits EzUtilTest

        <Test(), Sequential()> Public Sub _10のべき乗を計算できる(<Values(2, 1, 0, -1, -2)> ByVal exponent As Integer, _
                                                        <Values("100", "10", "1", "0.1", "0.01")> ByVal expected As String)
            Assert.That(EzUtil.Power(10D, exponent), [Is].EqualTo(CDec(expected)))
        End Sub

        <Test(), Sequential()> Public Sub _2のべき乗を計算できる(<Values(16, 1, 0, -1, -2)> ByVal exponent As Integer, _
                                                        <Values("65536", "2", "1", "0.5", "0.25")> ByVal expected As String)
            Assert.That(EzUtil.Power(2D, exponent), [Is].EqualTo(CDec(expected)))
        End Sub

    End Class

    Public Class CopyForSerializableTest : Inherits EzUtilTest
#Region "テストで利用するVo"
        <Serializable()> Protected Class TestVo
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
#End Region

        <Test()> Public Sub Integer型をコピーできる()
            Dim testee As Integer = 123
            Const EXPECTED As Integer = 123
            Dim actual As Integer = EzUtil.CopyForSerializable(testee)
            Assert.That(actual, [Is].EqualTo(EXPECTED))
            testee = 124
            Assert.That(actual, [Is].Not.EqualTo(testee))
            Assert.That(actual, [Is].EqualTo(EXPECTED))
        End Sub

        <Test()> Public Sub String型をコピーできる()
            Dim testee As String = "あいう"
            Const EXPECTED As String = "あいう"
            Dim actual As String = EzUtil.CopyForSerializable(testee)
            Assert.That(actual, [Is].EqualTo(EXPECTED))
            testee = "あいうえ"
            Assert.That(actual, [Is].Not.EqualTo(testee))
            Assert.That(actual, [Is].EqualTo(EXPECTED))
        End Sub

        <Test()> Public Sub 配列をコピーできる()
            Dim testee As String() = New String() {"A", "B"}
            Dim actual As String() = EzUtil.CopyForSerializable(testee)
            Assert.False(actual Is testee)
            Assert.That(actual.Count, [Is].EqualTo(2))
            Assert.That(actual(0), [Is].EqualTo("A"))
            Assert.That(actual(1), [Is].EqualTo("B"))
        End Sub

        <Test()> Public Sub オブジェクトをディープコピーできる()
            Dim testee As Object() = New Object() {1, "a", New String() {"あ", "い"}}
            Dim actual As Object() = EzUtil.CopyForSerializable(testee)
            Assert.False(actual Is testee)
            Assert.That(actual.Count, [Is].EqualTo(3))
            Assert.That(actual(0), [Is].EqualTo(1))
            Assert.That(actual(1), [Is].EqualTo("a"))
            Dim actuals As String() = DirectCast(actual(2), String())
            Assert.That(actuals(0), [Is].EqualTo("あ"))
            Assert.That(actuals(1), [Is].EqualTo("い"))
        End Sub

        <Test()> Public Sub シリアライズ可能であれば_自作Voもコピーできる()
            Dim testee As New TestVo
            With testee
                .ABoolean = False
                .ADateTime = New Date(2016, 11, 14)
                .ADecimal = 123D
                .ADouble = Nothing
                .AInt = 100
                .AIntArray = New Integer?() {1, Nothing, 3}
                .AString = "あ"
                .AStringArray = New String() {"A", "B"}
            End With

            Dim actual As TestVo = EzUtil.CopyForSerializable(testee)
            Assert.That(actual.ABoolean, [Is].EqualTo(False))
            Assert.That(actual.ABooleanArray, [Is].EqualTo(Nothing))
            Assert.That(actual.ADateTime, [Is].EqualTo(New Date(2016, 11, 14)))
            Assert.That(actual.ADecimal, [Is].EqualTo(123D))
            Assert.That(actual.ADouble, [Is].EqualTo(Nothing))
            Assert.That(actual.AInt, [Is].EqualTo(100))
            Assert.That(actual.AIntArray.Count, [Is].EqualTo(3))
            Assert.That(actual.AIntArray(0), [Is].EqualTo(1))
            Assert.That(actual.AIntArray(1), [Is].EqualTo(Nothing))
            Assert.That(actual.AIntArray(2), [Is].EqualTo(3))
            Assert.That(actual.AString, [Is].EqualTo("あ"))
            Assert.That(actual.AStringArray.Count, [Is].EqualTo(2))
            Assert.That(actual.AStringArray(0), [Is].EqualTo("A"))
            Assert.That(actual.AStringArray(1), [Is].EqualTo("B"))
        End Sub
    End Class

    Public Class CallIfTest : Inherits EzUtilTest

        <Test()> Public Sub Func_Nullだから_処理しない_実行時エラーにならない()
            Dim actual As Object = EzUtil.CallIf(Nothing, Function(o) o.ToString)
            Assert.That(actual, [Is].Null)
        End Sub

        <Test()> Public Sub Func_有効値だから_処理する()
            Dim actual As String = EzUtil.CallIf(123, Function(o) o.ToString)
            Assert.That(actual, [Is].EqualTo("123"))
        End Sub

        <Test()> Public Sub Sub_Nullなら処理しない()
            Dim argIfNull As Integer? = Nothing
            Dim actual As String = "A"
            EzUtil.CallIf(argIfNull, Sub(o) actual = o.ToString)
            Assert.That(actual, [Is].EqualTo("A"))
        End Sub

        <Test()> Public Sub Sub_有効値だから_処理する()
            Dim argIfNull As Integer? = 123
            Dim actual As String = "A"
            EzUtil.CallIf(argIfNull, Sub(o) actual = o.ToString)
            Assert.That(actual, [Is].EqualTo("123"))
        End Sub

        <Test()> Public Sub Func_Nullだから_処理しない_実行時エラーにならない_callback引数なし()
            Dim argIfNull As Integer? = Nothing
            Dim actual As Object = EzUtil.CallIf(argIfNull, Function() argIfNull.ToString)
            Assert.That(actual, [Is].Null)
        End Sub

        <Test()> Public Sub Func_有効値だから_処理する_callback引数なし()
            Dim argIfNull As Integer? = 123
            Dim actual As String = EzUtil.CallIf(argIfNull, Function() argIfNull.ToString)
            Assert.That(actual, [Is].EqualTo("123"))
        End Sub

        <Test()> Public Sub Sub_Nullなら処理しない_callback引数なし()
            Dim argIfNull As Integer? = Nothing
            Dim actual As String = "A"
            EzUtil.CallIf(argIfNull, Sub() actual = argIfNull.ToString)
            Assert.That(actual, [Is].EqualTo("A"))
        End Sub

        <Test()> Public Sub Sub_有効値だから_処理する_callback引数なし()
            Dim argIfNull As Integer? = 123
            Dim actual As String = "A"
            EzUtil.CallIf(argIfNull, Sub() actual = argIfNull.ToString)
            Assert.That(actual, [Is].EqualTo("123"))
        End Sub

    End Class

    Public Class GenerateHashSHA256Test : Inherits EzUtilTest

        <TestCase("ABC", "b5d4045c3f466fa91fe2cc6abe79232a1a57cdf104f7a26e716e0a1e2789df78")>
        <TestCase("20191212", "3ff174a83a1376730c5dc013a672b09459b7c480ea2fe4e9fea5e7a040ef0f9d")>
        <TestCase("daniel", "bd3dae5fb91f88a4f0978222dfd58f59a124257cb081486387cbae9df11fb879")>
        Public Sub SHA256ハッシュ値を生成できる(arg As String, expected As String)
            Dim actual As String = EzUtil.GenerateHashSHA256(arg)
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

    End Class

End Class
