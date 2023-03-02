Imports Fhi.Fw.Domain
Imports NUnit.Framework

Public MustInherit Class CompareUtilTest

#Region "TestPVO"

    Private Class StrPvo : Inherits PrimitiveValueObject(Of String)
        Public Sub New(value As String)
            MyBase.new(value)
        End Sub
    End Class

    Private Class IntPvo : Inherits PrimitiveValueObject(Of Integer)
        Public Sub New(value As Integer)
            MyBase.new(value)
        End Sub
    End Class

#End Region

    Public Class CompareTest : Inherits CompareUtilTest

        <TestCase("当アk", "当アk", CompareUtil.CompareResult.Equal)>
        <TestCase("123", "124", CompareUtil.CompareResult.GreaterThan)>
        <TestCase("ABC", "ABB", CompareUtil.CompareResult.LessThan)>
        Public Sub 文字列で比較できる(ls As String, rs As String, expected As CompareUtil.CompareResult)
            Dim actual As CompareUtil.CompareResult = CompareUtil.Compare(ls, rs)
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

        <TestCase(123, 123, CompareUtil.CompareResult.Equal)>
        <TestCase(99, 100, CompareUtil.CompareResult.GreaterThan)>
        <TestCase(12345, 12344, CompareUtil.CompareResult.LessThan)>
        Public Sub 整数値で比較できる(ls As Integer, rs As Integer, expected As CompareUtil.CompareResult)
            Dim actual As CompareUtil.CompareResult = CompareUtil.Compare(ls, rs)
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

        <TestCase(111.5, 111.5, CompareUtil.CompareResult.Equal)>
        <TestCase(999.0, 999.1, CompareUtil.CompareResult.GreaterThan)>
        <TestCase(123.5555, 123.5554, CompareUtil.CompareResult.LessThan)>
        Public Sub 固定小数点数値で比較できる(ls As Decimal, rs As Decimal, expected As CompareUtil.CompareResult)
            Dim actual As CompareUtil.CompareResult = CompareUtil.Compare(ls, rs)
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

        <TestCase(333.25F, 333.25F, CompareUtil.CompareResult.Equal)>
        <TestCase(1234.4F, 1234.5F, CompareUtil.CompareResult.GreaterThan)>
        <TestCase(99.125F, 99.124F, CompareUtil.CompareResult.LessThan)>
        Public Sub 単精度浮動小数点数値で比較できる(ls As Single, rs As Single, expected As CompareUtil.CompareResult)
            Dim actual As CompareUtil.CompareResult = CompareUtil.Compare(ls, rs)
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

        <TestCase(333.25, 333.25, CompareUtil.CompareResult.Equal)>
        <TestCase(1234.4, 1234.5, CompareUtil.CompareResult.GreaterThan)>
        <TestCase(99.125, 99.124, CompareUtil.CompareResult.LessThan)>
        Public Sub 倍精度浮動小数点数値で比較できる(ls As Double, rs As Double, expected As CompareUtil.CompareResult)
            Dim actual As CompareUtil.CompareResult = CompareUtil.Compare(ls, rs)
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

        <TestCase("2012/03/04", "2012/03/04", CompareUtil.CompareResult.Equal)>
        <TestCase("2015/12/12 15:30:00", "2015/12/12 15:30:00", CompareUtil.CompareResult.Equal)>
        <TestCase("2013/11/11", "2013/11/12", CompareUtil.CompareResult.GreaterThan)>
        <TestCase("2016/03/01 16:59:59", "2016/03/01 17:00:00", CompareUtil.CompareResult.GreaterThan)>
        <TestCase("2017/05/20", "2017/05/19", CompareUtil.CompareResult.LessThan)>
        <TestCase("2019/09/19 07:00:00", "2019/09/19 06:59:59", CompareUtil.CompareResult.LessThan)>
        Public Sub 日付で比較できる(ls As String, rs As String, expected As CompareUtil.CompareResult)
            Dim lsDate As Date = DateUtil.ConvDateValueToDateTime(ls).Value
            Dim rsDate As Date = DateUtil.ConvDateValueToDateTime(rs).Value
            Dim actual As CompareUtil.CompareResult = CompareUtil.Compare(lsDate, rsDate)
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

        <TestCase("もじ", "もじ", CompareUtil.CompareResult.Equal)>
        <TestCase("あああ", "ああい", CompareUtil.CompareResult.GreaterThan)>
        <TestCase("ううう", "ううい", CompareUtil.CompareResult.LessThan)>
        Public Sub PVOなら型を解釈して比較できる_文字列(ls As String, rs As String, expected As CompareUtil.CompareResult)
            Dim actual As CompareUtil.CompareResult = CompareUtil.Compare(New StrPvo(ls), New StrPvo(rs))
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

        <TestCase(123, 123, CompareUtil.CompareResult.Equal)>
        <TestCase(998, 999, CompareUtil.CompareResult.GreaterThan)>
        <TestCase(-15, -16, CompareUtil.CompareResult.LessThan)>
        Public Sub PVOなら型を解釈して比較できる_数値(ls As Integer, rs As Integer, expected As CompareUtil.CompareResult)
            Dim actual As CompareUtil.CompareResult = CompareUtil.Compare(New IntPvo(ls), New IntPvo(rs))
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

        <TestCase("", Nothing)>
        <TestCase(Nothing, "")>
        <TestCase(Nothing, Nothing)>
        Public Sub 引数がNULLなら_例外(ls As String, rs As String)
            Assert.That(Sub() CompareUtil.Compare(ls, rs),
                        Throws.TypeOf(GetType(ArgumentNullException)).And.Message.EqualTo("値を Null にすることはできません。").Or.EqualTo("Value cannot be null."))
        End Sub

        <Test()>
        Public Sub 非対応の型なら_例外()
            Assert.That(Sub() CompareUtil.Compare(True, True),
                        Throws.ArgumentException.And.Message.EqualTo("System.Boolean型は対応していない").Or.EqualTo("Value cannot be null."))
        End Sub

        <Test()> Public Sub 左辺と右辺で異なる型を渡すと_例外()
            Assert.That(Sub() CompareUtil.Compare("123", 123),
                        Throws.ArgumentException.And.Message.EqualTo("左辺と右辺の型は同じでないとダメ"))
        End Sub

        <Test> Public Sub PVOの型を無視して_内部の値とで比較できる()
            Assert.That(CompareUtil.Compare(New StrPvo("123"), "123", ignorePvo:=True), [Is].EqualTo(CompareUtil.CompareResult.Equal))
        End Sub

        <Test> Public Sub PVOの型を無視するといえど_内部値の型が違ければ例外()
            Assert.That(Sub() CompareUtil.Compare(New StrPvo("123"), 123, ignorePvo:=True),
                        Throws.ArgumentException.And.Message.EqualTo("左辺と右辺の型は同じでないとダメ"))
        End Sub
    End Class

    Public Class IsEqualTest : Inherits CompareUtilTest

        <TestCase("ほげ")>
        <TestCase(123)>
        <TestCase(33.3)>
        Public Sub 等しければTrue(value As Object)
            Assert.True(CompareUtil.IsEqual(value, value))
        End Sub

        <TestCase("ほげ", "ふが")>
        <TestCase(123, -123)>
        <TestCase(33.3, 33.4)>
        Public Sub 等しくなければFalse(ls As Object, rs As Object)
            Assert.False(CompareUtil.IsEqual(ls, rs))
        End Sub
    End Class

    Public Class IsGreaterThanTest : Inherits CompareUtilTest

        <TestCase("ほげ")>
        <TestCase(123)>
        <TestCase(33.3)>
        Public Sub 等しければFalse(value As Object)
            Assert.False(CompareUtil.IsGreaterThan(value, value))
        End Sub

        <TestCase("いい", "いう")>
        <TestCase(123, 124)>
        <TestCase(33.3, 33.4)>
        Public Sub 右辺のほうが大きければTrue(ls As Object, rs As Object)
            Assert.True(CompareUtil.IsGreaterThan(ls, rs))
        End Sub

        <TestCase("いい", "いあ")>
        <TestCase(123, 122)>
        <TestCase(33.3, 33.2)>
        Public Sub 右辺のほうが小さければFalse(ls As Object, rs As Object)
            Assert.False(CompareUtil.IsGreaterThan(ls, rs))
        End Sub
    End Class

    Public Class IsLessThanTest : Inherits CompareUtilTest

        <TestCase("ほげ")>
        <TestCase(123)>
        <TestCase(33.3)>
        Public Sub 等しければFalse(value As Object)
            Assert.False(CompareUtil.IsLessThan(value, value))
        End Sub

        <TestCase("いい", "いう")>
        <TestCase(123, 124)>
        <TestCase(33.3, 33.4)>
        Public Sub 右辺のほうが大きければFalse(ls As Object, rs As Object)
            Assert.False(CompareUtil.IsLessThan(ls, rs))
        End Sub

        <TestCase("いい", "いあ")>
        <TestCase(123, 122)>
        <TestCase(33.3, 33.2)>
        Public Sub 右辺のほうが小さければTrue(ls As Object, rs As Object)
            Assert.True(CompareUtil.IsLessThan(ls, rs))
        End Sub
    End Class

    Public Class IsGreaterEqualTest : Inherits CompareUtilTest

        <TestCase("ほげ")>
        <TestCase(123)>
        <TestCase(33.3)>
        Public Sub 等しければTrue(value As Object)
            Assert.True(CompareUtil.IsGreaterEqual(value, value))
        End Sub

        <TestCase("いい", "いう")>
        <TestCase(123, 124)>
        <TestCase(33.3, 33.4)>
        Public Sub 右辺のほうが大きければTrue(ls As Object, rs As Object)
            Assert.True(CompareUtil.IsGreaterEqual(ls, rs))
        End Sub

        <TestCase("いい", "いあ")>
        <TestCase(123, 122)>
        <TestCase(33.3, 33.2)>
        Public Sub 右辺のほうが小さければFalse(ls As Object, rs As Object)
            Assert.False(CompareUtil.IsGreaterEqual(ls, rs))
        End Sub
    End Class

    Public Class IsLessEqualTest : Inherits CompareUtilTest

        <TestCase("ほげ")>
        <TestCase(123)>
        <TestCase(33.3)>
        Public Sub 等しければTrue(value As Object)
            Assert.True(CompareUtil.IsLessEqual(value, value))
        End Sub

        <TestCase("いい", "いう")>
        <TestCase(123, 124)>
        <TestCase(33.3, 33.4)>
        Public Sub 右辺のほうが大きければFalse(ls As Object, rs As Object)
            Assert.False(CompareUtil.IsLessEqual(ls, rs))
        End Sub

        <TestCase("いい", "いあ")>
        <TestCase(123, 122)>
        <TestCase(33.3, 33.2)>
        Public Sub 右辺のほうが小さければTrue(ls As Object, rs As Object)
            Assert.True(CompareUtil.IsLessEqual(ls, rs))
        End Sub
    End Class

End Class
