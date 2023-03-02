Imports NUnit.Framework

Public MustInherit Class EnumUtilTest

    Public Enum TestingBit
        None = &H0
        Bit1 = &H1
        Bit2 = &H2
        Bit3 = &H4
        Bit4 = &H8
        Bit5 = &H10
    End Enum

    <Flags()> Public Enum TestingFlags
        None = &H0
        Fit1 = &H1
        Fit2 = &H2
        Fit3 = &H4
        Fit4 = &H8
        Fit5 = &H10
    End Enum

    Public Class ParseByNameTest : Inherits EnumUtilTest

        <TestCase("Bit1", TestingBit.Bit1)>
        <TestCase("None", TestingBit.None)>
        <TestCase("None", TestingFlags.None)>
        <TestCase("Fit3", TestingFlags.Fit3)>
        Public Sub TestParseEnumName(Of TEnum As Structure)(enumName As String, expected As TEnum)
            Assert.That(EnumUtil.ParseByName(Of TEnum)(enumName), [Is].EqualTo(expected))
        End Sub

        <TestCase("BIT1")>
        <TestCase("bit5")>
        <TestCase("NOne")>
        Public Sub 大文字小文字が一致しないとエラーになる(enumName As String)
            Try
                EnumUtil.ParseByName(Of TestingBit)(enumName)
                Assert.Fail()
            Catch expected As Exception
                Assert.That(expected.Message, [Is].EqualTo(String.Format("Enum {0} に'{1}'はありません", GetType(TestingBit).Name, enumName)))
            End Try
        End Sub

        <TestCase("BIT1", TestingBit.Bit1)>
        <TestCase("bit5", TestingBit.Bit5)>
        <TestCase("NOne", TestingBit.None)>
        <TestCase("NOne", TestingFlags.None)>
        <TestCase("fiT3", TestingFlags.Fit3)>
        Public Sub IgnoreCase_true_なら大文字小文字が一致しなくても_同一視する(Of TEnum As Structure)(enumName As String, expected As TEnum)
            Assert.That(EnumUtil.ParseByName(Of TEnum)(enumName, ignoresCase:=True), [Is].EqualTo(expected))
        End Sub

        <TestCase("bit10", True)>
        <TestCase("bit10", False)>
        <TestCase("abc", True)>
        <TestCase("abc", False)>
        Public Sub 存在しない名前はエラーになる(enumName As String, ignoreCase As Boolean)
            Try
                EnumUtil.ParseByName(Of TestingBit)(enumName, ignoreCase)
                Assert.Fail()
            Catch expected As Exception
                Assert.That(expected.Message, [Is].EqualTo(String.Format("Enum {0} に'{1}'はありません", GetType(TestingBit).Name, enumName)))
            End Try
        End Sub

    End Class

    Public Class ParseByNullableValueTest : Inherits EnumUtilTest
        Public Enum MyEnum
            ZERO = 0
            ONE
            TWO
        End Enum

        <TestCase(0, MyEnum.ZERO)>
        <TestCase(1, MyEnum.ONE)>
        <TestCase(2, MyEnum.TWO)>
        Public Sub Parseできる(enumValue As Integer, expected As MyEnum)
            Assert.That(EnumUtil.ParseByNullableValue(Of MyEnum)(enumValue), [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub null値ならnullのEnumになる()
            Dim enumValue As Integer? = Nothing
            Assert.That(EnumUtil.ParseByNullableValue(Of MyEnum)(enumValue), [Is].Null)
        End Sub

        <TestCase(-1)>
        <TestCase(3)>
        <TestCase(10)>
        Public Sub Enumに未定義なら例外になる(enumValue As Integer)
            Try
                EnumUtil.ParseByNullableValue(Of MyEnum)(enumValue)
                Assert.Fail()
            Catch ex As Exception
                Assert.That(ex.Message, [Is].EqualTo("Enum MyEnum に " & enumValue & " はありません"))
            End Try
        End Sub

        <Test()> Public Sub Enumに未定義の値でも_ignoresDefineがtrueなら_例外にしない()
            Dim actual As MyEnum? = EnumUtil.ParseByNullableValue(Of MyEnum)(10, ignoresDefine:=True)
            Assert.That(actual.HasValue, [Is].True)
            Assert.That(actual, [Is].TypeOf(GetType(MyEnum)))
        End Sub

    End Class

    Public Class GetEnumNameTest : Inherits EnumUtilTest

        <TestCase(TestingBit.Bit1, "Bit1")>
        <TestCase(TestingFlags.Fit4, "Fit4")>
        Public Sub Enumの名前を取得できる(Of TEnum As Structure)(enumValue As TEnum, name As String)
            Assert.That(EnumUtil.GetName(Of TEnum)(enumValue), [Is].EqualTo(name))
        End Sub

        <Test()> Public Sub 合成した値からは_取れない()
            Dim combineValue As TestingBit = TestingBit.Bit1 Or TestingBit.Bit3
            Assert.That(EnumUtil.GetName(Of TestingBit)(combineValue), [Is].Null)
            Assert.That(combineValue.ToString, [Is].EqualTo("5"))
        End Sub

        <Test()> Public Sub 合成した値からは_取れない2()
            Dim combineValue As TestingFlags = TestingFlags.Fit2 Or TestingFlags.Fit5
            Assert.That(EnumUtil.GetName(Of TestingBit)(combineValue), [Is].Null)
            Assert.That(combineValue.ToString, [Is].EqualTo("Fit2, Fit5"), "Flags属性が付いてると取れる")
        End Sub

    End Class

    Public Class GetEnumNamesTest : Inherits EnumUtilTest

        <Test()> Public Sub Enum型の定義した名前を取得できる()
            Assert.That(EnumUtil.GetNames(Of TestingFlags), [Is].EquivalentTo({"None", "Fit1", "Fit2", "Fit3", "Fit4", "Fit5"}))
        End Sub

    End Class

    Public Class ContainsEnumTest : Inherits EnumUtilTest

        <TestCase(TestingBit.Bit1, True)>
        <TestCase(TestingBit.Bit2, False)>
        <TestCase(TestingBit.Bit3, True)>
        Public Sub 結合した値に含まれるか判定できる(value As TestingBit, expected As Boolean)
            Const TESING_VALUE As TestingBit = TestingBit.Bit1 Or TestingBit.Bit3
            Assert.That((TESING_VALUE And value) = value, [Is].EqualTo(expected))
            Assert.That(EnumUtil.Contains(TESING_VALUE, value), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class GetValuesTest : Inherits EnumUtilTest
        Private Enum TestEnum
            AAA
            BBB
            CCC
        End Enum

        Private Structure HogeStructure
            Public Const PIYO As Integer = 0
            Public Const FUGA As Integer = 1
        End Structure

        <Test()>
        Public Sub 定義済みのEnum値をまとめて取得できる()
            Dim values As TestEnum() = EnumUtil.GetValues(Of TestEnum)()
            Assert.That(values, [Is].EquivalentTo([Enum].GetValues(GetType(TestEnum))))
        End Sub

        <Test()>
        Public Sub Enum値以外の型が指定されてたら例外を返す()
            Try
                EnumUtil.GetValues(Of HogeStructure)()
                Assert.Fail()
            Catch expected As ArgumentException
                Assert.That(expected.Message, [Is].EqualTo("型引数:HogeStructureはEnum型ではない"))
            End Try
        End Sub

    End Class

End Class
