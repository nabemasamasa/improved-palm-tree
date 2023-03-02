Imports System.Collections.Generic
Imports NUnit.Framework
Imports System.Threading
Imports Fhi.Fw.Domain

'''<summary>
'''This is a test class for StringUtilTest and is intended
'''to contain all StringUtilTest Unit Tests
'''</summary>
<TestFixture()> _
Public MustInherit Class StringUtilTest

    Private Class StrPvo : Inherits PrimitiveValueObject(Of String)
        Public Sub New(ByVal value As String)
            MyBase.New(value)
        End Sub
    End Class

    Public Class 通常 : Inherits StringUtilTest

        <Test()> Public Sub DecamelizeIgnoreNumber_数値の直前にアンダーバーで区切らない()
            Assert.AreEqual("RHAC0120", StringUtil.DecamelizeIgnoreNumber("Rhac0120"), "数字の前にアンダーバー付かない")
            Assert.AreEqual("123_RHAC", StringUtil.DecamelizeIgnoreNumber("123Rhac"), "数字の後はアンダーバーが付く")
        End Sub

        <Test()> Public Sub ToString_NullableIntegerで確認()
            Dim nullInteger As Nullable(Of Integer)
            Assert.IsNull(StringUtil.ToString(nullInteger))
            nullInteger = 123
            Assert.AreEqual("123", StringUtil.ToString(nullInteger))
        End Sub

        <Test()> Public Sub GetLengthByte_文字バイト数を取得()
            Assert.AreEqual(3, StringUtil.GetLengthByte("aaa"))
            Assert.AreEqual(6, StringUtil.GetLengthByte("ｑｑｑ"))
            Assert.AreEqual(7, StringUtil.GetLengthByte("ｍ_１２"))
        End Sub

        <Test()> Public Sub IncrementNumber()
            Assert.AreEqual("01", StringUtil.IncrementNumber("00"))
            Assert.AreEqual("10", StringUtil.IncrementNumber("09"))
            Assert.AreEqual("00", StringUtil.IncrementNumber("99"))
        End Sub

        <Test()> Public Sub IncrementAlphaNumber()
            Assert.AreEqual("01", StringUtil.IncrementAlphaNumber("00"))
            Assert.AreEqual("0A", StringUtil.IncrementAlphaNumber("09"))
            Assert.AreEqual("10", StringUtil.IncrementAlphaNumber("0Z"))
            Assert.AreEqual("00", StringUtil.IncrementAlphaNumber("ZZ"))
        End Sub

        <Test()> Public Sub TrimEndDecimalZero_小数点以下の末尾ゼロをTrimする()
            Assert.AreEqual("3.3", StringUtil.TrimEndDecimalZero("3.300"))
            Assert.AreEqual("4", StringUtil.TrimEndDecimalZero("4.00"))
            Assert.AreEqual("500", StringUtil.TrimEndDecimalZero("500.00"))
            Assert.AreEqual("123.4567", StringUtil.TrimEndDecimalZero("123.4567"), "Trimするゼロが無ければそのまま")
            Assert.AreEqual("6000", StringUtil.TrimEndDecimalZero("6000"), "整数部のゼロは取り除かない")
        End Sub

        <Test()> Public Sub IsZenkakuOnly_全角文字のみならtrue()
            Assert.AreEqual(True, StringUtil.IsZenkakuOnly("１"))
            Assert.AreEqual(True, StringUtil.IsZenkakuOnly("１２３"))
            Assert.AreEqual(True, StringUtil.IsZenkakuOnly("ａ"))
            Assert.AreEqual(True, StringUtil.IsZenkakuOnly("ａｂｃ"))

            Assert.AreEqual(False, StringUtil.IsZenkakuOnly("1"))
            Assert.AreEqual(False, StringUtil.IsZenkakuOnly("123"))
            Assert.AreEqual(False, StringUtil.IsZenkakuOnly("１2３"))
            Assert.AreEqual(False, StringUtil.IsZenkakuOnly("a"))
            Assert.AreEqual(False, StringUtil.IsZenkakuOnly("abc"))
            Assert.AreEqual(False, StringUtil.IsZenkakuOnly("aｂc"))
        End Sub

        <Test()> Public Sub IsHankakuOnly_半角文字のみならtrue()
            Assert.AreEqual(True, StringUtil.IsHankakuOnly("1"))
            Assert.AreEqual(True, StringUtil.IsHankakuOnly("123"))
            Assert.AreEqual(True, StringUtil.IsHankakuOnly("a"))
            Assert.AreEqual(True, StringUtil.IsHankakuOnly("abc"))

            Assert.AreEqual(False, StringUtil.IsHankakuOnly("１"))
            Assert.AreEqual(False, StringUtil.IsHankakuOnly("１２３"))
            Assert.AreEqual(False, StringUtil.IsHankakuOnly("１2３"))
            Assert.AreEqual(False, StringUtil.IsHankakuOnly("ａ"))
            Assert.AreEqual(False, StringUtil.IsHankakuOnly("ａｂｃ"))
            Assert.AreEqual(False, StringUtil.IsHankakuOnly("aｂc"))
        End Sub
    End Class

    Public Class DecamelizeTest : Inherits StringUtilTest

        <Test()> Public Sub Decamelize_通常ケース()
            Assert.AreEqual("USER_ID", StringUtil.Decamelize("userId"))
            Assert.AreEqual("M_USER", StringUtil.Decamelize("MUser"))
            Assert.AreEqual("M_USER", StringUtil.Decamelize("MUser"))
        End Sub

        <Test()> Public Sub Decamelize_数値の直前にアンダーバー()
            Assert.AreEqual("KINO_ID_1", StringUtil.Decamelize("KinoId1"))
            Assert.AreEqual("KINO_ID_23", StringUtil.Decamelize("KinoId23"))
            Assert.AreEqual("BUHIN_1_NO_2", StringUtil.Decamelize("Buhin1No2"))
            Assert.AreEqual("BUHIN_11_NO_22", StringUtil.Decamelize("Buhin11No22"))
        End Sub
    End Class

    Public Class NvlTest : Inherits StringUtilTest

        <Test()> Public Sub 判定する値が設定されていたら_そのまま値を返す(<Values("aa", "123")> value As String)
            Dim actual As String = StringUtil.Nvl(value, "NULL_VAL")
            Assert.That(actual, [Is].EqualTo(value))
        End Sub

        <Test()> Public Sub NULLの場合の文字列を設定しないメソッド_判定する値が設定されていたら_そのまま値を返す(<Values("aa", "123")> value As String)
            Dim actual As String = StringUtil.Nvl(value)
            Assert.That(actual, [Is].EqualTo(value))
        End Sub

        <Test()> Public Sub 判定する値がNULLだったら_NULLの場合の文字列を返す(<Values("abc", "111")> nullVal As String)
            Dim actual As String = StringUtil.Nvl(Nothing, nullVal)
            Assert.That(actual, [Is].EqualTo(nullVal))
        End Sub

        <Test()> Public Sub NULLの場合の文字列を設定しないメソッド_判定する値がNULLだったら_空の文字列を返す()
            Dim actual As String = StringUtil.Nvl(Nothing)
            Assert.That(actual, [Is].EqualTo(String.Empty))
        End Sub

        <TestCase()> Public Sub PrimitiveValueObject_値があれば_値の文字列を返す()
            Dim pvo As New StrPvo("abc")
            Dim actual As String = StringUtil.Nvl(pvo, "NullVal")
            Assert.That(actual, [Is].EqualTo("abc"))
        End Sub

        <TestCase()> Public Sub PrimitiveValueObject_値がNULLなら_NULLの場合の文字列を返す()
            Dim pvo As New StrPvo(Nothing)
            Dim actual As String = StringUtil.Nvl(pvo, "NullVal")
            Assert.That(actual, [Is].EqualTo("NullVal"))
        End Sub

        <Test()> Public Sub Nullable引数_で引数nullなら空文字が返る()
            Dim iNull As Integer? = Nothing
            Assert.AreEqual("", StringUtil.Nvl(iNull))

            Dim lNull As Long? = Nothing
            Assert.AreEqual("", StringUtil.Nvl(lNull))

            Dim dNull As Decimal? = Nothing
            Assert.AreEqual("", StringUtil.Nvl(dNull))

            Dim dtNull As DateTime? = Nothing
            Assert.AreEqual("", StringUtil.Nvl(dtNull))
        End Sub

        <Test()> Public Sub Nullable引数_で引数値有なら引数値が文字列で返る()
            Dim i As Integer? = 1
            Assert.AreEqual("1", StringUtil.Nvl(i))

            Dim l As Long? = 2
            Assert.AreEqual("2", StringUtil.Nvl(l))

            Dim dec As Decimal? = 3.14D
            Assert.AreEqual("3.14", StringUtil.Nvl(dec))

            Dim dt As DateTime? = CDate("2011/12/12 13:14:15")
            Assert.AreEqual("2011/12/12 13:14:15", StringUtil.Nvl(dt))
        End Sub
    End Class

    Public Class EvlTest : Inherits StringUtilTest

        <Test()> Public Sub 値があれば_第1引数の文字列を返す()
            Dim actual As String = StringUtil.Evl("abc", "EmptyVal")
            Assert.That(actual, [Is].EqualTo("abc"))
        End Sub

        <TestCase(Nothing)>
        <TestCase("")>
        <TestCase("   ")>
        Public Sub 値がNULL_空文字_空白のみのいずれかなら_第2引数の文字列を返す(value As String)
            Dim actual As String = StringUtil.Evl(value, "EmptyVal")
            Assert.That(actual, [Is].EqualTo("EmptyVal"))
        End Sub

        <Test()> Public Sub PrimitiveValueObject_値があれば_第1引数のPVOを文字列型にした値を返す()
            Dim pvo As New StrPvo("abc")
            Dim actual As String = StringUtil.Evl(pvo, "EmptyVal")
            Assert.That(actual, [Is].EqualTo("abc"))
        End Sub

        <TestCase(Nothing)>
        <TestCase("")>
        <TestCase("   ")>
        Public Sub PrimitiveValueObject_値がNULL_空文字_空白のみのいずれかなら_第2引数の文字列を返す(value As String)
            Dim pvo As New StrPvo(value)
            Dim actual As String = StringUtil.Evl(pvo, "EmptyVal")
            Assert.That(actual, [Is].EqualTo("EmptyVal"))
        End Sub

        <Test()> Public Sub PVO型を返すEvl_値があれば_第1引数の値オブジェクトを返す()
            Dim pvo As New StrPvo("abc")
            Dim actual As StrPvo = StringUtil.Evl(pvo, New StrPvo("EmptyVal"))
            Assert.That(actual, [Is].EqualTo(New StrPvo("abc")))
        End Sub

        <TestCase(Nothing)>
        <TestCase("")>
        <TestCase("   ")>
        Public Sub PVO型を返すEvl_値がNULL_空文字_空白のみのいずれかなら_第2引数の値オブジェクトを返す(value As String)
            Dim pvo As New StrPvo(value)
            Dim actual As StrPvo = StringUtil.Evl(pvo, New StrPvo("EmptyVal"))
            Assert.That(actual, [Is].EqualTo(New StrPvo("EmptyVal")))
        End Sub
    End Class

    Public Class EvlZeroTest : Inherits StringUtilTest

        <Test()> Public Sub 値があれば_第1引数の文字列を返す()
            Dim actual As String = StringUtil.EvlZero("123")
            Assert.That(actual, [Is].EqualTo("123"))
        End Sub

        <TestCase(Nothing)>
        <TestCase("")>
        <TestCase("   ")>
        Public Sub 値がNULL_空文字_空白のみのいずれかなら_文字列で0を返す(value As String)
            Dim actual As String = StringUtil.EvlZero(value)
            Assert.That(actual, [Is].EqualTo("0"))
        End Sub

        <Test()> Public Sub PrimitiveValueObject_値があれば_第1引数のPVOを文字列型にした値を返す()
            Dim pvo As New StrPvo("abc")
            Dim actual As String = StringUtil.Evl(pvo, "EmptyVal")
            Assert.That(actual, [Is].EqualTo("abc"))
        End Sub

        <TestCase(Nothing)>
        <TestCase("")>
        <TestCase("   ")>
        Public Sub PrimitiveValueObject_値がNULL_空文字_空白のみのいずれかなら_第2引数の文字列を返す(value As String)
            Dim pvo As New StrPvo(value)
            Dim actual As String = StringUtil.Evl(pvo, "EmptyVal")
            Assert.That(actual, [Is].EqualTo("EmptyVal"))
        End Sub
    End Class

    Public Class ToListTest : Inherits StringUtilTest

        <Test()> Public Sub 文字列の引数1つだけだったら()
            Dim result As List(Of String) = StringUtil.ToList("aa")
            Assert.AreEqual(1, result.Count)
            Assert.AreEqual("aa", result(0))
        End Sub

        <Test()> Public Sub 文字列の引数3つだったら()
            Dim result As List(Of String) = StringUtil.ToList("aa", "bb", "cc")
            Assert.AreEqual(3, result.Count)
            Assert.AreEqual("cc", result(2))
        End Sub

        <Test()> Public Sub 文字列配列の引数だったら()
            Dim params As String() = {"qq", "ww", "ee", "rr"}
            Dim result As List(Of String) = StringUtil.ToList(params)
            Assert.AreEqual(4, result.Count)
            Assert.AreEqual("rr", result(3))
        End Sub
    End Class

    Public Class FormatTest : Inherits StringUtilTest

        <Test()> Public Sub 数を整形()
            Assert.AreEqual("001", StringUtil.Format("000", 1))
            Assert.AreEqual("013", StringUtil.Format("000", 13))
            Assert.AreEqual("1234", StringUtil.Format("000", 1234))
            Assert.AreEqual("ABC002", StringUtil.Format("ABC000", 2), "文字列を組み合わせてもＯＫ")
        End Sub

        <Test()> Public Sub 日時を整形()
            Assert.AreEqual("20101112", StringUtil.Format("yyyyMMdd", CDate("2010/11/12 13:14:15")))
            Assert.AreEqual("131415", StringUtil.Format("HHmmss", CDate("2010/11/12 13:14:15")))
        End Sub

        <Test()> Public Sub 引数nullなら空文字を返す(<Values("#,###", "000")> ByVal formatString As String)
            Assert.That(StringUtil.Format(formatString, Nothing), [Is].Empty)
        End Sub
    End Class

    Public Class RemoveIfQuotedTest : Inherits StringUtilTest

        <Test()> Public Sub 囲まれていて外れるパターン()
            Assert.AreEqual("bbb", StringUtil.RemoveIfQuoted("""bbb"""))
            Assert.AreEqual("ccc", StringUtil.RemoveIfQuoted("'ccc'"))
            Assert.AreEqual("dd""ee", StringUtil.RemoveIfQuoted("""dd""ee"""))
            Assert.AreEqual("ff''gg", StringUtil.RemoveIfQuoted("""ff''gg"""))
        End Sub

        <Test()> Public Sub 囲まれていないから外れないパターン()
            Assert.AreEqual("aaa", StringUtil.RemoveIfQuoted("aaa"))
            Assert.AreEqual("'ddd""", StringUtil.RemoveIfQuoted("'ddd"""))
            Assert.AreEqual("""eee'", StringUtil.RemoveIfQuoted("""eee'"))
        End Sub
    End Class

    Public Class MakeFixedStringTest : Inherits StringUtilTest

        <Test(), Sequential()> Public Sub 固定長文字列の作成(<Values("aaaa", "ああ")> ByVal value As String, <Values("aaaa  ", "ああ  ")> ByVal expected As String)
            Assert.AreEqual(expected, StringUtil.MakeFixedString(value, 6))
        End Sub

        <Test(), Sequential()> Public Sub 固定長文字列の作成_右寄せ(<Values("1234", "５６")> ByVal value As String, <Values("  1234", "  ５６")> ByVal expected As String)
            Assert.AreEqual(expected, StringUtil.MakeFixedString(value, 6, alignsRight:=True))
        End Sub

        <Test()> Public Sub 固定長桁数が文字列長未満なら文字列をそのまま返す(<Values("bbb", "いい")> ByVal value As String)
            Assert.AreEqual(value, StringUtil.MakeFixedString(value, 1))
        End Sub

        <Test()> Public Sub Nothingでも動作する()
            Assert.That(StringUtil.MakeFixedString(Nothing, 2), [Is].EqualTo("  "))
        End Sub
    End Class

    Public Class IsMatchFilterPatternTest : Inherits StringUtilTest

        <Test()> Public Sub 拡張子無しファイルとパターン文字()
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("filename", "*"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("filename", "file"))
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("filename", "file*"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("filename", "*.jpg"))
        End Sub

        <Test()> Public Sub 拡張子マスクのパターン文字()
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("file.jpg", "*.jpg"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("file.jpg", "*.bmp"))
        End Sub

        <Test()> Public Sub 拡張子マスクのパターン文字_マルチピリオド()
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("multi.period.jpg", "*.jpg"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("multi.period.jpg", "*.bmp"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("multi.jpg.xls", "*.jpg"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("bmp.period.xls", "*.bmp"))
        End Sub

        <Test()> Public Sub 精査文字列にエスケープ文字が含まれても動作する()
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("file(0).jpg", "*.jpg"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("file(0).jpg", "*.bmp"))
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("file[0].jpg", "*.jpg"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("file[0].jpg", "*.bmp"))
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("contains space.jpg", "*.jpg"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("contains space.jpg", "*.bmp"))
        End Sub

        <Test()> Public Sub 精査文字列が大文字で拡張子マスクは小文字でも動作する()
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("FILE.JPG", "*.jpg"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("FILE.JPG", "*.bmp"))
        End Sub

        <Test()> Public Sub 複数の拡張子マスクで動作する()
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("file.jpg", "*.jpg", "*.xls", "*.txt"))
            Assert.IsFalse(StringUtil.IsMatchFilterPattern("file.jpg", "*.bmp", "*.xls", "*.txt"))
        End Sub

        <Test()> Public Sub パターン文字アスタアスタは例外的に全てを許容する()
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("filename", "*.*"))
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("filename.jpg", "*.*"))
            Assert.IsTrue(StringUtil.IsMatchFilterPattern("multi.period.bmp", "*.*"))
        End Sub
    End Class

    Public Class EqualsIgnoreCaseIfNullTest : Inherits StringUtilTest

        <Test> Public Sub 大文字小文字を区別せず一致ならTrue()
            Assert.IsTrue(StringUtil.EqualsIgnoreCaseIfNull("ABC", "aBc"))
        End Sub

        <Test> Public Sub 大文字小文字関係なく不一致ならFalse()
            Assert.IsFalse(StringUtil.EqualsIgnoreCaseIfNull("ABC", "XYZ"))
        End Sub

        <Test> Public Sub Null同士ならTrue()
            Assert.IsTrue(StringUtil.EqualsIgnoreCaseIfNull(Nothing, Nothing))
        End Sub

        <Test> Public Sub Nullと文字列同士ならFalse(<Values("", "AAA")> str As String)
            Assert.IsFalse(StringUtil.EqualsIgnoreCaseIfNull(Nothing, str))
        End Sub
    End Class

    Public Class ContainsIgnoreCaseTest : Inherits StringUtilTest

        <Test()> Public Sub 大文字小文字を区別せず_含み検出()
            Dim testee As String() = New String() {"AA", "BB", "CC"}
            Assert.IsTrue(StringUtil.ContainsIgnoreCase(testee, "bb"))
        End Sub

        <Test()> Public Sub 大文字小文字を区別せず_だがNULL指定で検出せず()
            Dim testee As String() = New String() {"AA", "BB", "CC"}
            Assert.IsFalse(StringUtil.ContainsIgnoreCase(testee, Nothing))
        End Sub

        <Test()> Public Sub 大文字小文字を区別せず_NULL入りだが含み検出()
            Dim testee As String() = New String() {Nothing, "BB", "CC"}
            Assert.IsTrue(StringUtil.ContainsIgnoreCase(testee, "bb"))
        End Sub

        <Test()> Public Sub 大文字小文字を区別せず_空文字入りだが含み検出()
            Dim testee As String() = New String() {String.Empty, "BB", "CC"}
            Assert.IsTrue(StringUtil.ContainsIgnoreCase(testee, "bb"))
        End Sub
    End Class

    Public Class ToIntegerTest : Inherits StringUtilTest

        <TestCase("123", 123), TestCase("-123", -123)> _
        Public Sub 通常の数値文字列(value As Object, expected As Integer)
            Assert.AreEqual(expected, StringUtil.ToInteger(value))
            Assert.AreEqual(expected, StringUtil.ToInteger(value.ToString))
        End Sub

        <Test()> Public Sub 文字列はnull(<Values("", Nothing, "a123")> value As String)
            Assert.IsNull(StringUtil.ToInteger(value))
        End Sub

        <Test()> Public Sub 文字列はnull_だから第二引数の値を返す(<Values("", Nothing, "a123")> value As String)
            Assert.That(StringUtil.ToInteger(value, 12), [Is].EqualTo(12))
        End Sub

    End Class

    Public Class ToDatetimeTest : Inherits StringUtilTest

        <TestCase("2018/1/2", "2018/1/2"), TestCase("1970/1/3", "1970/1/3")> _
        Public Sub 通常の数値文字列(value As Object, expected As String)
            Assert.AreEqual(CDate(expected), StringUtil.ToDatetime(value))
            Assert.AreEqual(CDate(expected), StringUtil.ToDatetime(value.ToString))
        End Sub

        <Test()> Public Sub 文字列はnull(<Values("", Nothing, "a123")> value As String)
            Assert.IsNull(StringUtil.ToDatetime(value))
        End Sub

        <Test()> Public Sub 文字列はnull_だから第二引数の値を返す(<Values("", Nothing, "a123")> value As String)
            Assert.That(StringUtil.ToDatetime(value, #3/4/2018#), [Is].EqualTo(#3/4/2018#))
        End Sub

    End Class

    Public Class ToDecimalTest : Inherits StringUtilTest

        <Test()> Public Sub 通常の数値文字列(<Values("123.456", "-123.567")> value As Object)
            Assert.AreEqual(CDec(value), StringUtil.ToDecimal(value))
            Assert.AreEqual(CDec(value), StringUtil.ToDecimal(value.ToString))
        End Sub

        <Test()> Public Sub 文字列はnull(<Values("", Nothing, "a123")> value As String)
            Assert.IsNull(StringUtil.ToDecimal(value))
        End Sub

        <Test()> Public Sub 文字列はnull_だから第二引数の値を返す(<Values("", Nothing, "a123")> value As String)
            Assert.That(StringUtil.ToDecimal(value, 12.3D), [Is].EqualTo(12.3D))
        End Sub

    End Class

    Public Class ToDoubleTest : Inherits StringUtilTest

        <Test()> Public Sub 通常の数値文字列(<Values("123.456", "-123.567")> value As Object)
            Assert.AreEqual(Convert.ToDouble(value), StringUtil.ToDouble(value))
            Assert.AreEqual(Convert.ToDouble(value), StringUtil.ToDouble(value.ToString))
        End Sub

        <Test()> Public Sub 文字列はnull(<Values("", Nothing, "a123")> value As String)
            Assert.IsNull(StringUtil.ToDouble(value))
        End Sub

        <Test()> Public Sub 文字列はnull_だから第二引数の値を返す(<Values("", Nothing, "a123")> value As String)
            Assert.That(StringUtil.ToDouble(value, 12.3D), [Is].EqualTo(12.3D))
        End Sub

    End Class

    Public Class ToLongTest : Inherits StringUtilTest

        <Test()> Public Sub 通常の数値文字列(<Values("1234567890123", "-567")> value As Object)
            Assert.AreEqual(Convert.ToInt64(value), StringUtil.ToDouble(value))
            Assert.AreEqual(Convert.ToInt64(value), StringUtil.ToDouble(value.ToString))
        End Sub

        <Test()> Public Sub 文字列はnull(<Values("", Nothing, "a123")> value As String)
            Assert.IsNull(StringUtil.ToDouble(value))
        End Sub

        <Test()> Public Sub 文字列はnull_だから第二引数の値を返す(<Values("", Nothing, "a123")> value As String)
            Assert.That(StringUtil.ToDouble(value, 12.3D), [Is].EqualTo(12.3D))
        End Sub

    End Class

    Public Class ToShortTest : Inherits StringUtilTest

        <Test()> Public Sub 通常の数値文字列(<Values("123", "255")> value As Object)
            Assert.AreEqual(Convert.ToByte(value), StringUtil.ToDouble(value))
            Assert.AreEqual(Convert.ToByte(value), StringUtil.ToDouble(value.ToString))
        End Sub

        <Test()> Public Sub 文字列はnull(<Values("", Nothing, "a123")> value As String)
            Assert.IsNull(StringUtil.ToDouble(value))
        End Sub

        <Test()> Public Sub 文字列はnull_だから第二引数の値を返す(<Values("", Nothing, "a123")> value As String)
            Assert.That(StringUtil.ToDouble(value, 12.3D), [Is].EqualTo(12.3D))
        End Sub

    End Class

    Public Class TrimEnd_末尾のスペースや不正な文字を除去 : Inherits StringUtilTest

        <Test()> Public Sub 末尾スペースは除去する()
            Assert.That(StringUtil.TrimEnd(" ab c "), [Is].EqualTo(" ab c"))
        End Sub

        <Test()> Public Sub 末尾の不正な文字コードも除去する(<Values(&H10, &H7F)> ByVal invalidChar As Integer)
            Dim bytes As Byte() = {&H41, Convert.ToByte(invalidChar)}
            Dim s As String = System.Text.Encoding.GetEncoding(932).GetString(bytes)
            Assert.That(s, [Is].Not.EqualTo("A"), "２文字だから一致する訳ない")
            Assert.That(s.TrimEnd, [Is].Not.EqualTo("A"), "スペースじゃないからTrimされない")
            Assert.That(StringUtil.TrimEnd(s), [Is].EqualTo("A"), "スペースじゃないけど不正な文字コードだからTrimする")
        End Sub

        <Test()> Public Sub 不正な文字コードだけでも除去する(<Values(&H10, &H7F, &H1A)> ByVal invalidChar As Integer)
            Dim bytes As Byte() = {Convert.ToByte(invalidChar), Convert.ToByte(invalidChar)}
            Dim invalidCharOnlyString As String = System.Text.Encoding.GetEncoding(932).GetString(bytes)
            Assert.That(invalidCharOnlyString, [Is].Not.EqualTo(""), "２文字だから一致する訳ない")
            Assert.That(invalidCharOnlyString.TrimEnd, [Is].Not.EqualTo(""), "標準TrimEndはスペースじゃないからTrimされない")
            Assert.That(StringUtil.TrimEnd(invalidCharOnlyString), [Is].EqualTo(""), "スペースじゃないけど不正な文字コードだからTrimする")
        End Sub

        <Test(), Sequential()> Public Sub UTF8含む全角文字末尾の不正な文字コードも除去する( _
                <Values("Ａ"c, "㌢"c, "℡"c, "￤"c, "ⅷ"c, "✔"c, "€"c)> ByVal zenkaku As Char)
            Const INVALID_CHAR_CODE As Integer = &H1A
            Dim s As String = zenkaku.ToString & System.Text.Encoding.GetEncoding(932).GetString({Convert.ToByte(INVALID_CHAR_CODE)})
            Assert.That(StringUtil.TrimEnd(s), [Is].EqualTo(zenkaku.ToString))
        End Sub

        <Test()> Public Sub スペースのみなら空文字になる()
            Assert.That(StringUtil.TrimEnd("  "), [Is].EqualTo(""))
        End Sub

        <Test()> Public Sub 空文字なら空文字()
            Assert.That(StringUtil.TrimEnd(""), [Is].EqualTo(""))
        End Sub

        <Test()> Public Sub NothingならNothing()
            Assert.That(StringUtil.TrimEnd(Nothing), [Is].Null)
        End Sub

    End Class

    Public Class TrimStart_開始位置のスペースや不正な文字を除去 : Inherits StringUtilTest

        <Test()> Public Sub 接頭スペースは除去する()
            Assert.That(StringUtil.TrimStart(" ab c "), [Is].EqualTo("ab c "))
        End Sub

        <Test()> Public Sub 接頭の不正な文字コードも除去する(<Values(&H10, &H7F)> ByVal invalidChar As Integer)
            Dim bytes As Byte() = {Convert.ToByte(invalidChar), &H41}
            Dim s As String = System.Text.Encoding.GetEncoding(932).GetString(bytes)
            Assert.That(s, [Is].Not.EqualTo("A"), "２文字だから一致する訳ない")
            Assert.That(s.TrimStart, [Is].Not.EqualTo("A"), "スペースじゃないからTrimされない")
            Assert.That(StringUtil.TrimStart(s), [Is].EqualTo("A"), "スペースじゃないけど不正な文字コードだからTrimする")
        End Sub

        <Test()> Public Sub 不正な文字コードだけでも除去する(<Values(&H10, &H7F, &H1A)> ByVal invalidChar As Integer)
            Dim bytes As Byte() = {Convert.ToByte(invalidChar), Convert.ToByte(invalidChar)}
            Dim invalidCharOnlyString As String = System.Text.Encoding.GetEncoding(932).GetString(bytes)
            Assert.That(invalidCharOnlyString, [Is].Not.EqualTo(""), "２文字だから一致する訳ない")
            Assert.That(invalidCharOnlyString.TrimStart, [Is].Not.EqualTo(""), "標準TrimStartはスペースじゃないからTrimされない")
            Assert.That(StringUtil.TrimStart(invalidCharOnlyString), [Is].EqualTo(""), "スペースじゃないけど不正な文字コードだからTrimする")
        End Sub

        <Test(), Sequential()> Public Sub UTF8含む全角文字接頭の不正な文字コードも除去する( _
                <Values("Ａ"c, "㌢"c, "℡"c, "￤"c, "ⅷ"c, "✔"c, "€"c)> ByVal zenkaku As Char)
            Const INVALID_CHAR_CODE As Integer = &H1A
            Dim s As String = System.Text.Encoding.GetEncoding(932).GetString({Convert.ToByte(INVALID_CHAR_CODE)}) & zenkaku.ToString
            Assert.That(StringUtil.TrimStart(s), [Is].EqualTo(zenkaku.ToString))
        End Sub

        <Test()> Public Sub スペースのみなら空文字になる()
            Assert.That(StringUtil.TrimStart("  "), [Is].EqualTo(""))
        End Sub

        <Test()> Public Sub 空文字なら空文字()
            Assert.That(StringUtil.TrimStart(""), [Is].EqualTo(""))
        End Sub

        <Test()> Public Sub NothingならNothing()
            Assert.That(StringUtil.TrimStart(Nothing), [Is].Null)
        End Sub

    End Class

    Public Class Trim_前後のスペースや不正な文字を除去 : Inherits StringUtilTest

        <Test()> Public Sub 末尾スペースは除去する()
            Assert.That(StringUtil.Trim(" ab c "), [Is].EqualTo("ab c"))
        End Sub

        <Test()> Public Sub 末尾の不正な文字コードも除去する(<Values(&H10, &H7F)> ByVal invalidChar As Integer)
            Dim bytes As Byte() = {Convert.ToByte(invalidChar), &H41, Convert.ToByte(invalidChar)}
            Dim s As String = System.Text.Encoding.GetEncoding(932).GetString(bytes)
            Assert.That(s, [Is].Not.EqualTo("A"), "3文字だから一致する訳ない")
            Assert.That(s.Trim, [Is].Not.EqualTo("A"), "スペースじゃないからTrimされない")
            Assert.That(StringUtil.Trim(s), [Is].EqualTo("A"), "スペースじゃないけど不正な文字コードだからTrimする")
        End Sub

        <Test()> Public Sub 不正な文字コードだけでも除去する(<Values(&H10, &H7F, &H1A)> ByVal invalidChar As Integer)
            Dim bytes As Byte() = {Convert.ToByte(invalidChar), Convert.ToByte(invalidChar)}
            Dim invalidCharOnlyString As String = System.Text.Encoding.GetEncoding(932).GetString(bytes)
            Assert.That(invalidCharOnlyString, [Is].Not.EqualTo(""), "２文字だから一致する訳ない")
            Assert.That(invalidCharOnlyString.Trim, [Is].Not.EqualTo(""), "標準TrimはスペースじゃないからTrimされない")
            Assert.That(StringUtil.Trim(invalidCharOnlyString), [Is].EqualTo(""), "スペースじゃないけど不正な文字コードだからTrimする")
        End Sub

        <Test(), Sequential()> Public Sub UTF8含む全角文字末尾の不正な文字コードも除去する( _
                <Values("Ａ"c, "㌢"c, "℡"c, "￤"c, "ⅷ"c, "✔"c, "€"c)> ByVal zenkaku As Char)
            Const INVALID_CHAR_CODE As Integer = &H1A
            Dim invalidString As String = System.Text.Encoding.GetEncoding(932).GetString({Convert.ToByte(INVALID_CHAR_CODE)})
            Dim s As String = invalidString & zenkaku.ToString & invalidString
            Assert.That(StringUtil.Trim(s), [Is].EqualTo(zenkaku.ToString))
        End Sub

        <Test()> Public Sub スペースのみなら空文字になる()
            Assert.That(StringUtil.Trim("  "), [Is].EqualTo(""))
        End Sub

        <Test()> Public Sub 空文字なら空文字()
            Assert.That(StringUtil.Trim(""), [Is].EqualTo(""))
        End Sub

        <Test()> Public Sub NothingならNothing()
            Assert.That(StringUtil.Trim(Nothing), [Is].Null)
        End Sub

    End Class

    Public Class CompareNullsLast_Null値を最後に並べるComparer値を返す : Inherits StringUtilTest

        <Test()> Public Sub Null同士なら一致_0()
            Assert.AreEqual(0, StringUtil.CompareNullsLast(Nothing, Nothing))
        End Sub

        <Test()> Public Sub どちらかがNullなら_Null側が最後にソートされる値を返す()
            Assert.AreEqual(1, StringUtil.CompareNullsLast(Nothing, "abc"), "> 0 だから1 - Null値のほうが大きい判定→最後にソート")
            Assert.AreEqual(-1, StringUtil.CompareNullsLast("abc", Nothing), "< 0 だから-1 - Null値のほうが大きい判定→最後にソート")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート()
            Assert.AreEqual(-1, StringUtil.CompareNullsLast("abc", "xyz"), "< 0 だから-1")
            Assert.AreEqual(1, StringUtil.CompareNullsLast("xyz", "abc"), "> 0 だから1")
            Assert.AreEqual(0, StringUtil.CompareNullsLast("abc", "abc"), "= 0")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート_Object型()
            Assert.AreEqual(-1, StringUtil.CompareNullsLast(DirectCast("abc", Object), DirectCast("xyz", Object)), "< 0 だから-1")
            Assert.AreEqual(1, StringUtil.CompareNullsLast(DirectCast("xyz", Object), DirectCast("abc", Object)), "> 0 だから1")
            Assert.AreEqual(0, StringUtil.CompareNullsLast(DirectCast("abc", Object), DirectCast("abc", Object)), "= 0")
        End Sub

        <Test()> Public Sub 実習_Null値は最後にソートする()
            Dim aList As New List(Of String)
            aList.AddRange(New String() {"C", "A", Nothing, "B"})
            aList.Sort(Function(x As String, y As String) StringUtil.CompareNullsLast(x, y))

            Assert.AreEqual(4, aList.Count)
            Assert.AreEqual("A", aList(0))
            Assert.AreEqual("B", aList(1))
            Assert.AreEqual("C", aList(2))
            Assert.IsNull(aList(3))
        End Sub
    End Class

    Public Class CompareDescNullsLast_Null値を最後に並べる降順Comparer値を返す : Inherits StringUtilTest

        <Test()> Public Sub Null同士なら一致_0()
            Assert.AreEqual(0, StringUtil.CompareDescNullsLast(Nothing, Nothing))
        End Sub

        <Test()> Public Sub どちらかがNullなら_Null側が最後にソートされる値を返す()
            Assert.AreEqual(1, StringUtil.CompareDescNullsLast(Nothing, "abc"), "> 0 だから1 - Null値のほうが大きい判定→最後にソート")
            Assert.AreEqual(-1, StringUtil.CompareDescNullsLast("abc", Nothing), "< 0 だから-1 - Null値のほうが大きい判定→最後にソート")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート()
            Assert.AreEqual(1, StringUtil.CompareDescNullsLast("abc", "z"), "< 0 だけど降順Sortしたいから反対の1")
            Assert.AreEqual(-1, StringUtil.CompareDescNullsLast("z", "abc"), "> 0 だけど降順Sortしたいから反対の-1")
            Assert.AreEqual(0, StringUtil.CompareDescNullsLast("xx", "xx"), "= 0")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート_Object型()
            Assert.AreEqual(1, StringUtil.CompareDescNullsLast(DirectCast("abc", Object), DirectCast("z", Object)), "< 0 だけど降順Sortしたいから反対の1")
            Assert.AreEqual(-1, StringUtil.CompareDescNullsLast(DirectCast("z", Object), DirectCast("abc", Object)), "> 0 だけど降順Sortしたいから反対の-1")
            Assert.AreEqual(0, StringUtil.CompareDescNullsLast(DirectCast("xx", Object), DirectCast("xx", Object)), "= 0")
        End Sub

        <Test()> Public Sub 実習_Null値は最後にソートする()
            Dim aList As New List(Of String)
            aList.AddRange(New String() {"C", "A", Nothing, "B"})
            aList.Sort(Function(x As String, y As String) StringUtil.CompareDescNullsLast(x, y))

            Assert.AreEqual(4, aList.Count)
            Assert.AreEqual("C", aList(0))
            Assert.AreEqual("B", aList(1))
            Assert.AreEqual("A", aList(2))
            Assert.IsNull(aList(3))
        End Sub
    End Class

    Public Class ConvToIndexTable_指定パラメータを変換テーブルのインデックス値で返す : Inherits StringUtilTest
        Const CONV_TABLE As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        <Test()> Public Sub 変換テスト()
            Assert.AreEqual(0, StringUtil.ConvToIndexFromTable(CONV_TABLE, "0"))
            Assert.AreEqual(1, StringUtil.ConvToIndexFromTable(CONV_TABLE, "1"))
            Assert.AreEqual(9, StringUtil.ConvToIndexFromTable(CONV_TABLE, "9"))
            Assert.AreEqual(10, StringUtil.ConvToIndexFromTable(CONV_TABLE, "A"))
            Assert.AreEqual(34, StringUtil.ConvToIndexFromTable(CONV_TABLE, "Y"))
            Assert.AreEqual(35, StringUtil.ConvToIndexFromTable(CONV_TABLE, "Z"))
        End Sub

    End Class

    Public Class MakeRandomString_ランダム文字列を作成する : Inherits StringUtilTest

        <Test()> Public Sub 候補が一文字だから_桁数に沿った文字列を返す()
            Assert.AreEqual("AAA", StringUtil.MakeRandomString("A", 3))
            Assert.AreEqual("BBBBBBBBBB", StringUtil.MakeRandomString("B", 10))
        End Sub

        <Test()> Public Sub 末尾文字も返される()
            For i As Integer = 1 To 100
                Dim actual As String = StringUtil.MakeRandomString("ABC", 3)
                If 0 <= actual.IndexOf("C") Then
                    Assert.Pass("指定文字の末尾も含まれている")
                End If
                Thread.Sleep(10)
            Next
            Assert.Fail("100回繰り返しても、指定文字の末尾が含まれなかった。要確認")
        End Sub

    End Class

    Public Class SplitForEnclosedDQ_正常データ : Inherits StringUtilTest

        Const COMMA As Char = ","c

        <Test()> Public Sub 先頭がDQで囲まれててもOK()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("""A B"" C", " "c)
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("A B", actuals(0))
            Assert.AreEqual("C", actuals(1))
        End Sub

        <Test()> Public Sub 末尾がDQで囲まれててもOK()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A ""B C""", " "c)
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("B C", actuals(1))
        End Sub

        <Test()> Public Sub 文中がDQで囲まれててもOK()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A ""B C"" D", " "c)
            Assert.AreEqual(3, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("B C", actuals(1))
            Assert.AreEqual("D", actuals(2))
        End Sub

        <Test()> Public Sub DQで囲まれてなければ普通に処理()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A B C", " "c)
            Assert.AreEqual(3, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("B", actuals(1))
            Assert.AreEqual("C", actuals(2))
        End Sub

        <Test()> Public Sub DQの空文字は_DQ無しの空文字になる()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A,""""", ","c)
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("", actuals(1))
        End Sub

        <Test()> Public Sub DQで囲まれたDQ始まりの文字列を取得できる()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A,""""""B""", ","c)
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("""B", actuals(1))
        End Sub

        <Test()> Public Sub DQで囲まれたDQ終わりの文字列を取得できる()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A,""B""""""", ","c)
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("B""", actuals(1))
        End Sub

        <Test()> Public Sub DQだけの文字列を取得できる()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A,""", COMMA)
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("""", actuals(1), "DQだけでも取得できる")
        End Sub

        Private Const DQ As String = """"
        <Test(), Sequential()> Public Sub DQで囲まれていない値はTrimする_行頭( _
                    <Values("x", " x ", DQ & " x " & DQ, "  " & DQ & " x " & DQ & " ", vbTab & DQ & vbTab & "x" & vbTab & DQ & vbTab, vbCrLf & DQ & vbCrLf & "x" & vbCrLf & DQ & vbCrLf)> ByVal value As String, _
                    <Values("x", "x", " x ", " x ", vbTab & "x" & vbTab, vbCrLf & "x" & vbCrLf)> ByVal actual As String)
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ(value & ",A", COMMA)
            Assert.That(actuals, [Is].Not.Empty)
            Assert.That(actuals(0), [Is].EqualTo(actual))
            Assert.That(actuals(1), [Is].EqualTo("A"))
            Assert.That(actuals.Length, [Is].EqualTo(2))
        End Sub

        <Test(), Sequential()> Public Sub DQで囲まれていない値はTrimする_行中( _
                    <Values("x", " x ", DQ & " x " & DQ, "  " & DQ & " x " & DQ & " ", vbTab & DQ & vbTab & "x" & vbTab & DQ & vbTab, vbCrLf & DQ & vbCrLf & "x" & vbCrLf & DQ & vbCrLf)> ByVal value As String, _
                    <Values("x", "x", " x ", " x ", vbTab & "x" & vbTab, vbCrLf & "x" & vbCrLf)> ByVal actual As String)
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A," & value & ",B", COMMA)
            Assert.That(actuals, [Is].Not.Empty)
            Assert.That(actuals(0), [Is].EqualTo("A"))
            Assert.That(actuals(1), [Is].EqualTo(actual))
            Assert.That(actuals(2), [Is].EqualTo("B"))
            Assert.That(actuals.Length, [Is].EqualTo(3))
        End Sub

        <Test(), Sequential()> Public Sub DQで囲まれていない値はTrimする_行末( _
                    <Values("x", " x ", DQ & " x " & DQ, "  " & DQ & " x " & DQ & " ", vbTab & DQ & vbTab & "x" & vbTab & DQ & vbTab, vbCrLf & DQ & vbCrLf & "x" & vbCrLf & DQ & vbCrLf)> ByVal value As String, _
                    <Values("x", "x", " x ", " x ", vbTab & "x" & vbTab, vbCrLf & "x" & vbCrLf)> ByVal actual As String)
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A," & value, COMMA)
            Assert.That(actuals, [Is].Not.Empty)
            Assert.That(actuals(0), [Is].EqualTo("A"))
            Assert.That(actuals(1), [Is].EqualTo(actual))
            Assert.That(actuals.Length, [Is].EqualTo(2))
        End Sub

    End Class

    Public Class ConvertValueToRuleStringTest : Inherits StringUtilTest

        <Test(), Sequential()> Public Sub _0始まりの文字列じゃないから桁上がりはBからなので注意( _
                <Values(0, 2, 3, 5, 6, 9)> ByVal value As Integer, _
                <Values("A", "C", "BA", "BC", "CA", "BAA")> ByVal expected As String)
            Const RULE_STRING As String = "ABC"
            Assert.That(StringUtil.ConvertValueToRuleString(RULE_STRING, value), [Is].EqualTo(expected))
        End Sub
    End Class

    Public Class SplitForEnclosedDQ_いじわる引数 : Inherits StringUtilTest

        <Test()> Public Sub 閉じてるようにみえるけど_文字列が続いているなら_囲まれた箇所を処理して返す()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A ""B C""E", " "c)
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("B CE", actuals(1))
        End Sub

        <Test()> Public Sub 閉じてるようにみえるけど_文字列が続いているなら_囲まれた箇所を処理して返す2()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A ""B """"C""E", " "c)
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("B ""CE", actuals(1), "DQで囲まれた箇所を処理するからDQは一つになる")
        End Sub

        <Test()> Public Sub 閉じてるようにみえるけど_文字列が続いているなら_囲まれた箇所を処理して返す3()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("0,""A"""",B""", ","c)
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("0", actuals(0))
            Assert.AreEqual("A"",B", actuals(1), "DQで囲まれた箇所を処理するからDQは一つになる")
        End Sub

        <Test()> Public Sub 開きわすれ()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A B C""", " "c)
            Assert.AreEqual(3, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("B", actuals(1))
            Assert.AreEqual("C""", actuals(2))
        End Sub

        <Test()> Public Sub 閉じわすれ()
            Dim actuals As String() = StringUtil.SplitForEnclosedDQ("A ""B C", " "c)
            Assert.AreEqual(2, actuals.Length)
            Assert.AreEqual("A", actuals(0))
            Assert.AreEqual("""B C", actuals(1))
        End Sub

    End Class

    Public Class Left_左から指定桁だけ返す : Inherits StringUtilTest

        <Test()> Public Sub Learning_VB_Left_桁が足りなくても()
            Assert.AreEqual("123", Left("123", 5))
            Assert.AreEqual("１２３", Left("１２３", 5), "全角文字でも文字数で返す")
        End Sub

        <Test()> Public Sub Learning_VB_Left_Null値はEmptyで返る()
            Assert.AreEqual("", Left(Nothing, 5), "NullはEmptyにされる")
        End Sub

        <Test()> Public Sub 正常系()
            Assert.AreEqual("123", StringUtil.Left("12345", 3))
            Assert.AreEqual("１２３", StringUtil.Left("１２３４５", 3), "全角文字でも文字数で返す")
        End Sub

        <Test()> Public Sub 桁が足りなければ_そのまま返す()
            Assert.AreEqual("123", StringUtil.Left("123", 5))
            Assert.AreEqual("１２３", StringUtil.Left("１２３", 5))
        End Sub

        <Test()> Public Sub Null値ならNull値で返す()
            Assert.IsNull(StringUtil.Left(Nothing, 5), "nullならnull")
        End Sub

    End Class

    Public Class IsAlphabet_アルファベット文字を判定 : Inherits StringUtilTest

        <Test()> Public Sub アルファベット文字ならtrue()
            Assert.IsTrue(StringUtil.IsAlphabetHalf("A"c))
            Assert.IsTrue(StringUtil.IsAlphabetHalf("Z"c))
            Assert.IsTrue(StringUtil.IsAlphabetHalf("a"c))
            Assert.IsTrue(StringUtil.IsAlphabetHalf("z"c))
            Assert.IsTrue(StringUtil.IsAlphabetHalf("d"c))
        End Sub

        <Test()> Public Sub アルファベット文字じゃないからfalse()
            Assert.IsFalse(StringUtil.IsAlphabetHalf("1"c))
            Assert.IsFalse(StringUtil.IsAlphabetHalf("#"c))
            Assert.IsFalse(StringUtil.IsAlphabetHalf("@"c))
            Assert.IsFalse(StringUtil.IsAlphabetHalf("~"c))
            Assert.IsFalse(StringUtil.IsAlphabetHalf("]"c))
            Assert.IsFalse(StringUtil.IsAlphabetHalf("ｱ"c), "半角カタカナ")
            Assert.IsFalse(StringUtil.IsAlphabetHalf("１"c), "全角数字")
            Assert.IsFalse(StringUtil.IsAlphabetHalf("ａ"c), "全角英小文字")
            Assert.IsFalse(StringUtil.IsAlphabetHalf("Ａ"c), "全角英大文字")
            Assert.IsFalse(StringUtil.IsAlphabetHalf("あ"c), "全角ひらがな")
            Assert.IsFalse(StringUtil.IsAlphabetHalf("ア"c), "全角カタカナ")
            Assert.IsFalse(StringUtil.IsAlphabetHalf("亜"c), "全角漢字")
        End Sub

    End Class

    Public Class IsAlphabetOrNumber_ : Inherits StringUtilTest

        <Test()> Public Sub 英数のみならtrue()
            Assert.IsTrue(StringUtil.IsAlphabetOrNumber("1234567890"), "数字のみならtrue")
            Assert.IsTrue(StringUtil.IsAlphabetOrNumber("abcdefghijklmnopqrstuvwxyz"), "英小文字のみならtrue")
            Assert.IsTrue(StringUtil.IsAlphabetOrNumber("ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "英大文字のみならtrue")
            Assert.IsTrue(StringUtil.IsAlphabetOrNumber("a1Az9Z"), "英数字混在しててもtrue")
        End Sub

        <Test()> Public Sub 英数以外ならfalse()
            Assert.IsFalse(StringUtil.IsAlphabetOrNumber("ｱ"), "半角カタカナ")
            Assert.IsFalse(StringUtil.IsAlphabetOrNumber("*"), "半角記号")
            Assert.IsFalse(StringUtil.IsAlphabetOrNumber("１"), "全角数字")
            Assert.IsFalse(StringUtil.IsAlphabetOrNumber("ａ"), "全角英小文字")
            Assert.IsFalse(StringUtil.IsAlphabetOrNumber("Ａ"), "全角英大文字")
            Assert.IsFalse(StringUtil.IsAlphabetOrNumber("あ"), "全角ひらがな")
            Assert.IsFalse(StringUtil.IsAlphabetOrNumber("ア"), "全角カタカナ")
            Assert.IsFalse(StringUtil.IsAlphabetOrNumber("亜"), "全角漢字")
        End Sub

        <Test()> Public Sub AllowChars_英数以外の許可する文字を指定できる()
            Assert.IsFalse(StringUtil.IsAlphabetOrNumber("1aZ-"), "allowCharsが無ければfalse")
            Assert.IsTrue(StringUtil.IsAlphabetOrNumber("1aZ-", "-"c), "allowCharsを指定してるからtrue")
        End Sub

    End Class

    Public Class ExtractEnclosedString_囲まれた文字列を返す : Inherits StringUtilTest

        <Test(), Sequential()> Public Sub AllowChars_英数以外の許可する文字を指定できる( _
                <Values("a'bc'd", "a[bcd]e", "x y z")> ByVal target As String, _
                <Values("'", "[", " ")> ByVal startStr As String, _
                <Values("'", "]", " ")> ByVal endStr As String, _
                <Values("bc", "bcd", "y")> ByVal expected As String)
            Assert.That(StringUtil.ExtractEnclosedString(target, startStr, endStr), [Is].EqualTo(expected))
        End Sub

        <Test(), Sequential()> Public Sub 始端終端に一致しなければ_nullを返す( _
                <Values("abcd", "abcd]e", "xyz")> ByVal target As String, _
                <Values("#", "[", " ")> ByVal startStr As String, _
                <Values("'", "]", " ")> ByVal endStr As String)
            Assert.That(StringUtil.ExtractEnclosedString(target, startStr, endStr), [Is].Null)
        End Sub

    End Class

    Public Class SplitNumberAlpha_ : Inherits StringUtilTest

        <Test(), Sequential()> Public Sub 数値とそれ以外に分割する( _
                <Values("abc", "12", "a1", "9x", "abc123xyz")> ByVal arg As String, _
                <Values("abc", "12", "a 1", "9 x", "abc 123 xyz")> ByVal expected As String)
            Dim actuals As String() = StringUtil.SplitNumberAlpha(arg)
            Assert.That(actuals, [Is].Not.Empty)
            Assert.That(Join(actuals, " "), [Is].EqualTo(expected))
        End Sub

        <Test(), Sequential()> Public Sub 空文字は長さ0の配列を返す()
            Dim actuals As String() = StringUtil.SplitNumberAlpha("")
            Assert.That(actuals, [Is].Not.Null)
            Assert.That(actuals.Length, [Is].EqualTo(0))
        End Sub

    End Class

    Public Class OmitIfLengthOverTest : Inherits StringUtilTest

        <Test(), Sequential()> Public Sub 最大長を超えていれば末尾トリプルドットで省略する( _
                <Values("abcdefg", "abcdefg", "abcdefg")> ByVal value As String, _
                <Values(7, 8, 5)> ByVal maxByteLength As Integer, _
                <Values("abcdefg", "abcdefg", "ab...")> ByVal expected As String)
            Assert.That(StringUtil.OmitIfLengthOver(value, maxByteLength), [Is].EqualTo(expected))
        End Sub

        <Test(), Sequential()> Public Sub 対象文字列が短ければ_末尾トリプルドットで省略せず_ただの切り捨て( _
                <Values("abc", "abc", "abc")> ByVal value As String, _
                <Values(3, 4, 2)> ByVal maxByteLength As Integer, _
                <Values("abc", "abc", "ab")> ByVal expected As String)
            Assert.That(StringUtil.OmitIfLengthOver(value, maxByteLength), [Is].EqualTo(expected))
        End Sub

        <Test(), Sequential()> Public Sub 最大長を超えていれば末尾トリプルドットで省略する_全角含む( _
                <Values("abcあいうdef", "abcあいうdef", "abcあいうdef", "abcあいうdef")> ByVal value As String, _
                <Values(12, 11, 10, 5)> ByVal maxByteLength As Integer, _
                <Values("abcあいうdef", "abcあい...", "abcあい...", "ab...")> ByVal expected As String)
            Assert.That(StringUtil.OmitIfLengthOver(value, maxByteLength), [Is].EqualTo(expected))
        End Sub

        <Test(), Sequential()> Public Sub 最大長を超えていれば末尾文字で省略する_全角の末尾文字でもOK( _
                <Values("abcdefg", "abcde", "あいう", "aいうえ")> ByVal value As String, _
                <Values("abc…", "abcde", "あ…", "aい…")> ByVal expected As String)
            Assert.That(StringUtil.OmitIfLengthOver(value, 5, "…"), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class ToUpper_nullはnothingあれば大文字で返す : Inherits StringUtilTest
        <Test()> Public Sub 半角小文字のときは半角大文字を返す()
            Assert.AreEqual(StringUtil.ToUpper("a"), "A")
        End Sub

        <Test()> Public Sub nullのときはnothingを返す()
            Assert.IsNull(StringUtil.ToUpper(Nothing))
        End Sub

        <Test()> Public Sub 半角大文字のときは半角大文字を返す()
            Assert.AreEqual(StringUtil.ToUpper("A"), "A")
        End Sub

        <Test()> Public Sub 全角小文字のときは全角大文字を返す()
            Assert.AreEqual(StringUtil.ToUpper("ａ"), "Ａ")
        End Sub

        <Test()> Public Sub 全角大文字のときは全角大文字を返す()
            Assert.AreEqual(StringUtil.ToUpper("Ａ"), "Ａ")
        End Sub

        <Test(), Sequential()> Public Sub 大文字の小文字の判別の無いものは全てそのままを返す(<Values("1", "１", "ｱ", "あ", "ア", "亜", "#")> ByVal value As String)
            Assert.AreEqual(StringUtil.ToUpper(value), value)
        End Sub
    End Class

    Public Class ToDateTimeString_常にyyyyMMddHmmssの文字列にする : Inherits StringUtilTest

        <Test(), Sequential()> Public Sub 年月日があれば_常にyyyyMMddHmmssの文字列になる( _
                <Values("2010/01/02 3:04:05", "2010/01/02 15:16:17", "2010/01/02 03:04:05")> ByVal arg As String, _
                <Values("2010/01/02 3:04:05", "2010/01/02 15:16:17", "2010/01/02 3:04:05")> ByVal expected As String)
            Dim sut As DateTime? = CDate(arg)
            Assert.That(StringUtil.ToDateTimeString(sut), [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub 引数NullならNull返す()
            Dim sut As DateTime? = Nothing
            Assert.That(StringUtil.ToDateTimeString(sut), [Is].Null)
        End Sub

        <Test(), Sequential()> Public Sub 年月日が無ければ_Hmmssの文字列になる( _
                <Values("3:4:5", "03:04:05")> ByVal arg As String, _
                <Values("3:04:05", "3:04:05")> ByVal expected As String)
            Dim sut As DateTime? = CDate(arg)
            Assert.That(StringUtil.ToDateTimeString(sut), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class ToZenkaku : Inherits StringUtilTest

        <Test(), Sequential()> Public Sub 半角文字が全角文字になる( _
                <Values("a", "Z", "1", "ｦ", " ", "ａ1ｂ2ｃ")> ByVal arg As String, _
                <Values("ａ", "Ｚ", "１", "ヲ", "　", "ａ１ｂ２ｃ")> ByVal expected As String)
            Assert.That(StringUtil.ToZenkaku(arg), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class ToHankakuTest : Inherits StringUtilTest

        <Test(), Sequential()> Public Sub 全角文字が半角文字になる( _
                <Values("ａ", "Ｚ", "１", "ヲ", "　", "ａ１ｂ２ｃ")> ByVal arg As String, _
                <Values("a", "Z", "1", "ｦ", " ", "a1b2c")> ByVal expected As String)
            Assert.That(StringUtil.ToHankaku(arg), [Is].EqualTo(expected))
        End Sub

        <Test(), Sequential()> Public Sub 半角文字は半角文字のまま( _
                <Values("a", "Z", "1", "ｦ", " ", "ａ1ｂ2ｃ3d")> ByVal arg As String, _
                <Values("a", "Z", "1", "ｦ", " ", "a1b2c3d")> ByVal expected As String)
            Assert.That(StringUtil.ToHankaku(arg), [Is].EqualTo(expected))
        End Sub

        <Test(), Sequential()> Public Sub 記号_ひらがな_漢字は_変換できる半角文字が無いから_変換しない( _
                <Values("①", "あ", "亜")> ByVal arg As String, _
                <Values("①", "あ", "亜")> ByVal expected As String)
            Assert.That(StringUtil.ToHankaku(arg), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class IsGaijiForJISX0208Test : Inherits StringUtilTest

        <Test(), Sequential()> Public Sub 機種依存文字と呼ばれる文字は真となる( _
                <Values("①"c, "㌢"c, "℡"c, "￤"c, "ⅷ"c, "✔"c, "€"c)> ByVal gaiji As Char)
            Assert.That(StringUtil.IsGaijiForJISX0208(gaiji), [Is].True)
        End Sub

        <Test(), Sequential()> Public Sub 機種依存文字と呼ばれない文字は偽となる( _
                <Values("a"c, "Z"c, "1"c, "ｦ"c, "?"c, " "c, "ａ"c, "Ｚ"c, "１"c, "あ"c, "ヲ"c, "　"c, "亜"c)> ByVal notGaiji As Char)
            Assert.That(StringUtil.IsGaijiForJISX0208(notGaiji), [Is].False)
        End Sub

    End Class

    Public Class 文字列からスペースを取り除いて返す : Inherits StringUtilTest

        <Test()> Public Sub 半角スペースを取り除く()
            Assert.AreEqual("1234", StringUtil.RemoveSpaceCharacters("12 3  4"))
        End Sub

        <Test()> Public Sub 全角スペースを取り除く()
            Assert.AreEqual("あいうえ", StringUtil.RemoveSpaceCharacters("あい　う　　え"))
        End Sub

        <Test()> Public Sub スペースが無ければそのまま返す()
            Assert.AreEqual("1234あいうえ", StringUtil.RemoveSpaceCharacters("1234あいうえ"))
        End Sub

        <Test()> Public Sub 引数NothingならNothingを返す()
            Assert.AreEqual(Nothing, StringUtil.RemoveSpaceCharacters(Nothing))
        End Sub
    End Class

    Public Class IsEmpty_PrimitiveValueObject_Test : Inherits StringUtilTest

        <Test()> Public Sub 自身がNullなら_真になる()
            Dim arg As StrPvo = Nothing
            Assert.That(StringUtil.IsEmpty(arg), [Is].True)
            Assert.That(StringUtil.IsNotEmpty(arg), [Is].False)
        End Sub

        <TestCase(Nothing)>
        <TestCase("")>
        <TestCase("  ")>
        Public Sub 値がNullなら_真になる(arg As String)
            Assert.That(StringUtil.IsEmpty(New StrPvo(arg)), [Is].True)
            Assert.That(StringUtil.IsNotEmpty(New StrPvo(arg)), [Is].False)
        End Sub

        <TestCase("aa")>
        <TestCase("ああ")>
        Public Sub 値があれば_偽になる(arg As String)
            Assert.That(StringUtil.IsEmpty(New StrPvo(arg)), [Is].False)
            Assert.That(StringUtil.IsNotEmpty(New StrPvo(arg)), [Is].True)
        End Sub

    End Class

    Public Class SplitBySizeTest : Inherits StringUtilTest

        <Test()>
        Public Sub 桁数ごとで文字列を分割できる()
            Assert.That(StringUtil.SplitBySize("abcdef", 2), [Is].EquivalentTo({"ab", "cd", "ef"}))
        End Sub

        <Test()>
        Public Sub 最大長を考慮して割り切れなくても_分割できる()
            Assert.That(StringUtil.SplitBySize("1234567", 4), [Is].EquivalentTo({"1234", "567"}))
        End Sub

        <Test()>
        Public Sub 空文字なら空の配列を返す()
            Assert.That(StringUtil.SplitBySize("", 3), [Is].EquivalentTo(New String() {}))
        End Sub

        <Test()>
        Public Sub 桁数が対象文字列を超える場合でも_配列として取得できる()
            Assert.That(StringUtil.SplitBySize("hohoge", 10), [Is].EquivalentTo({"hohoge"}))
        End Sub

    End Class

    Public Class ExtractNumberTest : Inherits StringUtilTest

        <TestCase("1,a", 1, "先頭の数値")>
        <TestCase("a2b", 2, "中間の数値")>
        <TestCase("a,3", 3, "末尾の数値")>
        <TestCase("444,a", 444, "3桁")>
        <TestCase("a55555b", 55555, "5桁")>
        <TestCase("a,666", 666, "3桁")>
        <TestCase("77,88a999", 77, "複数あっても最初に見つかった数値になる")>
        Public Sub 最初に見つかった数値を抜き出せる(text As String, expected As Integer, message As String)
            Assert.That(StringUtil.ExtractNumber(text), [Is].EqualTo(expected), message)
        End Sub

        <TestCase("一", "漢数字")>
        <TestCase("７７", "全角数値")>
        <TestCase("abc", "半角文字")>
        <TestCase("あいうえお", "全角文字")>
        Public Sub 半角数値が無ければ_例外にする(text As String, comment As String)
            Assert.That(Sub() StringUtil.ExtractNumber(text), Throws.TypeOf(Of ArgumentException).And.Message.StringStarting("値 '" & text & "' に数値は含まれない"))
        End Sub

    End Class

    Public Class ConvByteToWithSiPrefixTest : Inherits StringUtilTest

        <Test> Public Sub _1000バイト未満はそのまま表示する()
            Dim actual As String = StringUtil.ConvByteToWithSiPrefix(999UL)
            Assert.That(actual, [Is].EqualTo("999B"))
        End Sub

        <TestCase(1000UL, "1KB")>
        <TestCase(435132UL, "424.9KB")>
        <TestCase(15355148UL, "14.6MB")>
        <TestCase(18663173564UL, "17.4GB")>
        <TestCase(14353864841343UL, "13.1TB")>
        Public Sub バイト数をSi接頭語付き表記に変換できる(num As ULong, expected As String)
            Dim actual As String = StringUtil.ConvByteToWithSiPrefix(num)
            Assert.That(actual, [Is].EqualTo(expected))
        End Sub

        <Test> Public Sub TBより上は変換しない()
            Dim actual As String = StringUtil.ConvByteToWithSiPrefix(2251799813685248UL)
            Assert.That(actual, [Is].EqualTo("2048TB"))
        End Sub
    End Class

End Class
