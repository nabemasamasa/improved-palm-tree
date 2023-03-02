Imports System.Globalization
Imports NUnit.Framework

Public MustInherit Class DateUtilTest
    Private Function CCDate(obj As Object) As DateTime
        Return DateUtil.CCDate(obj)
    End Function

    Public Class ConvYyyymmddToSlashYyyymmdd : Inherits DateUtilTest

        <Test()> Public Sub 正常系()
            Assert.AreEqual("2010/07/07", DateUtil.ConvYyyymmddToSlashYyyymmdd("20100707"))
            Assert.AreEqual("2010/07/07", DateUtil.ConvYyyymmddToSlashYyyymmdd(20100707))
        End Sub

        <Test()> Public Sub Null引数ならNull()
            Dim s As String = Nothing
            Assert.IsNull(DateUtil.ConvYyyymmddToSlashYyyymmdd(s))

            Dim i As Nullable(Of Integer)
            Assert.IsNull(DateUtil.ConvYyyymmddToSlashYyyymmdd(i))
        End Sub

    End Class

    Public Class ConvHhmmssToColonHhmmss : Inherits DateUtilTest

        <Test()> Public Sub 正常系()
            Assert.AreEqual("01:23:45", DateUtil.ConvHhmmssToColonHhmmss("12345"))
            Assert.AreEqual("00:01:23", DateUtil.ConvHhmmssToColonHhmmss(123))
        End Sub

        <Test()> Public Sub Null引数ならNull()
            Dim s As String = Nothing
            Assert.IsNull(DateUtil.ConvHhmmssToColonHhmmss(s))

            Dim i As Nullable(Of Integer)
            Assert.IsNull(DateUtil.ConvHhmmssToColonHhmmss(i))
        End Sub

    End Class

    Public Class ConvTimeToColonHhmmssTest : Inherits DateUtilTest

        <Test(), Sequential()> Public Sub 日時を時分秒のコロン区切りにする(
                <Values("1:23:45", "2010/01/02 03:04:05")> ByVal dateTimeStr As String,
                <Values("01:23:45", "03:04:05")> ByVal expected As String)
            Assert.That(DateUtil.ConvTimeToColonHhmmss(CCDate(dateTimeStr)), [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub nullならnullを返す()
            Assert.That(DateUtil.ConvTimeToColonHhmmss(Nothing), [Is].Null)
        End Sub

    End Class

    Public Class ConvTimeToColonHhmmTest : Inherits DateUtilTest

        <Test(), Sequential()> Public Sub 日時を時分秒のコロン区切りにする(
                <Values("1:23:45", "2010/01/02 03:04:05")> ByVal dateTimeStr As String,
                <Values("01:23", "03:04")> ByVal expected As String)
            Assert.That(DateUtil.ConvTimeToColonHhmm(CCDate(dateTimeStr)), [Is].EqualTo(expected))
        End Sub

        <Test()> Public Sub nullならnullを返す()
            Assert.That(DateUtil.ConvTimeToColonHhmm(Nothing), [Is].Null)
        End Sub

    End Class

    Public Class ConvTimeToXxx : Inherits DateUtilTest

        <Test()> Public Sub HHMMSS形式へ変換できる()
            Assert.That(DateUtil.ConvTimeToInteger(CCDate("2018/01/02 03:04:05")), [Is].EqualTo(30405), "数値なので先頭の0は消える")
            Assert.That(DateUtil.ConvTimeToHhmmss(CCDate("2018/01/02 03:04:05")), [Is].EqualTo("030405"))
        End Sub

        <Test()> Public Sub HHMM形式へ変換できる()
            Assert.That(DateUtil.ConvTimeToHhmm(CCDate("2018/01/02 03:04:05")), [Is].EqualTo("0304"))
        End Sub
    End Class

    Public Class 通常Test : Inherits DateUtilTest

        <Test()> Public Sub TestConvYyyymmddToSlashYyyymmdd_Empty引数でもNull()
            Dim s As String = String.Empty
            Assert.IsNull(DateUtil.ConvYyyymmddToSlashYyyymmdd(s))
        End Sub

        <Test()> Public Sub TestConvYyyymmddToSlashYyyymmdd_99999999なら最大日付()
            Assert.AreEqual("9999/12/31", DateUtil.ConvYyyymmddToSlashYyyymmdd("99999999"))
        End Sub

        <Test()> Public Sub TestConvSlashYyyymmddToYyyymmdd()
            Assert.AreEqual("20100707", DateUtil.ConvSlashYyyymmddToYyyymmdd("2010/07/07"))
            Assert.AreEqual(20100707, DateUtil.ConvSlashYyyymmddToYyyymmddAsInteger("2010/07/07"))
        End Sub

        <Test()> Public Sub TestConvSlashYyyymmToYyyymm()
            Assert.AreEqual("201007", DateUtil.ConvSlashYyyymmToYyyymm("2010/07"))
            Assert.AreEqual(201007, DateUtil.ConvSlashYyyymmToYyyymmAsInteger("2010/07"))
        End Sub

        <Test()> Public Sub TestConvSlashYyyymmddToYyyymmdd_Null引数ならNull()
            Assert.IsNull(DateUtil.ConvSlashYyyymmddToYyyymmdd(Nothing))
            Assert.IsNull(DateUtil.ConvSlashYyyymmddToYyyymmddAsInteger(Nothing))
        End Sub
        <Test()> Public Sub TestConvSlashYyyymmddToYyyymmdd_Empty引数でもNull()
            Assert.IsNull(DateUtil.ConvSlashYyyymmddToYyyymmdd(""))
            Assert.IsNull(DateUtil.ConvSlashYyyymmddToYyyymmddAsInteger(String.Empty))
        End Sub

        <Test()> Public Sub TestConvSlashYyyymmToYyyymm_Null引数ならNull()
            Assert.IsNull(DateUtil.ConvSlashYyyymmToYyyymm(Nothing))
            Assert.IsNull(DateUtil.ConvSlashYyyymmToYyyymmAsInteger(Nothing))
        End Sub

        <Test()> Public Sub TestConvSlashYyyymmToYyyymm_Empty引数でもNull()
            Assert.IsNull(DateUtil.ConvSlashYyyymmToYyyymm(""))
            Assert.IsNull(DateUtil.ConvSlashYyyymmToYyyymmAsInteger(String.Empty))
        End Sub

        <Test()> Public Sub ConvSlashYyyymmddToYyyymmdd_最大日付()
            Assert.AreEqual("99999999", DateUtil.ConvSlashYyyymmddToYyyymmdd("9999/12/31"))
            Assert.AreEqual(99999999, DateUtil.ConvSlashYyyymmddToYyyymmddAsInteger("9999/12/31"))
        End Sub

    End Class

    Public Class ConvYyyymmddToDate : Inherits DateUtilTest

        <Test()> Public Sub 正常系()
            Assert.AreEqual(CCDate("2010/05/07"), DateUtil.ConvYyyymmddToDate("20100507"))
            Assert.AreEqual(CCDate("2010/04/07"), DateUtil.ConvYyyymmddToDate(20100407))
        End Sub

        <Test(), Sequential()> Public Sub 正常系_日時版(
                <Values("20100307", "20100408", "20100509")> ByVal yyyymmdd As String,
                <Values("321", Nothing, "0")> ByVal hhmmss As String,
                <Values("2010/03/07 0:03:21", "2010/04/08", "2010/05/09")> ByVal expected As String)
            Dim ymd As Integer? = Nothing
            Dim hms As Integer? = Nothing
            If IsNumeric(yyyymmdd) Then
                ymd = CInt(yyyymmdd)
            End If
            If IsNumeric(hhmmss) Then
                hms = CInt(hhmmss)
            End If
            Assert.That(DateUtil.ConvYyyymmddToDate(ymd, hms), [Is].EqualTo(CCDate(expected)))
        End Sub

        <Test()> Public Sub Null引数ならNull()
            Dim s As String = Nothing
            Assert.IsNull(DateUtil.ConvYyyymmddToDate(s))

            Dim i As Nullable(Of Integer)
            Assert.IsNull(DateUtil.ConvYyyymmddToDate(i))

            Assert.IsNull(DateUtil.ConvYyyymmddToDate(i, i))
        End Sub

        <Test()> Public Sub T99999999なら最大日付()
            Assert.AreEqual(#12/31/9999#, DateUtil.ConvYyyymmddToDate("99999999"))
            Assert.AreEqual(#12/31/9999#, DateUtil.ConvYyyymmddToDate(99999999))
            Assert.AreEqual(#12/31/9999 11:11:11 AM#, DateUtil.ConvYyyymmddToDate(99999999, 111111))
        End Sub

    End Class

    Public Class ConvDateValueToInteger_ : Inherits DateUtilTest

        <Test()> Public Sub ConvDateValueToInteger_Object型の変数にDateTime型()
            Dim aDate As Object = CCDate("2010/07/07")
            Assert.AreEqual(20100707, DateUtil.ConvDateValueToInteger(aDate))
        End Sub

        <Test()> Public Sub ConvDateValueToInteger_Object型の変数にDateTime型_最大日付()
            Dim aDate As Object = DateUtil.MAX_VALUE_DATETIME
            Assert.AreEqual(DateUtil.MAX_VALUE_INTEGER, DateUtil.ConvDateValueToInteger(aDate))
        End Sub

        <Test()> Public Sub ConvDateValueToInteger_Object型の変数にString型(<Values("2010/07/05", "2010-07-05")> ByVal aDate As Object)
            Assert.AreEqual(20100705, DateUtil.ConvDateValueToInteger(aDate), "スラッシュ区切りでもハイフン区切りでも大丈夫")
        End Sub

        <Test()> Public Sub ConvDateValueToInteger_Object型の変数にString型_最大日付()
            Dim aDate As Object = DateUtil.MAX_VALUE_WITH_SLASH_AS_STRING
            Assert.AreEqual(DateUtil.MAX_VALUE_INTEGER, DateUtil.ConvDateValueToInteger(aDate))
        End Sub

        <Test()> Public Sub ConvDateValueToInteger_Object型の変数に無効な日付書式のString型()
            Dim aDate As Object = "2010/07/"
            Assert.IsNull(DateUtil.ConvDateValueToInteger(aDate))
        End Sub

    End Class

    Public Class 通常Test2 : Inherits DateUtilTest

        <Test()> Public Sub ConvTimeToInteger_通常()
            Assert.AreEqual(123456, DateUtil.ConvTimeToInteger(CCDate("2010/02/03 12:34:56")))
        End Sub

        <Test()> Public Sub ConvObjectToDateTime_Object型の変数にDateTime型()
            Dim aDate As Object = CCDate("2010/07/07")
            Assert.AreEqual(CCDate("2010/07/07"), DateUtil.ConvDateValueToDateTime(aDate))
        End Sub

        <Test()> Public Sub ConvObjectToDateTime_Object型の変数にString型()
            Dim aDate As Object = "2010/07/05"
            Assert.AreEqual(CCDate("2010/07/05"), DateUtil.ConvDateValueToDateTime(aDate))
        End Sub

        <Test()> Public Sub ConvObjectToDateTime_Object型の変数に無効な日付書式のString型()
            Dim aDate As Object = "2010/07/"
            Assert.IsNull(DateUtil.ConvDateValueToDateTime(aDate))
        End Sub

        '<Test()> Public Sub ConvObjectToDateTime_Object型の変数にInteger型()
        '    Dim aDate As Object = 20100101
        '    Assert.AreEqual(CDate("2010/01/01"), DateUtil.ConvDateValueToDateTime(aDate))
        'End Sub


        <Test()> Public Sub ConvYyyymmddToYyyymm_通常()
            Assert.AreEqual(123456, DateUtil.ConvYyyymmddToYyyymm(12345678))
        End Sub

        <Test()> Public Sub ConvYyyymmddToYyyymm_NullならNullを返す()
            Dim aDate As Integer? = Nothing
            Assert.IsNull(DateUtil.ConvYyyymmddToYyyymm(aDate))
        End Sub

        <Test()> Public Sub ConvYyyymmddToYyyymm_8桁未満は例外発生()
            Try
                DateUtil.ConvYyyymmddToYyyymm(1234567)
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub ConvYyyymmToSlashYyyymm_通常()
            Dim aDate As Integer = 201003
            Assert.AreEqual("2010/03", DateUtil.ConvYyyymmToSlashYyyymm(aDate))
        End Sub

        <Test()> Public Sub ConvYyyymmToSlashYyyymm_NullならNullを返す()
            Assert.IsNull(DateUtil.ConvYyyymmToSlashYyyymm(Nothing))
        End Sub

        <Test()> Public Sub ConvYyyymmToSlashYyyymm_6桁未満は例外発生()
            Try
                DateUtil.ConvYyyymmToSlashYyyymm(12345)
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub ToTimeMillis_1970_01_01_からのミリ秒にして返す()
            Assert.AreEqual(0&, DateUtil.ToTimeMillis(DateUtil.CCDate("1970/01/01 0:00:00")))
            Assert.AreEqual(1 * 60 * 60 * 1000 + 1 * 60 * 1000 + 1 * 1000 + 1&, DateUtil.ToTimeMillis(DateUtil.CCDate("1970/01/01 1:01:01.001")))
            Assert.AreEqual((2 - 1) * 24 * 60 * 60 * 1000, DateUtil.ToTimeMillis(DateUtil.CCDate("1970/01/02 0:00:00")))
        End Sub

    End Class

    Public Class CompareNullsLast_Null値を最後に並べるComparer値を返す : Inherits DateUtilTest

        <Test()> Public Sub Null同士なら一致_0()
            Assert.AreEqual(0, DateUtil.CompareNullsLast(Nothing, Nothing))
        End Sub

        <Test()> Public Sub どちらかがNullなら_Null側が最後にソートされる値を返す()
            Assert.AreEqual(1, DateUtil.CompareNullsLast(Nothing, CCDate("12/12/12")), "> 0 だから1 - Null値のほうが大きい判定→最後にソート")
            Assert.AreEqual(-1, DateUtil.CompareNullsLast(CCDate("12/12/12"), Nothing), "< 0 だから-1 - Null値のほうが大きい判定→最後にソート")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート()
            Assert.AreEqual(-1, DateUtil.CompareNullsLast(CCDate("01/01/01"), CCDate("12/12/12")), "< 0 だから-1")
            Assert.AreEqual(1, DateUtil.CompareNullsLast(CCDate("12/12/12"), CCDate("01/01/01")), "> 0 だから1")
            Assert.AreEqual(0, DateUtil.CompareNullsLast(CCDate("01/01/01"), CCDate("01/01/01")), "= 0")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート_Object型()
            Assert.AreEqual(-1, DateUtil.CompareNullsLast("01/01/01", "12/12/12"), "< 0 だから-1")
            Assert.AreEqual(1, DateUtil.CompareNullsLast(CCDate("12/12/12"), "01/01/01"), "> 0 だから1")
            Assert.AreEqual(0, DateUtil.CompareNullsLast("01/01/01", CCDate("01/01/01")), "= 0")
        End Sub

        <Test()> Public Sub 実習_Null値は最後にソートする()
            Dim aList As New List(Of DateTime?)
            aList.AddRange(New DateTime?() {CCDate("2012/12/12"), CCDate("2001/01/01"), Nothing, CCDate("2010/11/11")})
            aList.Sort(Function(x As DateTime?, y As DateTime?) DateUtil.CompareNullsLast(x, y))

            Assert.AreEqual(4, aList.Count)
            Assert.AreEqual(CCDate("2001/01/01"), aList(0))
            Assert.AreEqual(CCDate("2010/11/11"), aList(1))
            Assert.AreEqual(CCDate("2012/12/12"), aList(2))
            Assert.IsNull(aList(3))
        End Sub
    End Class

    Public Class CompareDescNullsLast_Null値を最後に並べる降順Comparer値を返す : Inherits DateUtilTest

        <Test()> Public Sub Null同士なら一致_0()
            Assert.AreEqual(0, DateUtil.CompareDescNullsLast(Nothing, Nothing))
        End Sub

        <Test()> Public Sub どちらかがNullなら_Null側が最後にソートされる値を返す()
            Assert.AreEqual(1, DateUtil.CompareDescNullsLast(Nothing, CCDate("12/12/12")), "> 0 だから1 - Null値のほうが大きい判定→最後にソート")
            Assert.AreEqual(-1, DateUtil.CompareDescNullsLast(CCDate("12/12/12"), Nothing), "< 0 だから-1 - Null値のほうが大きい判定→最後にソート")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート()
            Assert.AreEqual(1, DateUtil.CompareDescNullsLast(CCDate("01/01/01"), CCDate("12/12/12")), "< 0 だけど降順Sortしたいから反対の1")
            Assert.AreEqual(-1, DateUtil.CompareDescNullsLast(CCDate("12/12/12"), CCDate("01/01/01")), "> 0 だけど降順Sortしたいから反対の-1")
            Assert.AreEqual(0, DateUtil.CompareDescNullsLast(CCDate("01/01/01"), CCDate("01/01/01")), "= 0")
        End Sub

        <Test()> Public Sub Null値でなければ_普通にソート_Object型()
            Assert.AreEqual(1, DateUtil.CompareDescNullsLast("01/01/01", "12/12/12"), "< 0 だけど降順Sortしたいから反対の1")
            Assert.AreEqual(-1, DateUtil.CompareDescNullsLast(CCDate("12/12/12"), "01/01/01"), "> 0 だけど降順Sortしたいから反対の-1")
            Assert.AreEqual(0, DateUtil.CompareDescNullsLast("01/01/01", CCDate("01/01/01")), "= 0")
        End Sub

        <Test()> Public Sub 実習_Null値は最後にソートする()
            Dim aList As New List(Of DateTime?)
            aList.AddRange(New DateTime?() {CCDate("2012/12/12"), CCDate("2001/01/01"), Nothing, CCDate("2010/11/11")})
            aList.Sort(Function(x As DateTime?, y As DateTime?) DateUtil.CompareDescNullsLast(x, y))

            Assert.AreEqual(4, aList.Count)
            Assert.AreEqual(CCDate("2012/12/12"), aList(0))
            Assert.AreEqual(CCDate("2010/11/11"), aList(1))
            Assert.AreEqual(CCDate("2001/01/01"), aList(2))
            Assert.IsNull(aList(3))
        End Sub
    End Class

    Public Class ConvDateValueToSlashYyyymmdd : Inherits DateUtilTest

        <Test()> Public Sub 正常系()
            Assert.AreEqual("2010/07/07", DateUtil.ConvDateValueToSlashYyyymmdd(CCDate("2010/07/07")))
            Assert.AreEqual("2010/06/06", DateUtil.ConvDateValueToSlashYyyymmdd("2010/06/06"))
        End Sub

        <Test()> Public Sub Null引数ならNull()
            Dim s As String = Nothing
            Assert.IsNull(DateUtil.ConvYyyymmddToSlashYyyymmdd(s))

            Dim d As DateTime? = Nothing
            Assert.IsNull(DateUtil.ConvYyyymmddToSlashYyyymmdd(s))
        End Sub

    End Class

    Public Class ConvYymmddToYyyymmdd : Inherits DateUtilTest

        <Test()> Public Sub 正常系()
            Assert.AreEqual(19950505, DateUtil.ConvYymmddToYyyymmdd(950505))
            Assert.AreEqual(19991231, DateUtil.ConvYymmddToYyyymmdd(991231))
            Assert.AreEqual(20000101, DateUtil.ConvYymmddToYyyymmdd(101), "000101だから2000/1/1")
            Assert.AreEqual(20020202, DateUtil.ConvYymmddToYyyymmdd(20202))
        End Sub

        <Test()> Public Sub Null引数ならNull()
            Assert.IsNull(DateUtil.ConvYymmddToYyyymmdd(DirectCast(Nothing, Integer?)))
        End Sub

        <TestCase("000101", 20000101)>
        <TestCase("290101", 20290101)>
        <TestCase("300101", 19300101)>
        <TestCase("990101", 19990101)>
        <TestCase("00040101", 40101)>
        <TestCase("18990101", 18990101)>
        <TestCase(Nothing, Nothing)>
        Public Sub 文字列引数ならよりわかりやすい(arg As String, expected As Integer?)
            Assert.That(DateUtil.ConvYymmddToYyyymmdd(arg), [Is].EqualTo(expected))
        End Sub

        <TestCase("101")>
        <TestCase("0101")>
        <TestCase("1201")>
        Public Sub 文字列引数のときは年を明示しないとエラー(arg As String)
            Try
                DateUtil.ConvYymmddToYyyymmdd(arg)
                Assert.Fail()
            Catch expected As ArgumentException
                Assert.That(expected.Message, [Is].EqualTo("年が不明:" & arg))
            End Try
        End Sub

    End Class

    Public Class ConvYyyymmToDate : Inherits DateUtilTest

        <Test(), Sequential()> Public Sub 正常系(
                <Values("201402", "201412", "201501")> ByVal yyyymm As String,
                <Values("2014/2/1", "2014/12/1", "2015/1/1")> ByVal expected As String)
            Assert.That(DateUtil.ConvYyyymmToDate(yyyymm), [Is].EqualTo(CCDate(expected)))
        End Sub

        <Test(), Sequential()> Public Sub 異常ならnullが返る(
                <Values("ABC", "20141215", "2014", Nothing)> ByVal yyyymm As String)
            Assert.That(DateUtil.ConvYyyymmToDate(yyyymm), [Is].Null)
        End Sub

        <Test(), Sequential()> Public Sub 異常でsystemYyyymmIfInvalidがtrueならシステム日付の1日が返る_(
                <Values("ABC", "20141215", "2014", Nothing)> ByVal yyyymm As String)
            Assert.That(DateUtil.ConvYyyymmToDate(yyyymm, systemYyyymmIfInvalid:=True), [Is].EqualTo(CCDate(String.Format("{0}/{1}/01", Date.Now.Year, Date.Now.Month))))
        End Sub

    End Class

    Public Class ConvDateToXxxxx_日付値を変換する : Inherits DateUtilTest

        <Test()> Public Sub ConvDateToInteger_通常()
            Assert.AreEqual(20100201, DateUtil.ConvDateToInteger(CCDate("2010/02/01")))
            Assert.AreEqual(20100203, DateUtil.ConvDateToInteger(CCDate("2010/02/03 12:34:56")), "時分秒は無視")
        End Sub

        <Test()> Public Sub ConvDateToInteger_最大日付は99999999になる()
            Assert.AreEqual(99999999, DateUtil.ConvDateToInteger(#12/31/9999#))
            Assert.AreEqual(99999999, DateUtil.ConvDateToInteger(#12/31/9999 12:34:56 PM#), "時分秒は無視")
        End Sub

        <Test()> Public Sub ConvDateToInteger_NullならNull()
            Assert.IsNull(DateUtil.ConvDateToInteger(Nothing))
        End Sub

        <Test()> Public Sub ConvDateToSlashYyyymmdd_通常()
            Assert.AreEqual("2012/03/03", DateUtil.ConvDateToSlashYyyymmdd(CCDate("2012/03/03")))
            Assert.AreEqual("2012/04/04", DateUtil.ConvDateToSlashYyyymmdd(CCDate("2012/04/04 05:05:05")), "時分秒は無視")
        End Sub

        <Test()> Public Sub ConvDateToSlashYyyymmdd_最大日付()
            Assert.AreEqual("9999/12/31", DateUtil.ConvDateToSlashYyyymmdd(#12/31/9999#))
            Assert.AreEqual("9999/12/31", DateUtil.ConvDateToSlashYyyymmdd(#12/31/9999 12:34:56 PM#), "時分秒は無視")
        End Sub

        <Test()> Public Sub ConvDateToSlashYyyymmdd_NullならNull()
            Assert.IsNull(DateUtil.ConvDateToSlashYyyymmdd(Nothing))
        End Sub

        <Test()> Public Sub ConvDateToHyphenYyyymmdd_通常()
            Assert.That(DateUtil.ConvDateToHyphenYyyymmdd(CCDate("2012/03/03")), [Is].EqualTo("2012-03-03"))
            Assert.That(DateUtil.ConvDateToHyphenYyyymmdd(CCDate("2012/04/04 05:05:05")), [Is].EqualTo("2012-04-04"), "時分秒は無視")
        End Sub

        <Test()> Public Sub ConvDateToHyphenYyyymmdd_最大日付()
            Assert.That(DateUtil.ConvDateToHyphenYyyymmdd(#12/31/9999#), [Is].EqualTo("9999-12-31"))
            Assert.That(DateUtil.ConvDateToHyphenYyyymmdd(#12/31/9999 12:34:56 PM#), [Is].EqualTo("9999-12-31"), "時分秒は無視")
        End Sub

        <Test()> Public Sub ConvDateToHyphenYyyymmdd_NullならNull()
            Assert.That(DateUtil.ConvDateToHyphenYyyymmdd(Nothing), [Is].Null)
        End Sub

        <Test()> Public Sub ConvDateToYyyymmdd_通常()
            Assert.AreEqual("20120303", DateUtil.ConvDateToYyyymmdd(CCDate("2012/03/03")))
            Assert.AreEqual("20120404", DateUtil.ConvDateToYyyymmdd(CCDate("2012/04/04 05:05:05")), "時分秒は無視")
        End Sub

        <Test()> Public Sub ConvDateToYyyymmdd_最大日付は99999999になる()
            Assert.AreEqual("99999999", DateUtil.ConvDateToYyyymmdd(#12/31/9999#))
            Assert.AreEqual("99999999", DateUtil.ConvDateToYyyymmdd(#12/31/9999 12:34:56 PM#), "時分秒は無視")
        End Sub

        <Test()> Public Sub ConvDateToYyyymmdd_NullならNull()
            Assert.IsNull(DateUtil.ConvDateToYyyymmdd(Nothing))
        End Sub

        <Test()> Public Sub ConvDateToYyyymm_通常()
            Assert.AreEqual("201203", DateUtil.ConvDateToYyyymm(CCDate("2012/03/03")))
            Assert.AreEqual("201204", DateUtil.ConvDateToYyyymm(CCDate("2012/04/04 05:05:05")), "時分秒は無視")
        End Sub

        <Test()> Public Sub ConvDateToYyyymm_最大日付は999999になる()
            Assert.AreEqual("999999", DateUtil.ConvDateToYyyymm(#12/31/9999#))
            Assert.AreEqual("999999", DateUtil.ConvDateToYyyymm(#12/31/9999 12:34:56 PM#), "時分秒は無視")
        End Sub

        <Test()> Public Sub ConvDateToYyyymm_NullならNull()
            Assert.IsNull(DateUtil.ConvDateToYyyymm(Nothing))
        End Sub

    End Class

    Public Class ConvDateTimeToXxx : Inherits DateUtilTest

        <TestCase("2019/10/10 10:10:00", "10月10日（木） 10:10", "10月10日（Thu） 10:10")>
        <TestCase("2020/04/04 08:05:00", "4月4日（土） 08:05", "4月4日（Sat） 08:05")>
        <TestCase("2020/04/04 08:05:59", "4月4日（土） 08:05", "4月4日（Sat） 08:05")>
        <TestCase("2020/04/04", "4月4日（土） 00:00", "4月4日（Sat） 00:00")>
        Public Sub ConvDateTimeToMddowhhmm_通常(ByVal dateTimeStr As String, ByVal expected As String, ByVal expectedEng As String)
            Assert.That(DateUtil.ConvDateTimeToMddowhhmm(CCDate(dateTimeStr)), [Is].EqualTo(expected).Or.EqualTo(expectedEng))
        End Sub

        <Test()> Public Sub ConvDateTimeToMddowhhmm_NullならNull()
            Assert.That(DateUtil.ConvDateTimeToMddowhhmm(Nothing), [Is].Null)
        End Sub

    End Class

    Public Class AddXxxxx_加算した結果を返す : Inherits DateUtilTest
        <Test()> Public Sub 通常_年()
            Assert.AreEqual(CCDate("2013/03/03"), DateUtil.AddYears(CCDate("2012/03/03"), 1))
        End Sub

        <Test()> Public Sub 通常_日()
            Assert.AreEqual(CCDate("2012/03/04"), DateUtil.AddDays(CCDate("2012/03/03"), 1))
        End Sub

        <Test()> Public Sub 通常_一年前()
            Assert.AreEqual(CCDate("2011/03/03"), DateUtil.AddYears(CCDate("2012/03/03"), -1))
        End Sub

        <Test()> Public Sub 通常_一週間前()
            Assert.AreEqual(CCDate("2012/02/25"), DateUtil.AddDays(CCDate("2012/03/03"), -7))
        End Sub

        <Test()> Public Sub 通常_加算年が0の場合はなにもしない()
            Assert.AreEqual(CCDate("2012/03/03"), DateUtil.AddYears(CCDate("2012/03/03"), 0))
        End Sub

        <Test()> Public Sub 通常_加算日が0の場合は何もしない()
            Assert.AreEqual(CCDate("2012/03/03"), DateUtil.AddDays(CCDate("2012/03/03"), 0))
        End Sub

        <Test()> Public Sub 通常_加算年がNullの場合はなにもしない()
            Assert.AreEqual(CCDate("2012/03/03"), DateUtil.AddYears(CCDate("2012/03/03"), Nothing))
        End Sub

        <Test()> Public Sub 通常_加算日がNullの場合は何もしない()
            Assert.AreEqual(CCDate("2012/03/03"), DateUtil.AddDays(CCDate("2012/03/03"), Nothing))
        End Sub

        <Test()> Public Sub 通常_年_NullならNull()
            Assert.IsNull(DateUtil.AddYears(Nothing, 1))
        End Sub

        <Test()> Public Sub 通常_日_NullならNull()
            Assert.IsNull(DateUtil.AddDays(Nothing, 1))
        End Sub

    End Class

    Public Class MakeEndOfLastBusinessPeriod_直前の半期末にする : Inherits DateUtilTest

        <Test(), Sequential()> Public Sub 指定日未満の半期末へ_指定日が半期末なら前半期末になる(
                <Values("2015/4/1", "2015/3/31", "2014/10/1", "2014/9/30")> ByVal dateTheToMarch As String,
                <Values("2015/3/31", "2014/9/30", "2014/9/30", "2014/3/31")> ByVal expectDate As String)
            Assert.That(DateUtil.MakeEndOfLastBusinessPeriod(CCDate(dateTheToMarch)), [Is].EqualTo(CCDate(expectDate)))
        End Sub

    End Class

    Public Class MakeStartOfCurrentBusinessPeriodTest : Inherits DateUtilTest

        <TestCase(20110401, 20110401)>
        <TestCase(20110930, 20110401)>
        <TestCase(20111001, 20111001)>
        <TestCase(20111231, 20111001)>
        <TestCase(20120101, 20111001)>
        <TestCase(20120331, 20111001)>
        Public Sub GetDateBeginCurrentPeriod_指定日付の当期の開始日付を取得する(yyyymmdd As Integer, expected As Integer)
            Assert.That(DateUtil.MakeStartOfCurrentBusinessPeriod(yyyymmdd), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class TruncateEndOfBusinessPeriod_半期末に切り捨てる : Inherits DateUtilTest

        <Test()> Public Sub _3月末に切り捨てられる(<Values(20150331, 20150401, 20150915, 20150929)> ByVal dateTheToMarch As Integer)
            Assert.That(DateUtil.TruncateEndOfBusinessPeriod(DateUtil.ConvYyyymmddToDate(dateTheToMarch).Value), [Is].EqualTo(CCDate("2015/03/31")))
        End Sub

        <Test()> Public Sub _9月末に切り捨てられる(<Values(20140930, 20141001, 20150101, 20150330)> ByVal dateTheToMarch As Integer)
            Assert.That(DateUtil.TruncateEndOfBusinessPeriod(DateUtil.ConvYyyymmddToDate(dateTheToMarch).Value), [Is].EqualTo(CCDate("2014/09/30")))
        End Sub

    End Class

    Public Class CCDate_常にYMDの年月日の並びで解釈するCDate : Inherits DateUtilTest

        <Test()> Sub システム設定が和暦でも_西暦から日付型を生成して_西暦を取得できる(<Values(1988, 1989, 2000, 2001)> ByVal year As Integer)
            Assert.That(DateUtil.CCDate(year & "/11/22"), [Is].EqualTo(New Date(year, 11, 22)))
            Assert.That(DateUtil.CCDate(year & "/02/03 04:05:06"), [Is].EqualTo(New Date(year, 2, 3, 4, 5, 6)))
            Assert.That(DateUtil.CCDate(year & "/03/04 05:06:07.888"), [Is].EqualTo(New Date(year, 3, 4, 5, 6, 7, 888)))
        End Sub

        <Test()> Sub システム設定が和暦でも_年が2桁以下でも西暦として_日付型を生成する(<Values("1", "20", "29", "30", "0003")> ByVal year As String)
            Assert.That(DateUtil.CCDate(year & "/11/22"), [Is].EqualTo(New Date(DateUtil.CorrectTo4DigitYear(year), 11, 22)))
            Assert.That(DateUtil.CCDate(year & "/02/03 04:05:06"), [Is].EqualTo(New Date(DateUtil.CorrectTo4DigitYear(year), 2, 3, 4, 5, 6)))
            Assert.That(DateUtil.CCDate(year & "/03/04 05:06:07.888"), [Is].EqualTo(New Date(DateUtil.CorrectTo4DigitYear(year), 3, 4, 5, 6, 7, 888)))
        End Sub

        <Test()> Sub システム設定が和暦でも_月日だけの指定なら_今年として日付型を生成する()
            Assert.That(DateUtil.CCDate("10/13"), [Is].EqualTo(New Date(DateTime.Now.Year, 10, 13)))
        End Sub

        <Test()> Sub 注意_システム設定が和暦だと_yyyy指定は和暦になる()
            Assert.That(DateUtil.CCDate("2000/11/22").ToString("yyyy"), [Is].EqualTo(If(0 <> DateUtil.YEAR_DIFF_OF_JAPANESE_CALENDAR, "12", "2000")))
        End Sub

        <Test()> Sub 時分指定のときは_年月日00010101になる()
            Assert.That(DateUtil.CCDate("10:13"), [Is].EqualTo(New Date(1, 1, 1, 10, 13, 0)))
            Assert.That(DateUtil.CCDate("10:13"), [Is].EqualTo(CDate("10:13")))
        End Sub

        <Test()> Sub null指定のときは_00010101_000000になる()
            Assert.That(DateUtil.CCDate(Nothing), [Is].EqualTo(New Date(1, 1, 1, 0, 0, 0)))
            Assert.That(DateUtil.CCDate(Nothing), [Is].EqualTo(CDate(Nothing)))
        End Sub

        <Test()> Sub 不正な日付書式は例外になる(
                <Values("2012/02/", "/12/01", "/03/", "10", "10:11:", ":12:", "04/10 10:11")> ByVal invalidDateTime As String)
            Try
                DateUtil.CCDate(invalidDateTime)
                Assert.Fail()
            Catch expected As InvalidCastException
                Assert.That(expected.Message, [Is].EqualTo(String.Format("String ""{0}"" から型 'Date' への変換は無効です。", invalidDateTime)) _
                            .Or.EqualTo(String.Format("Conversion from string ""{0}"" to type 'Date' is not valid.", invalidDateTime)), "言語「日本語」と言語「英語」のメッセージ")
            End Try
        End Sub

    End Class

    Public Class IsDateTest : Inherits DateUtilTest

        <TestCase("04/05", True, "MM/ddはok")>
        <TestCase("04:05", True, "hh:mmもok")>
        <TestCase("02/03 04:05", False, "MM/dd hh:mmはNGの謎")>
        Public Sub Learning_IsDateの挙動確認(dateTimeStr As String, expected As Boolean, comment As String)
            Assert.That(IsDate(dateTimeStr), [Is].EqualTo(expected), comment)
        End Sub

        <TestCase("2020/10/20", True)>
        <TestCase("10/20", True)>
        <TestCase("10/11/12", True)>
        <TestCase("20/11/22", True)>
        <TestCase("80/11/22", True)>
        <TestCase("2020/20/20", False)>
        <TestCase("20/20", False)>
        <TestCase("13/20/13", False)>
        <TestCase("10:11:12.333", True)>
        <TestCase("10:11:12", True)>
        <TestCase("10:11", True)>
        <TestCase("30:11:12.333", False)>
        <TestCase("30:11:12", False)>
        <TestCase("30:11", False)>
        <TestCase("2020/10/20 10:11", True)>
        <TestCase("2020/10/20 10:11:12", True)>
        <TestCase("2020/10/20 10:11:12.333", True)>
        <TestCase("80/10/20 10:11", True)>
        <TestCase("80/10/20 10:11:12", True)>
        <TestCase("80/10/20 10:11:12.333", True)>
        <TestCase("10/20 10:11", False)>
        <TestCase("10/20 10:11:12", False)>
        <TestCase("10/20 10:11:12.333", False)>
        <TestCase("2020/10/20 25:11", False)>
        <TestCase("2020/10/20 25:11:12", False)>
        <TestCase("2020/10/20 25:11:12.333", False)>
        <TestCase("2020/10/20 10:11:12:09", False)>
        <TestCase("February 12, 1969", True)>
        <TestCase("4:35:47 PM", True)>
        <TestCase("February 12, 1969 4:35:47 PM", True)>
        Public Sub 年月日は常にYMDの並びとしてIsDate判定をする(arg As String, expected As Boolean)
            Assert.That(DateUtil.IsDate(arg), [Is].EqualTo(expected))
        End Sub

    End Class

    <Ignore()>
    Public Class DateTimeパフォーマンス : Inherits EzUtilTest

        Private Function Average(ByVal times As List(Of TimeSpan)) As Double
            Dim result As Double
            For Each time As TimeSpan In times
                result += time.TotalMilliseconds
            Next
            Return result / times.Count
        End Function

        <Test()> Public Sub Hoge()
            Dim result As New Dictionary(Of String, List(Of TimeSpan))
            Dim emptyTime As New List(Of TimeSpan)
            result.Add("DateTime.Now", New List(Of TimeSpan))
            result.Add("DateTime.UtcNow", New List(Of TimeSpan))
            result.Add("DateUtil.GetNowDateTime", New List(Of TimeSpan))

            Const COUNT As Integer = 1024 * 1024
            Dim sw As New Stopwatch

            For j As Integer = 0 To 5
                sw.Reset()
                sw.Start()
                For i As Integer = 0 To COUNT - 1
                Next
                sw.Stop()
                emptyTime.Add(sw.Elapsed)

                sw.Reset()
                sw.Start()
                For i As Integer = 0 To COUNT - 1
                    Dim d As DateTime = DateTime.Now
                Next
                sw.Stop()
                result("DateTime.Now").Add(sw.Elapsed)

                sw.Reset()
                sw.Start()
                For i As Integer = 0 To COUNT - 1
                    Dim d As DateTime = DateTime.UtcNow
                Next
                sw.Stop()
                result("DateTime.UtcNow").Add(sw.Elapsed)

                sw.Reset()
                sw.Start()
                For i As Integer = 0 To COUNT - 1
                    Dim d As DateTime = DateUtil.GetNowDateTime
                Next
                sw.Stop()
                result("DateUtil.GetNowDateTime").Add(sw.Elapsed)
            Next

            emptyTime.RemoveAt(0)
            Dim emptyAverage As Double = Average(emptyTime)
            For Each pair As KeyValuePair(Of String, List(Of TimeSpan)) In result
                pair.Value.RemoveAt(0)
                Dim summary As Double = Average(pair.Value) - emptyAverage
                Dim per As Double = summary / COUNT
                Console.WriteLine(String.Format("{0} | {1}(ms) | {2}(ms)", pair.Key, summary, per * 1000))
            Next
        End Sub

    End Class

    Public Class BusinessYearAndPeriodTest : Inherits DateUtilTest

        <TestCase("2018/03/01", 2017)>
        <TestCase("2018/04/01", 2018)>
        <TestCase("2018/09/01", 2018)>
        <TestCase("2018/10/01", 2018)>
        <TestCase("2019/01/01", 2018)>
        <TestCase("2019/04/01", 2019)>
        Public Sub 年度を特定できる(ByVal aDate As DateTime, ByVal expected As Integer)
            Assert.That(DateUtil.DetectBusinessYear(aDate), [Is].EqualTo(expected))
        End Sub

        <TestCase("2018/03/01", DateUtil.BusinessPeriod.SECOND_PERIOD)>
        <TestCase("2018/04/01", DateUtil.BusinessPeriod.FIRST_PERIOD)>
        <TestCase("2018/09/01", DateUtil.BusinessPeriod.FIRST_PERIOD)>
        <TestCase("2018/10/01", DateUtil.BusinessPeriod.SECOND_PERIOD)>
        <TestCase("2019/01/01", DateUtil.BusinessPeriod.SECOND_PERIOD)>
        <TestCase("2019/04/01", DateUtil.BusinessPeriod.FIRST_PERIOD)>
        Public Sub 期を特定できる(ByVal aDate As DateTime, ByVal expected As DateUtil.BusinessPeriod)
            Assert.That(DateUtil.DetectBusinessPeriod(aDate), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class IsThisBusinessPeriodTest : Inherits DateUtilTest

        <TestCase("2018/04/01", "2018/09/24")>
        <TestCase("2018/09/01", "2018/09/24")>
        Public Sub 今期ならTrue_DateTime型(ByVal aDate As DateTime, ByVal now As DateTime)
            Assert.True(DateUtil.IsThisBusinessPeriod(aDate, now))
        End Sub

        <TestCase("2018/03/01", "2018/09/24")>
        <TestCase("2018/10/01", "2018/09/24")>
        <TestCase("2019/01/01", "2018/09/24")>
        <TestCase("2019/04/01", "2018/09/24")>
        Public Sub 今期でないならFalse_DateTime型(ByVal aDate As DateTime, ByVal now As DateTime)
            Assert.False(DateUtil.IsThisBusinessPeriod(aDate, now))
        End Sub

        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2018/04/01")>
        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2018/09/30")>
        Public Sub 今期ならTrue_年度と期で判定(ByVal businessYear As Integer, ByVal businessPeriod As DateUtil.BusinessPeriod, ByVal now As DateTime)
            Assert.True(DateUtil.IsThisBusinessPeriod(businessYear, businessPeriod, now))
        End Sub

        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2018/03/31")>
        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2018/10/01")>
        <TestCase(2018, DateUtil.BusinessPeriod.SECOND_PERIOD, "2018/09/30")>
        <TestCase(2018, DateUtil.BusinessPeriod.SECOND_PERIOD, "2019/04/01")>
        Public Sub 今期でないならFalse_年度と期で判定(ByVal businessYear As Integer, ByVal businessPeriod As DateUtil.BusinessPeriod, ByVal now As DateTime)
            Assert.False(DateUtil.IsThisBusinessPeriod(businessYear, businessPeriod, now))
        End Sub

    End Class

    Public Class IsPrevBusinessPeriodTest : Inherits DateUtilTest

        <TestCase("2017/12/31", "2018/09/24")>
        <TestCase("2018/01/01", "2018/09/24")>
        <TestCase("2018/03/31", "2018/09/24")>
        Public Sub 前期ならTrue_DateTime型で判定(ByVal aDate As DateTime, ByVal now As DateTime)
            Assert.True(DateUtil.IsPrevBusinessPeriod(aDate, now))
        End Sub

        <TestCase("2018/04/01", "2018/09/24")>
        <TestCase("2018/09/01", "2018/09/24")>
        <TestCase("2018/10/01", "2018/09/24")>
        <TestCase("2019/01/01", "2018/09/24")>
        <TestCase("2019/04/01", "2018/09/24")>
        Public Sub 前期でないならFalse_DateTime型で判定(ByVal aDate As DateTime, ByVal now As DateTime)
            Assert.False(DateUtil.IsPrevBusinessPeriod(aDate, now))
        End Sub

        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2018/10/01")>
        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2019/03/31")>
        Public Sub 前期ならTrue_年度と期で判定(ByVal businessYear As Integer, ByVal businessPeriod As DateUtil.BusinessPeriod, ByVal now As DateTime)
            Assert.True(DateUtil.IsPrevBusinessPeriod(businessYear, businessPeriod, now))
        End Sub

        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2018/09/30")>
        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2019/04/01")>
        <TestCase(2018, DateUtil.BusinessPeriod.SECOND_PERIOD, "2019/03/31")>
        <TestCase(2018, DateUtil.BusinessPeriod.SECOND_PERIOD, "2019/10/01")>
        Public Sub 前期でないならFalse_年度と期で判定(ByVal businessYear As Integer, ByVal businessPeriod As DateUtil.BusinessPeriod, ByVal now As DateTime)
            Assert.False(DateUtil.IsPrevBusinessPeriod(businessYear, businessPeriod, now))
        End Sub

    End Class

    Public Class IsNextBusinessPeriodTest : Inherits DateUtilTest

        <TestCase("2018/10/01", "2018/09/24")>
        <TestCase("2019/01/01", "2018/09/24")>
        <TestCase("2019/03/31", "2018/09/24")>
        Public Sub 来期ならTrue_DateTime型(ByVal aDate As DateTime, ByVal now As DateTime)
            Assert.True(DateUtil.IsNextBusinessPeriod(aDate, now))
        End Sub

        <TestCase("2018/04/01", "2018/09/24")>
        <TestCase("2018/09/30", "2018/09/24")>
        <TestCase("2019/04/01", "2018/09/24")>
        <TestCase("2019/09/30", "2018/09/24")>
        Public Sub 来期でないならFalse_DateTime型(ByVal aDate As DateTime, ByVal now As DateTime)
            Assert.False(DateUtil.IsNextBusinessPeriod(aDate, now))
        End Sub

        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2017/10/01")>
        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2018/03/31")>
        Public Sub 来期ならTrue_年度と期で判定(ByVal businessYear As Integer, ByVal businessPeriod As DateUtil.BusinessPeriod, ByVal now As DateTime)
            Assert.True(DateUtil.IsNextBusinessPeriod(businessYear, businessPeriod, now))
        End Sub

        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2017/09/30")>
        <TestCase(2018, DateUtil.BusinessPeriod.FIRST_PERIOD, "2018/04/01")>
        <TestCase(2018, DateUtil.BusinessPeriod.SECOND_PERIOD, "2018/03/31")>
        <TestCase(2018, DateUtil.BusinessPeriod.SECOND_PERIOD, "2018/10/01")>
        Public Sub 来期でないならFalse_年度と期で判定(ByVal businessYear As Integer, ByVal businessPeriod As DateUtil.BusinessPeriod, ByVal now As DateTime)
            Assert.False(DateUtil.IsNextBusinessPeriod(businessYear, businessPeriod, now))
        End Sub

    End Class

    Public Class ConvTimeToTotalSecondsTest : Inherits DateUtilTest

        <TestCase("1:23:45", (1 * 60 * 60) + (23 * 60) + 45)>
        <TestCase("01:23:45", (1 * 60 * 60) + (23 * 60) + 45)>
        <TestCase("1:23", (1 * 60 * 60) + (23 * 60))>
        <TestCase("1:2:3", (1 * 60 * 60) + (2 * 60) + 3)>
        Public Sub 日時の時分秒合計がとれる(ByVal dateTimeStr As String, ByVal expected As Integer)
            Assert.That(DateUtil.ConvTimeToTotalSeconds(dateTimeStr), [Is].EqualTo(expected))
        End Sub

        <TestCase(Nothing)>
        <TestCase("")>
        <TestCase("hoge")>
        <TestCase("8:")>
        <TestCase("1")>
        <TestCase("::")>
        Public Sub 日時変換できないものはエラーになる(ByVal arg As String)
            Try
                DateUtil.ConvTimeToTotalSeconds(arg)
                Assert.Fail()
            Catch ex As ArgumentException
                Assert.That(ex.Message, [Is].EqualTo("時刻を指定すべき"))
            End Try

        End Sub

    End Class

    Public Class TruncMillisecondsTest : Inherits DateUtilTest

        <Test()> Public Sub ミリ秒未満も切捨てできる()
            Const expected As Date = #6/17/2019 8:38:56 AM#
            Const ticks As Long = 636963575360966120L
            Dim arg As New Date(ticks)
            Assert.That(arg, [Is].Not.EqualTo(expected), "ミリ秒以下不一致")
            Assert.That(arg.AddMilliseconds(arg.Millisecond * -1), [Is].Not.EqualTo(expected), "ミリ秒未満が不一致らしい")
            Assert.That(DateUtil.TruncMilliseconds(arg), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class CorrectTo4DigitYearTest : Inherits DateUtilTest

        <TestCase("00", 2000)>
        <TestCase("1", 2001)>
        <TestCase("29", 2029)>
        <TestCase("30", 1930)>
        <TestCase("69", 1969)>
        <TestCase("0003", 3)>
        <TestCase("0029", 29)>
        <TestCase("1901", 1901)>
        Sub 年を4桁年に補正できる(year As String, expected As Integer)
            Assert.That(DateUtil.CorrectTo4DigitYear(year), [Is].EqualTo(expected))
        End Sub

    End Class

    Public Class RoundMillSecondsTest : Inherits DateUtilTest

        <TestCase("12:34:56.123", "12:34:56.000")>
        <TestCase("12:34:56.987", "12:34:57.000")>
        <TestCase("12:34:56.500", "12:34:57.000")>
        <TestCase("12:34:56.499", "12:34:56.000")>
        <TestCase("12:34:56.000", "12:34:56.000")>
        Public Sub ミリ秒を四捨五入して秒に加算できる(aTime As String, expected As String)
            Dim actual As DateTime? = DateUtil.RoundMillSeconds(DateTime.Parse("2020/01/02 " & aTime))
            Assert.That(expected, [Is].EqualTo(String.Format("{0:hh:mm:ss.fff}", actual)))
        End Sub

    End Class

End Class
