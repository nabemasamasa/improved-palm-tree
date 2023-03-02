''' <summary>
''' 日付操作のユーティリティ
''' </summary>
''' <remarks></remarks>
Public Class DateUtil

#Region "Nested classes..."
    ''' <summary>期に関する情報</summary>
    Public Enum BusinessPeriod
        ''' <summary>上期</summary>
        FIRST_PERIOD = 1
        ''' <summary>下期</summary>
        SECOND_PERIOD = 2
    End Enum
    Private Class MyCultureInfo
        ''' <summary>常に年を西暦として、年月日の並びはYMD固定としたCultureInfo</summary>
        Public Shared ReadOnly YMD_TIME As System.Globalization.CultureInfo
        Shared Sub New()
            YMD_TIME = New System.Globalization.CultureInfo("ja-JP", False)
            ' OSが和暦設定でも西暦として認識させる
            YMD_TIME.DateTimeFormat.Calendar = New System.Globalization.GregorianCalendar
        End Sub
    End Class
#End Region

    ''' <summary>年月日の最大表現(Integer)</summary>
    Public Const MAX_VALUE_INTEGER As Integer = 99999999
    ''' <summary>年月日の最大表現(スラッシュ区切り)</summary>
    Public Const MAX_VALUE_WITH_SLASH_AS_STRING As String = "9999/12/31"
    ''' <summary>年月日の最大表現(ハイフン区切り)</summary>
    Public Const MAX_VALUE_WITH_HYPHEN_AS_STRING As String = "9999-12-31"
    ''' <summary>年月日の最大表現(DateTime)</summary>
    Public Shared ReadOnly MAX_VALUE_DATETIME As DateTime = #12/31/9999#
    ''' <summary>和暦と西暦の差分 ※システム設定が西暦だと常に0</summary>
    Public Shared ReadOnly YEAR_DIFF_OF_JAPANESE_CALENDAR As Integer = CDate("100/1/1").Year - (#1/1/0100#).Year

    ''' <summary>
    ''' 日付を表す数値がEmptyかを返す
    ''' </summary>
    ''' <param name="yyyymmdd">日付を表す数値</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Private Shared Function IsEmpty(ByVal yyyymmdd As Nullable(Of Integer)) As Boolean
        Return Not yyyymmdd.HasValue OrElse yyyymmdd.Value = 0
    End Function

    ''' <summary>
    ''' YYYYMMDD書式からYYYY/MM/DD書式に変換する
    ''' </summary>
    ''' <param name="yyyymmdd">YYYYMMDD書式の文字列</param>
    ''' <returns>YYYY/MM/DD書式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvYyyymmddToSlashYyyymmdd(ByVal yyyymmdd As String) As String
        If StringUtil.IsEmpty(yyyymmdd) Then
            Return Nothing
        End If
        Return ConvYyyymmddToSlashYyyymmdd(CInt(yyyymmdd))
    End Function

    ''' <summary>
    ''' YYYYMMDD書式からYYYY/MM/DD書式に変換する
    ''' </summary>
    ''' <param name="yyyymmdd">YYYYMMDD書式の数値</param>
    ''' <returns>YYYY/MM/DD書式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvYyyymmddToSlashYyyymmdd(ByVal yyyymmdd As Nullable(Of Integer)) As String
        If IsEmpty(yyyymmdd) Then
            Return Nothing
        End If
        If yyyymmdd = MAX_VALUE_INTEGER Then
            Return MAX_VALUE_WITH_SLASH_AS_STRING
        End If
        Return Format(yyyymmdd, "0000/00/00")
    End Function

    ''' <summary>
    ''' HHMMSS書式からHH:MM:SS書式に変換する
    ''' </summary>
    ''' <param name="hhmmss">HHMMSS書式の文字列</param>
    ''' <returns>HH:MM:SS書式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvHhmmssToColonHhmmss(ByVal hhmmss As String) As String
        If StringUtil.IsEmpty(hhmmss) Then
            Return Nothing
        End If
        Return ConvHhmmssToColonHhmmss(CInt(hhmmss))
    End Function

    ''' <summary>
    ''' HHMMSS書式からHH:MM:SS書式に変換する
    ''' </summary>
    ''' <param name="hhmmss">HHMMSS書式の数値</param>
    ''' <returns>HH:MM:SS書式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvHhmmssToColonHhmmss(ByVal hhmmss As Integer?) As String
        If IsEmpty(hhmmss) Then
            Return Nothing
        End If
        Return Format(hhmmss, "00:00:00")
    End Function

    ''' <summary>
    ''' 時分秒を合計秒にして返す
    ''' </summary>
    ''' <param name="colonHhmmss">HH:MM:SS書式の文字列</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvTimeToTotalSeconds(ByVal colonHhmmss As String) As Integer
        Dim aDate As Date
        If Not Date.TryParse(colonHhmmss, aDate) Then
            Throw New ArgumentException("時刻を指定すべき")
        End If

        Return CInt(aDate.TimeOfDay.TotalSeconds)
    End Function

    ''' <summary>
    ''' YYYYMMDD書式から日付型に変換する
    ''' </summary>
    ''' <param name="yyyymmdd">YYYYMMDD書式の文字列</param>
    ''' <returns>日付型</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvYyyymmddToDate(ByVal yyyymmdd As String) As Nullable(Of Date)
        Return ConvYyyymmddToDate(StringUtil.ToInteger(yyyymmdd))
    End Function

    ''' <summary>
    ''' YYYYMMDD書式から日付型に変換する
    ''' </summary>
    ''' <param name="yyyymmdd">YYYYMMDD書式の数値</param>
    ''' <returns>日付値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvYyyymmddToDate(ByVal yyyymmdd As Nullable(Of Integer)) As Nullable(Of Date)
        If IsEmpty(yyyymmdd) Then
            Return Nothing
        End If
        If yyyymmdd = MAX_VALUE_INTEGER Then
            Return MAX_VALUE_DATETIME
        End If
        Return CCDate(ConvYyyymmddToSlashYyyymmdd(yyyymmdd))
    End Function

    ''' <summary>
    ''' YYYYMMDD書式から日付型に変換する
    ''' </summary>
    ''' <param name="yyyymmdd">YYYYMMDD書式の数値</param>
    ''' <param name="hhmmss">HHMMSS書式の数値</param>
    ''' <returns>日時値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvYyyymmddToDate(ByVal yyyymmdd As Integer?, ByVal hhmmss As Integer?) As DateTime?
        If Not yyyymmdd.HasValue AndAlso Not hhmmss.HasValue Then
            Return Nothing
        End If
        Dim colonHhmmss As String = ConvHhmmssToColonHhmmss(EzUtil.Nvl(hhmmss, 0))
        If Not yyyymmdd.HasValue Then
            Return CDate(colonHhmmss)
        End If
        If Not InternalIsDateVB(colonHhmmss) Then
            Return ConvYyyymmddToDate(yyyymmdd)
        End If
        Return New DateTime(ConvYyyymmddToDate(yyyymmdd).Value.Ticks + CDate(colonHhmmss).Ticks)
    End Function

    ''' <summary>
    ''' YYYY/MM/DD書式からYYYYMMDD書式に変換する
    ''' </summary>
    ''' <param name="slashYyyymmdd">YYYY/MM/DD書式の文字列</param>
    ''' <returns>YYYYMMDD書式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvSlashYyyymmddToYyyymmdd(ByVal slashYyyymmdd As String) As String
        Dim result As Integer? = ConvSlashYyyymmddToYyyymmddAsInteger(slashYyyymmdd)
        If Not result.HasValue Then
            Return Nothing
        End If
        Return result.Value.ToString
    End Function

    ''' <summary>
    ''' YYYY/MM/DD書式からYYYYMMDD書式に変換する
    ''' </summary>
    ''' <param name="slashYyyymmdd">YYYY/MM/DD書式の文字列</param>
    ''' <returns>YYYYMMDD書式の数値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvSlashYyyymmddToYyyymmddAsInteger(ByVal slashYyyymmdd As String) As Nullable(Of Integer)
        If StringUtil.IsEmpty(slashYyyymmdd) Then
            Return Nothing
        End If
        If MAX_VALUE_WITH_SLASH_AS_STRING.Equals(slashYyyymmdd) Then
            Return MAX_VALUE_INTEGER
        End If
        Dim dominical As Date = CCDate(slashYyyymmdd)
        Return dominical.Year * 10000 + dominical.Month * 100 + dominical.Day
    End Function

    ''' <summary>
    ''' 日時の年月日を YYYYMMDD 形式の数値で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>YYYYMMDD形式の数値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateToInteger(ByVal aDateTime As DateTime?) As Int32?
        If Not aDateTime.HasValue Then
            Return Nothing
        End If
        Dim effectiveDateTime As DateTime = aDateTime.Value
        If TruncTime(effectiveDateTime).CompareTo(MAX_VALUE_DATETIME) = 0 Then
            Return MAX_VALUE_INTEGER
        End If
        Return effectiveDateTime.Year * 10000 + effectiveDateTime.Month * 100 + effectiveDateTime.Day
    End Function

    ''' <summary>
    ''' 日時の年月日を YYYYMM 形式の文字列で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>YYYYMM形式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateToYyyymm(ByVal aDateTime As DateTime?) As String
        Dim result As Integer? = ConvDateToYyyymmAsInteger(aDateTime)
        If Not result.HasValue Then
            Return Nothing
        End If
        Return result.Value.ToString
    End Function

    ''' <summary>
    ''' 日時の年月日を YYYYMM 形式の数値で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>YYYYMM形式の数値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateToYyyymmAsInteger(ByVal aDateTime As DateTime?) As Int32?
        Dim result As Integer? = ConvDateToInteger(aDateTime)
        If Not result.HasValue Then
            Return Nothing
        End If
        Return CInt(Math.Truncate(result.Value / 100))
    End Function

    ''' <summary>
    ''' 日付を YYYYMMDD 形式の文字列で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>YYYYMMDD 形式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateToYyyymmdd(ByVal aDateTime As DateTime?) As String
        Dim result As Integer? = ConvDateToInteger(aDateTime)
        If Not result.HasValue Then
            Return Nothing
        End If
        Return result.Value.ToString
    End Function

    ''' <summary>
    ''' 日時の年月日を YYYYMMDD 形式の数値で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>YYYYMMDD形式の数値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateToYyyymmddAsInteger(ByVal aDateTime As DateTime?) As Integer?
        Return ConvDateToInteger(aDateTime)
    End Function

    ''' <summary>
    ''' 日付を YYYY/MM/DD 形式の文字列で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>YYYY/MM/DD 形式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateToSlashYyyymmdd(ByVal aDateTime As DateTime?) As String
        Return PerformConvDateToFormat(aDateTime, "yyyy/MM/dd", MAX_VALUE_WITH_SLASH_AS_STRING)
    End Function

    Private Shared Function PerformConvDateToFormat(ByVal aDateTime As Date?, ByVal dateFormat As String, ByVal valueWhenMax As String) As String
        If Not aDateTime.HasValue Then
            Return Nothing
        End If
        Dim effectiveDateTime As DateTime = aDateTime.Value
        If TruncTime(effectiveDateTime).CompareTo(MAX_VALUE_DATETIME) = 0 Then
            Return valueWhenMax
        End If
        Return aDateTime.Value.ToString(dateFormat, MyCultureInfo.YMD_TIME)
    End Function

    ''' <summary>
    ''' 日付を YYYY-MM-DD 形式の文字列で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>YYYY-MM-DD 形式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateToHyphenYyyymmdd(ByVal aDateTime As DateTime?) As String
        Return PerformConvDateToFormat(aDateTime, "yyyy-MM-dd", MAX_VALUE_WITH_HYPHEN_AS_STRING)
    End Function

    ''' <summary>
    ''' 日時の時分秒を HHMMSS 形式の文字列で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>HHMMSS形式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvTimeToHhmmss(ByVal aDateTime As DateTime?) As String
        If Not aDateTime.HasValue Then
            Return Nothing
        End If
        Return aDateTime.Value.ToString("HHmmss")
    End Function

    ''' <summary>
    ''' 日時の時分秒を HHMM 形式の文字列で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>HHMM形式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvTimeToHhmm(ByVal aDateTime As DateTime?) As String
        If Not aDateTime.HasValue Then
            Return Nothing
        End If
        Return aDateTime.Value.ToString("HHmm")
    End Function

    ''' <summary>
    ''' 日時の時分秒を HHMMSS 形式の数値で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>HHMMSS形式の数値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvTimeToInteger(ByVal aDateTime As DateTime?) As Int32?
        If Not aDateTime.HasValue Then
            Return Nothing
        End If
        Return Convert.ToInt32(ConvTimeToHhmmss(aDateTime))
    End Function

    ''' <summary>
    ''' 日時の時分秒を HH:MM:SS書式に変換する
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>HH:MM:SS書式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvTimeToColonHhmmss(ByVal aDateTime As DateTime?) As String
        If Not aDateTime.HasValue Then
            Return Nothing
        End If
        Return ConvHhmmssToColonHhmmss(ConvTimeToInteger(aDateTime))
    End Function

    ''' <summary>
    ''' 日時の時分秒を HH:MM書式に変換する
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>HH:MM書式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvTimeToColonHhmm(ByVal aDateTime As DateTime?) As String
        If Not aDateTime.HasValue Then
            Return Nothing
        End If
        Return aDateTime.Value.ToString("HH:mm")
    End Function

    ''' <summary>
    ''' 日付を表す値を DateTime にして返す
    ''' </summary>
    ''' <param name="dateValue">日付を表す値</param>
    ''' <returns>DateTime型</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateValueToDateTime(ByVal dateValue As Object) As DateTime?
        If TypeOf dateValue Is DateTime Then
            Return DirectCast(dateValue, DateTime)
        ElseIf TypeOf dateValue Is String Then
            Dim yyyymmdd As String = DirectCast(dateValue, String)
            If IsNumeric(dateValue) Then
                Return ConvYyyymmddToDate(yyyymmdd)
            ElseIf IsDate(dateValue) Then
                Return CCDate(dateValue)
            ElseIf MAX_VALUE_WITH_SLASH_AS_STRING.Equals(dateValue) Then
                Return MAX_VALUE_DATETIME
            End If
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 日付を表す値をYYYYMMDD形式の数値で返す
    ''' </summary>
    ''' <param name="dateValue">日付を表す値</param>
    ''' <returns>YYYYMMDD形式の数値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateValueToInteger(ByVal dateValue As Object) As Int32?
        Dim aDate As DateTime? = ConvDateValueToDateTime(dateValue)
        If Not aDate.HasValue Then
            Return Nothing
        End If
        Return ConvDateToInteger(aDate)
    End Function

    ''' <summary>
    ''' 日付を表す値を YYYY/MM/DD 形式の文字列で返す
    ''' </summary>
    ''' <param name="dateValue">日付を表す値</param>
    ''' <returns>YYYY/MM/DD 形式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateValueToSlashYyyymmdd(ByVal dateValue As Object) As String
        Dim aDate As DateTime? = ConvDateValueToDateTime(dateValue)
        If Not aDate.HasValue Then
            Return Nothing
        End If
        Return ConvDateToSlashYyyymmdd(aDate)
    End Function

    ''' <summary>
    ''' 日時を M月d日（曜日） HH:mm 形式の文字列で返す
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>YYYY-MM-DD 形式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDateTimeToMddowhhmm(ByVal aDateTime As DateTime?) As String
        If Not aDateTime.HasValue Then
            Return Nothing
        End If
        Return aDateTime.Value.ToString("M月d日（ddd） HH:mm")
    End Function

    ''' <summary>
    ''' 日時から、時分秒を切り捨てる
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>日付</returns>
    ''' <remarks></remarks>
    Public Shared Function TruncTime(ByVal aDateTime As Date) As Date
        Return aDateTime.Date
    End Function

    ''' <summary>
    ''' 日時から、ミリ秒以下を切り捨てる
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <returns>日時</returns>
    ''' <remarks></remarks>
    Public Shared Function TruncMilliseconds(ByVal aDateTime As Date) As Date
        Return aDateTime.Date.AddSeconds(Math.Truncate(aDateTime.TimeOfDay.TotalSeconds))
    End Function

    ''' <summary>
    ''' YYYYMMDD形式の数値をYYYYMM形式の数値にして返す
    ''' </summary>
    ''' <param name="yyyymmdd">YYYYMMDD形式の数値</param>
    ''' <returns>YYYYMM形式の数値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvYyyymmddToYyyymm(ByVal yyyymmdd As Integer?) As Integer?
        If IsEmpty(yyyymmdd) Then
            Return Nothing
        End If
        'Return ConvYyyymmddToYyyymm(yyyymmdd.Value)
        Dim yyyymmddStr As String = yyyymmdd.Value.ToString
        If yyyymmddStr.Length < 8 Then
            Throw New ArgumentException("年月日を指定すべき")
        End If
        Return CInt(yyyymmddStr.Substring(0, 6))
    End Function

    ''' <summary>
    ''' YYYYMM形式の数値をYYYY/MM形式の文字列にして返す
    ''' </summary>
    ''' <param name="yyyymm">YYYYMM形式の数値</param>
    ''' <returns>YYYY/MM形式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvYyyymmToSlashYyyymm(ByVal yyyymm As Integer?) As String
        If IsEmpty(yyyymm) Then
            Return Nothing
        End If
        Dim yyyymmStr As String = yyyymm.Value.ToString
        If yyyymmStr.Length < 6 Then
            Throw New ArgumentException("年月を指定すべき")
        End If
        Return Format(yyyymm, "0000/00")
    End Function

    ''' <summary>
    ''' YYYY/MM書式からYYYYMM書式に変換する
    ''' </summary>
    ''' <param name="slashYyyymmdd">YYYY/MM書式の文字列</param>
    ''' <returns>YYYYMM書式の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvSlashYyyymmToYyyymm(ByVal slashYyyymmdd As String) As String
        Dim result As Integer? = ConvSlashYyyymmToYyyymmAsInteger(slashYyyymmdd)
        If Not result.HasValue Then
            Return Nothing
        End If
        Return result.Value.ToString
    End Function

    ''' <summary>
    ''' YYYY/MM書式からYYYYMM書式に変換する
    ''' </summary>
    ''' <param name="slashYyyymm">YYYY/MM書式の文字列</param>
    ''' <returns>YYYYMM書式の数値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvSlashYyyymmToYyyymmAsInteger(ByVal slashYyyymm As String) As Nullable(Of Integer)
        If StringUtil.IsEmpty(slashYyyymm) Then
            Return Nothing
        End If
        Dim dominical As Date = CCDate(slashYyyymm & "/01")
        Return dominical.Year * 100 + dominical.Month
    End Function

    ''' <summary>
    ''' ミリ秒で表される現在の時間を返す
    ''' </summary>
    ''' <returns>1970/01/01 0:00:00 からのミリ秒</returns>
    ''' <remarks>java.lang.System.currentTimeMillis 相当</remarks>
    Public Shared Function CurrentTimeMillis() As Long

        Return ToTimeMillis(GetNowDateTime)
    End Function

    ''' <summary>
    ''' 1970/01/01 0:00:00 からのミリ秒にして返す
    ''' </summary>
    ''' <param name="aDate">日時</param>
    ''' <returns>ミリ秒</returns>
    ''' <remarks></remarks>
    Public Shared Function ToTimeMillis(ByVal aDate As DateTime) As Long

        Return Convert.ToInt64((aDate.Ticks - 621355968000000000&) / 10000)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後にSortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareNullsLast(ByVal x As Object, ByVal y As Object) As Integer
        Return CompareNullsLast(ConvDateValueToDateTime(x), ConvDateValueToDateTime(y))
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後に降順Sortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareDescNullsLast(ByVal x As Object, ByVal y As Object) As Integer
        Return CompareDescNullsLast(ConvDateValueToDateTime(x), ConvDateValueToDateTime(y))
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後にSortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareNullsLast(ByVal x As DateTime?, ByVal y As DateTime?) As Integer
        Dim result As Integer? = EzUtil.CompareObjectNullsLast(x, y)
        If result.HasValue Then
            Return result.Value
        End If
        Return x.Value.CompareTo(y.Value)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後に降順Sortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareDescNullsLast(ByVal x As DateTime?, ByVal y As DateTime?) As Integer
        Dim result As Integer? = EzUtil.CompareObjectNullsLast(x, y)
        If result.HasValue Then
            Return result.Value
        End If
        Return y.Value.CompareTo(x.Value)
    End Function

    ''' <summary>
    ''' 6桁年月日を8桁年月日にして返す
    ''' </summary>
    ''' <param name="yymmdd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvYymmddToYyyymmdd(ByVal yymmdd As Integer?) As Integer?
        If Not yymmdd.HasValue Then
            Return Nothing
        End If
        Return ConvYymmddToYyyymmdd(If(yymmdd.Value < 100000, StringUtil.Format("000000", yymmdd.Value), yymmdd.ToString))
    End Function
    ''' <summary>
    ''' 6桁年月日を8桁年月日にして返す
    ''' </summary>
    ''' <param name="yymmdd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvYymmddToYyyymmdd(ByVal yymmdd As String) As Integer?
        If yymmdd Is Nothing Then
            Return Nothing
        End If
        If yymmdd.Length <= 4 Then
            Throw New ArgumentException("年が不明:" & yymmdd)
        End If
        Return CInt(CorrectTo4DigitYear(yymmdd.Substring(0, yymmdd.Length - 4)) & yymmdd.Substring(yymmdd.Length - 4))
    End Function

    Private Class TimeZoneLocalHolder
        ''' <summary>UTCから現地時刻の差 … UTC + Offset = 現地時刻</summary>
        Public Shared ReadOnly OffsetFromUtc As Long = DateTime.Now.Ticks - DateTime.UtcNow.Ticks
    End Class

    ''' <summary>
    ''' 端末日時を取得する
    ''' </summary>
    ''' <returns>端末日時</returns>
    ''' <remarks></remarks>
    Public Shared Function GetNowDateTime() As DateTime
        ' .NET標準の DateTime.Now は、内部で 
        ' DirectCast(TimeZone.CurrentTimeZone, CurrentSystemTimeZone).GetUtcOffsetFromUniversalTime(time, False)
        ' しているが、これが結構遅い
        ' その処理を端折って、下記の通り取得すると 328(ms)→30(ms)
        Return New DateTime(DateTime.UtcNow.Ticks + TimeZoneLocalHolder.OffsetFromUtc, DateTimeKind.Local)
    End Function

    ''' <summary>
    ''' 年を加算した日時を取得する
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <param name="addendYear">加算年</param>
    ''' <returns>年を加算した日時</returns>
    ''' <remarks></remarks>
    Public Shared Function AddYears(ByVal aDateTime As Date?, ByVal addendYear As Integer) As Date?
        If Not aDateTime.HasValue Then
            Return aDateTime
        End If
        Return aDateTime.Value.AddYears(addendYear)
    End Function

    ''' <summary>
    ''' 日を加算した日時を取得する
    ''' </summary>
    ''' <param name="aDateTime">日時</param>
    ''' <param name="addendDay">加算日</param>
    ''' <returns>日を加算した日時</returns>
    ''' <remarks></remarks>
    Public Shared Function AddDays(ByVal aDateTime As Date?, ByVal addendDay As Integer) As Date?
        If Not aDateTime.HasValue Then
            Return aDateTime
        End If
        Return aDateTime.Value.AddDays(addendDay)
    End Function

    ''' <summary>
    ''' YYYYMM値をDate値にする（日にちは1日固定）
    ''' </summary>
    ''' <param name="yyyymm">年月(YYYYMM)</param>
    ''' <param name="systemYyyymmIfInvalid">yyyymmが不正のとき、システム年月を返す場合、true</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvYyyymmToDate(ByVal yyyymm As String, Optional ByVal systemYyyymmIfInvalid As Boolean = False) As Date?
        Dim yyyymm2 As String = yyyymm
        If EzUtil.IsEmpty(yyyymm2) OrElse yyyymm2.Length <> 6 OrElse Not IsDate(DateUtil.ConvYyyymmddToSlashYyyymmdd(yyyymm2 & "01")) Then
            If Not systemYyyymmIfInvalid Then
                Return Nothing
            End If
            yyyymm2 = DateUtil.ConvDateToYyyymm(GetNowDateTime)
        End If
        Return DateUtil.ConvYyyymmddToDate(yyyymm2 & "01")
    End Function

    ''' <summary>
    ''' (直前の/最後の)半期末にする
    ''' </summary>
    ''' <param name="aDate">対象日付</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function MakeEndOfLastBusinessPeriod(ByVal aDate As DateTime?) As DateTime?
        If Not aDate.HasValue Then
            Return Nothing
        End If
        Dim yearForCDate As Integer = aDate.Value.Year
        If aDate.Value.Month <= 3 Then
            Return New Date(yearForCDate - 1, 9, 30)
        ElseIf aDate.Value.Month <= 9 Then
            Return New Date(yearForCDate, 3, 31)
        Else ' aDate.Value.Month <= 12
            Return New Date(yearForCDate, 9, 30)
        End If
    End Function

    ''' <summary>
    ''' 指定日付の当期の開始日付を取得する
    ''' </summary>
    ''' <param name="yyyymmdd">指定日付</param>
    ''' <returns>期の開始日付(整数型日付)</returns>
    ''' <remarks></remarks>
    Public Shared Function MakeStartOfCurrentBusinessPeriod(yyyymmdd As Integer?) As Integer?
        If yyyymmdd Is Nothing Then
            Return Nothing
        End If

        Dim baseDate As DateTime = DateUtil.ConvYyyymmddToDate(yyyymmdd).Value
        Dim fiscalYear As Integer = DateUtil.DetectBusinessYear(baseDate)
        Dim halfPeriod As BusinessPeriod = DateUtil.DetectBusinessPeriod(baseDate)

        Dim resultDate As DateTime
        If BusinessPeriod.FIRST_PERIOD = halfPeriod Then
            resultDate = New Date(fiscalYear, 4, 1)
        Else
            resultDate = New Date(fiscalYear, 10, 1)
        End If
        Return DateUtil.ConvDateToInteger(resultDate)
    End Function

    ''' <summary>
    ''' 日付を半期末に切り捨てる
    ''' </summary>
    ''' <param name="aDate">対象日付</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function TruncateEndOfBusinessPeriod(ByVal aDate As DateTime?) As DateTime?
        If Not aDate.HasValue Then
            Return Nothing
        End If
        Dim yearForCDate As Integer = aDate.Value.Year
        If aDate.Value.Month * 100 + aDate.Value.Day < 331 Then
            Return New Date(yearForCDate - 1, 9, 30)
        ElseIf aDate.Value.Month * 100 + aDate.Value.Day < 930 Then
            Return New Date(yearForCDate, 3, 31)
        Else ' aDate.Value.Month <= 12
            Return New Date(yearForCDate, 9, 30)
        End If
    End Function

    ''' <summary>
    ''' ローカルPC コントロールパネル|地域と言語 の日付書式に左右されない日時文字列にする
    ''' </summary>
    ''' <param name="aDate">日時</param>
    ''' <returns>yyyy/mm/dd h24:mm:ss</returns>
    ''' <remarks>日時値`2010/01/02 03:04:05`をToStringすると、XP="2010/01/02 3:04:05" Win7="10/01/02 3:04:05"</remarks>
    Public Shared Function ToFixedString(ByVal aDate As DateTime?) As String
        Return StringUtil.ToDateTimeString(aDate)
    End Function

    ''' <summary>
    ''' システムが和暦設定か？を返す
    ''' </summary>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Friend Shared Function IsJapaneseCalendarOnSystem() As Boolean
        Return 0 < YEAR_DIFF_OF_JAPANESE_CALENDAR
    End Function

    Private Shared Function InternalIsDateVB(ymdTime As Object) As Boolean
        Return Microsoft.VisualBasic.IsDate(ymdTime)
    End Function

    ''' <summary>
    ''' 常に年を西暦として、年月日の並びはYMD固定として日付型を作成する
    ''' </summary>
    ''' <param name="ymdTime">年月日の並びはYMD固定とした西暦の日時文字列</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' システム設定が和暦だと以下が問題<br/>
    ''' 1) CDate("2000/01/01") は平成2000年1月1日と解釈される<br/>
    ''' 2) (#2000/1/1#).ToString("yyyyMMdd") は 120101 になる(12は平成12年の意味)<br/>
    ''' システム設定が西暦だと以下が問題<br/>
    ''' 1) 西暦が2桁だとシステム設定によるが1930-2030等に解釈される
    ''' </remarks>
    Public Shared Function CCDate(ymdTime As Object) As DateTime
        Dim ymdTimeString As String = StringUtil.ToString(ymdTime)
        If Not IsDate(ymdTime) Then
            ' CDateでエラーにする
            Return CDate(ymdTime)
        ElseIf Not ymdTimeString.Contains("/") AndAlso Not ymdTimeString.Contains("-") Then
            ' 年月日がないならCDateで変換
            Return CDate(ymdTime)
        End If
        Return Date.Parse(ymdTimeString, MyCultureInfo.YMD_TIME)
    End Function

    ''' <summary>
    ''' 常に年を西暦として、年月日の並びはYMD固定として日時かどうか判定する
    ''' </summary>
    ''' <param name="ymdTime">年月日の並びはYMD固定とした西暦の日時文字列</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsDate(ymdTime As Object) As Boolean
        If InternalIsDateVB(ymdTime) Then
            Return True
        End If
        ' YMD文字列の引数に対し、OS日付書式MDYの場合を想定して以下で救済
        Dim dateParts As String() = Nothing
        Dim timeParts As String() = Nothing
        Dim ymdTimeString As String = StringUtil.ToString(ymdTime)
        For Each aPart As String In Split(ymdTimeString, " ")
            If aPart.Contains("/") Then
                dateParts = Split(aPart, "/")
            ElseIf aPart.Contains("-") Then
                dateParts = Split(aPart, "-")
            ElseIf aPart.Contains(":") Then
                timeParts = Split(aPart, ":")
            ElseIf IsNumeric(aPart) Then
                Return False
            End If
        Next
        If dateParts Is Nothing AndAlso timeParts Is Nothing Then
            Return False
        End If
        If dateParts IsNot Nothing Then
            If Not dateParts.All(Function(aPart) IsNumeric(aPart)) Then
                Return False
            End If
            If 3 < dateParts.Length Then
                Return False
            ElseIf dateParts.Length = 3 Then
                If Not InternalIsDateVB(dateParts(1) & "/" & dateParts(2)) Then
                    Return False
                End If
            ElseIf dateParts.Length = 2 Then
                If Not InternalIsDateVB(dateParts(0) & "/" & dateParts(1)) Then
                    Return False
                End If
            End If
        End If
        If timeParts IsNot Nothing Then
            If dateParts Is Nothing Then
                Return InternalIsDateVB(ymdTimeString)
            End If
            If dateParts.Length < 3 Then
                Return False
            End If
            If Not InternalIsDateVB(Join(timeParts, ":")) Then
                Return False
            End If
        End If
        Return True
    End Function

    ''' <summary>
    ''' 年度を特定する
    ''' </summary>
    ''' <param name="aDate">判定したい日時</param>
    ''' <returns>年度</returns>
    ''' <remarks></remarks>
    Public Shared Function DetectBusinessYear(ByVal aDate As DateTime) As Integer
        Return If(aDate.Month < 4, aDate.Year - 1, aDate.Year)
    End Function

    ''' <summary>
    ''' 期を特定する
    ''' </summary>
    ''' <param name="aDate">判定したい日時</param>
    ''' <returns>今期なら1、来期なら2</returns>
    ''' <remarks></remarks>
    Public Shared Function DetectBusinessPeriod(ByVal aDate As DateTime) As BusinessPeriod
        If 3 < aDate.Month AndAlso aDate.Month < 10 Then
            Return BusinessPeriod.FIRST_PERIOD
        End If
        Return BusinessPeriod.SECOND_PERIOD
    End Function

    ''' <summary>
    ''' 今期かどうか
    ''' </summary>
    ''' <param name="aDate">判定したい日時</param>
    ''' <returns>今期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsThisBusinessPeriod(ByVal aDate As DateTime) As Boolean
        Return IsThisBusinessPeriod(aDate, GetNowDateTime())
    End Function

    ''' <summary>
    ''' 今期かどうか
    ''' </summary>
    ''' <param name="aDate">判定したい日時</param>
    ''' <param name="now">現在日時</param>
    ''' <returns>今期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsThisBusinessPeriod(ByVal aDate As DateTime, ByVal now As DateTime) As Boolean
        Return IsThisBusinessPeriod(DetectBusinessYear(aDate), DetectBusinessPeriod(aDate), now)
    End Function

    ''' <summary>
    ''' 今期かどうか
    ''' </summary>
    ''' <param name="businessYear">年度</param>
    ''' <param name="aBusinessPeriod">期</param>
    ''' <returns>今期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsThisBusinessPeriod(ByVal businessYear As Integer, ByVal aBusinessPeriod As BusinessPeriod) As Boolean
        Return IsThisBusinessPeriod(businessYear, aBusinessPeriod, GetNowDateTime())
    End Function

    ''' <summary>
    ''' 今期かどうか
    ''' </summary>
    ''' <param name="businessYear">年度</param>
    ''' <param name="aBusinessPeriod">期</param>
    ''' <param name="now">現在日時</param>
    ''' <returns>今期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsThisBusinessPeriod(ByVal businessYear As Integer, ByVal aBusinessPeriod As BusinessPeriod, ByVal now As DateTime) As Boolean
        Dim nowYear As Integer = DetectBusinessYear(now)
        Dim nowPeriod As BusinessPeriod = DetectBusinessPeriod(now)
        Return (nowYear = businessYear AndAlso nowPeriod.Equals(aBusinessPeriod))
    End Function

    ''' <summary>
    ''' 前期かどうか
    ''' </summary>
    ''' <param name="aDate">判定したい日時</param>
    ''' <returns>前期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsPrevBusinessPeriod(ByVal aDate As DateTime) As Boolean
        Return IsPrevBusinessPeriod(aDate, GetNowDateTime())
    End Function

    ''' <summary>
    ''' 前期かどうか
    ''' </summary>
    ''' <param name="aDate">判定したい日時</param>
    ''' <param name="now">現在日時</param>
    ''' <returns>前期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsPrevBusinessPeriod(ByVal aDate As DateTime, ByVal now As DateTime) As Boolean
        Return IsPrevBusinessPeriod(DetectBusinessYear(aDate), DetectBusinessPeriod(aDate), now)
    End Function

    ''' <summary>
    ''' 前期かどうか
    ''' </summary>
    ''' <param name="businessYear">年度</param>
    ''' <param name="aBusinessPeriod">期</param>
    ''' <returns>前期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsPrevBusinessPeriod(ByVal businessYear As Integer, ByVal aBusinessPeriod As BusinessPeriod) As Boolean
        Return IsPrevBusinessPeriod(businessYear, aBusinessPeriod, GetNowDateTime())
    End Function

    ''' <summary>
    ''' 前期かどうか
    ''' </summary>
    ''' <param name="businessYear">年度</param>
    ''' <param name="aBusinessPeriod">期</param>
    ''' <param name="now">現在日時</param>
    ''' <returns>前期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsPrevBusinessPeriod(ByVal businessYear As Integer, ByVal aBusinessPeriod As BusinessPeriod, ByVal now As DateTime) As Boolean
        Dim nowYear As Integer = DetectBusinessYear(now)
        Dim nowPeriod As BusinessPeriod = DetectBusinessPeriod(now)
        If nowPeriod = 1 Then
            Return (businessYear = nowYear - 1 AndAlso aBusinessPeriod.Equals(BusinessPeriod.SECOND_PERIOD))
        End If
        Return (businessYear = nowYear AndAlso aBusinessPeriod.Equals(BusinessPeriod.FIRST_PERIOD))
    End Function

    ''' <summary>
    ''' 来期かどうか
    ''' </summary>
    ''' <param name="aDate">判定したい日時</param>
    ''' <returns>来期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNextBusinessPeriod(ByVal aDate As DateTime) As Boolean
        Return IsNextBusinessPeriod(aDate, GetNowDateTime())
    End Function

    ''' <summary>
    ''' 来期かどうか
    ''' </summary>
    ''' <param name="aDate">判定したい日時</param>
    ''' <param name="now">現在日時</param>
    ''' <returns>来期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNextBusinessPeriod(ByVal aDate As DateTime, ByVal now As DateTime) As Boolean
        Return IsNextBusinessPeriod(DetectBusinessYear(aDate), DetectBusinessPeriod(aDate), now)
    End Function

    ''' <summary>
    ''' 来期かどうか
    ''' </summary>
    ''' <param name="businessYear">年度</param>
    ''' <param name="aBusinessPeriod">期</param>
    ''' <returns>来期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNextBusinessPeriod(ByVal businessYear As Integer, ByVal aBusinessPeriod As BusinessPeriod) As Boolean
        Return IsNextBusinessPeriod(businessYear, aBusinessPeriod, GetNowDateTime())
    End Function

    ''' <summary>
    ''' 来期かどうか
    ''' </summary>
    ''' <param name="businessYear">年度</param>
    ''' <param name="aBusinessPeriod">期</param>
    ''' <param name="now">現在日時</param>
    ''' <returns>来期ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNextBusinessPeriod(ByVal businessYear As Integer, ByVal aBusinessPeriod As BusinessPeriod, ByVal now As DateTime) As Boolean
        Dim nowYear As Integer = DetectBusinessYear(now)
        Dim nowPeriod As Integer = DetectBusinessPeriod(now)
        If nowPeriod = 2 Then
            Return (businessYear = nowYear + 1 AndAlso aBusinessPeriod.Equals(BusinessPeriod.FIRST_PERIOD))
        End If
        Return (businessYear = nowYear AndAlso aBusinessPeriod.Equals(BusinessPeriod.SECOND_PERIOD))
    End Function

    ''' <summary>
    ''' 2桁年であれば4桁年に補正する
    ''' </summary>
    ''' <param name="year"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CorrectTo4DigitYear(year As String) As Integer
        If Not IsNumeric(year) Then
            Throw New ArgumentException("yearを解釈できない " & year, "year")
        End If
        Dim twoDigitYearMax As Integer = (New System.Globalization.GregorianCalendar).TwoDigitYearMax
        Dim threshold = twoDigitYearMax Mod 100
        Dim currentCentury = twoDigitYearMax - threshold
        Dim lastCentury = currentCentury - 100
        Dim y As Integer = CInt(year)
        Return If(year.Length <= 2, If(y <= threshold, currentCentury, lastCentury), 0) + y
    End Function

    ''' <summary>
    ''' 日時からミリ秒を四捨五入する
    ''' </summary>
    ''' <param name="aDateTime"></param>
    ''' <returns></returns>
    Public Shared Function RoundMillSeconds(aDateTime As DateTime) As DateTime
        Return TruncMilliseconds(aDateTime).AddSeconds(Math.Round(aDateTime.Millisecond / 1000, MidpointRounding.AwayFromZero))
    End Function

End Class
