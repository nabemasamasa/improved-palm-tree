Imports Fhi.Fw.Domain
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
''' 文字列操作のユーティリティ
''' </summary>
''' <remarks></remarks>
Public Class StringUtil
    ''' <summary>
    ''' キャメル記法を"_"記法して返す
    ''' </summary>
    ''' <param name="str">キャメル記法の文字列</param>
    ''' <returns>"_"記法の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function Decamelize(ByVal str As String) As String
        If str = String.Empty Then
            Return str
        End If
        Dim chars() As Char = str.ToCharArray
        Dim result As New StringBuilder
        result.Append(chars(0))
        For i As Integer = 1 To chars.Length - 1
            If "A"c <= chars(i) AndAlso chars(i) <= "Z"c Then
                result.Append("_"c)
            ElseIf "0"c <= chars(i) AndAlso chars(i) <= "9"c Then
                Dim numberStr As String = chars(i)
                For j As Integer = i + 1 To chars.Length - 1
                    If Not ("0"c <= chars(j) AndAlso chars(j) <= "9"c) Then
                        Exit For
                    End If
                    numberStr &= chars(j)
                    i = j
                Next
                result.Append("_"c).Append(numberStr)
                Continue For
            End If
            result.Append(chars(i))
        Next
        Return result.ToString.ToUpper
    End Function

    ''' <summary>
    ''' キャメル記法を"_"記法して返す(数字文字列はアンダーバーで区切らない)
    ''' </summary>
    ''' <param name="str">キャメル記法の文字列</param>
    ''' <returns>"_"記法の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function DecamelizeIgnoreNumber(ByVal str As String) As String
        If str = String.Empty Then
            Return str
        End If
        Dim chars() As Char = str.ToCharArray
        Dim result As New StringBuilder
        result.Append(chars(0))
        For i As Integer = 1 To chars.Length - 1
            If "A"c <= chars(i) AndAlso chars(i) <= "Z"c Then
                result.Append("_"c)
            End If
            result.Append(chars(i))
        Next
        Return result.ToString.ToUpper
    End Function

    ''' <summary>
    ''' Empty値、またはNull値か、を返す
    ''' </summary>
    ''' <param name="str">判定する文字列</param>
    ''' <returns>Empty値かNull値の場合、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmpty(ByVal str As Object) As Boolean
        Return str Is Nothing OrElse If(TypeOf str Is PrimitiveValueObject,
                                        IsEmpty(DirectCast(str, PrimitiveValueObject).Value),
                                        str.ToString.Trim.Length = 0)
    End Function

    ''' <summary>
    ''' Empty値以外でかつ、Null値以外か、を返す
    ''' </summary>
    ''' <param name="str">判定する文字列</param>
    ''' <returns>Empty値以外でかつ、Null値以外の場合、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEmpty(ByVal str As Object) As Boolean
        Return Not IsEmpty(str)
    End Function

    ''' <summary>
    ''' 数値をInteger型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <returns>Integer型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToInteger(ByVal number As Object) As Integer?
        Return ToInteger(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値をInteger型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Integer型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToInteger(ByVal number As Object, ByVal valueIfNull As Integer?) As Integer?
        Return ToInteger(StringUtil.ToString(number), valueIfNull)
    End Function

    ''' <summary>
    ''' 数値文字列をInteger型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <returns>Integer型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToInteger(ByVal number As String) As Integer?
        Return ToInteger(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値文字列をInteger型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Integer型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToInteger(ByVal number As String, ByVal valueIfNull As Integer?) As Integer?
        If Not IsNumeric(number) Then
            Return valueIfNull
        End If
        Return CInt(number)
    End Function

    ''' <summary>
    ''' 値をDatetime型にして返す
    ''' </summary>
    ''' <param name="dateValue">値</param>
    ''' <returns>Datetime型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDatetime(ByVal dateValue As Object) As DateTime?
        Return ToDatetime(dateValue, Nothing)
    End Function

    ''' <summary>
    ''' 値をDatetime型にして返す
    ''' </summary>
    ''' <param name="dateValue">値</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Datetime型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDatetime(ByVal dateValue As Object, ByVal valueIfNull As DateTime?) As DateTime?
        Return ToDatetime(StringUtil.ToString(dateValue), valueIfNull)
    End Function

    ''' <summary>
    ''' 値文字列をDatetime型にして返す
    ''' </summary>
    ''' <param name="dateValue">値文字列</param>
    ''' <returns>Datetime型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDatetime(ByVal dateValue As String) As DateTime?
        Return ToDatetime(dateValue, Nothing)
    End Function

    ''' <summary>
    ''' 値文字列をDatetime型にして返す
    ''' </summary>
    ''' <param name="dateValue">値文字列</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Datetime型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDatetime(ByVal dateValue As String, ByVal valueIfNull As DateTime?) As DateTime?
        If Not IsDate(dateValue) Then
            Return valueIfNull
        End If
        Return CDate(dateValue)
    End Function

    ''' <summary>
    ''' 数値をDecimal型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <returns>Decimal型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDecimal(ByVal number As Object) As Decimal?
        Return ToDecimal(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値をDecimal型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Decimal型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDecimal(ByVal number As Object, ByVal valueIfNull As Decimal?) As Decimal?
        Return ToDecimal(StringUtil.ToString(number), valueIfNull)
    End Function

    ''' <summary>
    ''' 数値文字列をDecimal型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <returns>Decimal型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDecimal(ByVal number As String) As Decimal?
        Return ToDecimal(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値文字列をDecimal型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Decimal型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDecimal(ByVal number As String, ByVal valueIfNull As Decimal?) As Decimal?
        If Not IsNumeric(number) Then
            Return valueIfNull
        End If
        Return CDec(number)
    End Function

    ''' <summary>
    ''' 数値をDouble型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <returns>Double型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDouble(ByVal number As Object) As Double?
        Return ToDouble(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値をDouble型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Double型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDouble(ByVal number As Object, ByVal valueIfNull As Double?) As Double?
        Return ToDouble(StringUtil.ToString(number), valueIfNull)
    End Function

    ''' <summary>
    ''' 数値文字列をDouble型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <returns>Double型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDouble(ByVal number As String) As Double?
        Return ToDouble(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値文字列をDouble型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Double型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToDouble(ByVal number As String, ByVal valueIfNull As Double?) As Double?
        If Not IsNumeric(number) Then
            Return valueIfNull
        End If
        Return CDbl(number)
    End Function

    ''' <summary>
    ''' 数値をLong型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <returns>Long型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToLong(ByVal number As Object) As Long?
        Return ToLong(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値をLong型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Long型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToLong(ByVal number As Object, ByVal valueIfNull As Long?) As Long?
        Return ToLong(StringUtil.ToString(number), valueIfNull)
    End Function

    ''' <summary>
    ''' 数値文字列をLong型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <returns>Long型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToLong(ByVal number As String) As Long?
        Return ToLong(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値文字列をLong型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Long型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToLong(ByVal number As String, ByVal valueIfNull As Long?) As Long?
        If Not IsNumeric(number) Then
            Return valueIfNull
        End If
        Return CLng(number)
    End Function

    ''' <summary>
    ''' 数値をShort型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <returns>Short型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToShort(ByVal number As Object) As Short?
        Return ToShort(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値をShort型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Short型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToShort(ByVal number As Object, ByVal valueIfNull As Short?) As Short?
        Return ToShort(StringUtil.ToString(number), valueIfNull)
    End Function

    ''' <summary>
    ''' 数値文字列をShort型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <returns>Short型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToShort(ByVal number As String) As Short?
        Return ToShort(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値文字列をShort型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Short型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToShort(ByVal number As String, ByVal valueIfNull As Short?) As Short?
        If Not IsNumeric(number) Then
            Return valueIfNull
        End If
        Return CShort(number)
    End Function

    ''' <summary>
    ''' 数値をByte型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <returns>Byte型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToByte(ByVal number As Object) As Byte?
        Return ToByte(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値をByte型にして返す
    ''' </summary>
    ''' <param name="number">数値</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Byte型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToByte(ByVal number As Object, ByVal valueIfNull As Byte?) As Byte?
        Return ToByte(StringUtil.ToString(number), valueIfNull)
    End Function

    ''' <summary>
    ''' 数値文字列をByte型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <returns>Byte型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToByte(ByVal number As String) As Byte?
        Return ToByte(number, Nothing)
    End Function

    ''' <summary>
    ''' 数値文字列をByte型にして返す
    ''' </summary>
    ''' <param name="number">数値文字列</param>
    ''' <param name="valueIfNull">null置き換え値</param>
    ''' <returns>Byte型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToByte(ByVal number As String, ByVal valueIfNull As Byte?) As Byte?
        If Not IsNumeric(number) Then
            Return valueIfNull
        End If
        Return CByte(number)
    End Function

    ''' <summary>
    ''' Objectを文字列にして返す
    ''' </summary>
    ''' <param name="obj">数値</param>
    ''' <returns>文字列</returns>
    ''' <remarks>NullならNullのまま返す。Convert.ToString() はNullの時、0 が返る</remarks>
    Public Overloads Shared Function ToString(ByVal obj As Object) As String
        If obj Is Nothing Then
            Return Nothing
        End If
        If TypeOf obj Is DateTime OrElse TypeOf obj Is DateTime? Then
            Return ToDateTimeString(DirectCast(obj, DateTime?))
        End If
        Return obj.ToString
    End Function

    ''' <summary>
    ''' ローカルPC コントロールパネル|地域と言語 の日付書式に左右されない日時文字列にする
    ''' </summary>
    ''' <param name="aDate">日時</param>
    ''' <returns>yyyy/mm/dd h24:mm:ss</returns>
    ''' <remarks>日時値`2010/01/02 03:04:05`をToStringすると、XP="2010/01/02 3:04:05" Win7="10/01/02 3:04:05"</remarks>
    Public Shared Function ToDateTimeString(ByVal aDate As DateTime?) As String
        If Not aDate.HasValue Then
            Return Nothing
        End If
        Dim dateValue As Date = aDate.Value
        If dateValue.Year <= 1 Then
            Return dateValue.ToString("H:mm:ss")
        End If
        Return dateValue.ToString("yyyy/MM/dd H:mm:ss")
    End Function

    ''' <summary>
    ''' NULL値だったら空文字変換 Null Value Logic
    ''' </summary>
    ''' <param name="obj">判定object</param>
    ''' <returns>対応した文字列</returns>
    ''' <remarks>PrimitiveValueObject型を返したい場合は、`EzUtil#Nvl`を利用すべき</remarks>
    Public Shared Function Nvl(ByVal obj As Object) As String
        Return Nvl(obj, String.Empty)
    End Function

    ''' <summary>
    ''' NULL値だったら変換 Null Value Logic
    ''' </summary>
    ''' <param name="obj">判定object</param>
    ''' <param name="nullVal">置き換え文字列</param>
    ''' <returns>対応した文字列</returns>
    ''' <remarks>PrimitiveValueObject型を返したい場合は、`EzUtil#Nvl`を利用すべき</remarks>
    Public Shared Function Nvl(ByVal obj As Object, ByVal nullVal As String) As String
        If obj Is Nothing Then
            Return nullVal
        End If
        If TypeOf obj Is PrimitiveValueObject _
           AndAlso DirectCast(obj, PrimitiveValueObject).Value Is Nothing Then
            Return nullVal
        End If
        Return ToString(obj)
    End Function

    ''' <summary>
    ''' NULL値or空文字だったら、"0"に変換  Empty Value Logic
    ''' </summary>
    ''' <param name="obj">判定object</param>
    ''' <returns>対応した文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function EvlZero(ByVal obj As Object) As String
        Return Evl(obj, "0")
    End Function

    ''' <summary>
    ''' NULL値or空文字だったら変換  Empty Value Logic
    ''' </summary>
    ''' <param name="str">判定文字列</param>
    ''' <param name="emptyVal">置き換え文字列</param>
    ''' <returns>対応した文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function Evl(ByVal str As Object, ByVal emptyVal As String) As String
        If IsEmpty(str) Then
            Return emptyVal
        End If
        Return ToString(str)
    End Function

    ''' <summary>
    ''' 保持している値がNULL値or空文字の値オブジェクトだったら変換  Empty Value Logic
    ''' </summary>
    ''' <param name="value">判定値オブジェクト</param>
    ''' <param name="emptyValue">置き換え値オブジェクト</param>
    ''' <returns>対応した値オブジェクト</returns>
    ''' <remarks></remarks>
    Public Shared Function Evl(Of T As PrimitiveValueObject)(ByVal value As T, ByVal emptyValue As T) As T
        If IsEmpty(value) Then
            Return emptyValue
        End If
        Return value
    End Function

    ''' <summary>
    ''' 文字列を List(Of String) にして返す
    ''' </summary>
    ''' <param name="elements">Listの要素たち</param>
    ''' <returns>List(Of String)</returns>
    ''' <remarks></remarks>
    Public Shared Function ToList(ByVal ParamArray elements() As String) As List(Of String)
        Return elements.ToList()
    End Function

    ''' <summary>
    ''' 値を整形する
    ''' </summary>
    ''' <param name="formatString">整形書式</param>
    ''' <param name="value">値</param>
    ''' <returns>整形した値</returns>
    ''' <remarks></remarks>
    Public Shared Function Format(ByVal formatString As String, ByVal value As Object) As String
        Dim sb As New StringBuilder
        Return String.Format(sb.Append("{0:").Append(formatString).Append("}").ToString, value)
    End Function

    ''' <summary>
    ''' 値を固定長に整形する
    ''' </summary>
    ''' <param name="value">値</param>
    ''' <param name="length">固定長桁数</param>
    ''' <param name="alignsRight">右寄せにする場合、true</param>
    ''' <returns>整形した値</returns>
    ''' <remarks></remarks>
    Public Shared Function MakeFixedString(ByVal value As String, ByVal length As Integer, Optional ByVal alignsRight As Boolean = False) As String
        Const SPACE As Char = " "c
        Dim sb As New StringBuilder(value)
        Dim lengthByte As Integer = GetLengthByte(If(value, ""))
        If lengthByte < length Then
            If alignsRight Then
                sb.Insert(0, SPACE, length - lengthByte)
            Else
                sb.Append(SPACE, length - lengthByte)
            End If
        End If
        Return sb.ToString
    End Function

    ''' <summary>
    ''' 引用符に囲まれていたら取り外して返す
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns>引用符が取り除かれた文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function RemoveIfQuoted(ByVal str As String) As String
        If str Is Nothing Then
            Return str
        End If
        If (str.StartsWith("""") AndAlso str.EndsWith("""")) _
                OrElse (str.StartsWith("'") AndAlso str.EndsWith("'")) Then
            Return str.Substring(1, str.Length - 2)
        End If
        Return str
    End Function

    ''' <summary>
    ''' 文字を繰り返して返す
    ''' </summary>
    ''' <param name="c">文字</param>
    ''' <param name="count">繰り返し数</param>
    ''' <returns>繰り返した文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function Repeat(ByVal c As Char, ByVal count As Integer) As String
        Dim sb As New StringBuilder
        sb.Append(c, count)
        Return sb.ToString
    End Function

    ''' <summary>
    ''' 文字バイト数を取得
    ''' </summary>
    ''' <param name="val">文字列</param>
    ''' <returns>バイト数</returns>
    ''' <remarks></remarks>
    Public Shared Function GetLengthByte(ByVal val As String) As Integer
        Return Encoding.GetEncoding("SHIFT_JIS").GetByteCount(val)
    End Function

    ''' <summary>
    ''' 全角文字のみの文字列かを返す
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsZenkakuOnly(ByVal str As String) As Boolean
        Return GetLengthByte(str) = str.Length * 2
    End Function

    ''' <summary>
    ''' 半角文字のみの文字列かを返す
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsHankakuOnly(ByVal str As String) As Boolean
        Return GetLengthByte(str) = str.Length
    End Function

    ''' <summary>
    ''' （ファイル選択ダイアログ等で使用する）フィルターに一致するファイル名かを返す
    ''' </summary>
    ''' <param name="fileName">ファイル名</param>
    ''' <param name="filterPatterns">フィルター[]  ex) *.jpg</param>
    ''' <returns>判定結果</returns>
    ''' <remarks>*.* は全ての文字列と一致する</remarks>
    Public Shared Function IsMatchFilterPattern(ByVal fileName As String, ByVal ParamArray filterPatterns As String()) As Boolean

        For Each filterPattern As String In filterPatterns
            If "*.*".Equals(filterPattern) Then
                Return True
            End If
            If Regex.IsMatch(fileName, String.Format("^{0}$", Regex.Escape(filterPattern).Replace("\*", ".*")), RegexOptions.IgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' ルールに従ってインクリメントした文字列を返す
    ''' </summary>
    ''' <param name="src">対象の文字列</param>
    ''' <param name="ordinalRule">ルール</param>
    ''' <param name="place">対象の桁</param>
    ''' <param name="isFixed">固定長の場合、true</param>
    ''' <returns>インクリメント後の文字列</returns>
    ''' <remarks></remarks>
    Private Shared Function IncrementStr(ByVal src As String, ByVal ordinalRule As String, ByVal place As Integer, ByVal isFixed As Boolean) As String
        Dim charPlace As Integer = src.Length - 1 - place
        Dim sb As New StringBuilder
        If charPlace < 0 Then
            If isFixed Then
                ' 固定長指定の場合、そのまま返す
                Return src
            Else
                ' ordinalRule.Chars(0)は10進数でいう所の"0"なので桁上がりの値は .Chars(1) となる
                Return sb.Append(ordinalRule.Chars(If(1 < ordinalRule.Length, 1, 0))).Append(src).ToString
            End If
        End If
        For i As Integer = 0 To ordinalRule.Length - 1
            If src.Chars(charPlace) = ordinalRule.Chars(i) Then
                sb.Append(src.Substring(0, charPlace))
                If i + 1 < ordinalRule.Length Then
                    sb.Append(ordinalRule.Chars(i + 1))
                    sb.Append(src.Substring(charPlace + 1))
                    Return sb.ToString
                Else
                    sb.Append(ordinalRule.Chars(0))
                    sb.Append(src.Substring(charPlace + 1))
                    Return IncrementStr(sb.ToString, ordinalRule, place + 1, isFixed)
                End If
            End If
        Next
        Throw New ArgumentException(String.Format("'{0}' の {1}桁目 '{2}' はインクリメントルール '{3}' に該当しない", _
                                                  src, place + 1, src.Substring(charPlace, 1), ordinalRule))
    End Function

    ''' <summary>
    ''' 固定長数値をインクリメントして返す
    ''' </summary>
    ''' <param name="fixedNumber">固定長数値</param>
    ''' <returns>インクリメント後の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function IncrementNumber(ByVal fixedNumber As String) As String
        Const NUMBER As String = "0123456789"
        Return IncrementStr(fixedNumber, NUMBER, 0, True)
    End Function

    ''' <summary>
    ''' 固定長英数をインクリメントして返す
    ''' </summary>
    ''' <param name="fixedAlphaNumber">固定長英数</param>
    ''' <returns>インクリメント後の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function IncrementAlphaNumber(ByVal fixedAlphaNumber As String) As String
        Const ALPHA_NUMBER As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Return IncrementStr(fixedAlphaNumber, ALPHA_NUMBER, 0, True)
    End Function

    ''' <summary>
    ''' 数値をルール文字列に従って返す
    ''' </summary>
    ''' <param name="ruleString">ルール文字列</param>
    ''' <param name="value">数値</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvertValueToRuleString(ByVal ruleString As String, ByVal value As Integer) As String
        Dim length As Integer = ruleString.Length
        Dim truncate As Integer = CInt(Math.Truncate(value / length))
        If truncate <= 0 Then
            Return ruleString.Chars(value)
        End If
        Dim fuga As Integer = value Mod length
        Return ConvertValueToRuleString(ruleString, truncate) & ruleString.Chars(fuga)
    End Function

    ''' <summary>
    ''' 大文字小文字を区別せずに、Nullを考慮して同値かを返す
    ''' </summary>
    ''' <param name="a"></param>
    ''' <param name="b"></param>
    ''' <returns>同値ならTrue</returns>
    Public Shared Function EqualsIgnoreCaseIfNull(a As Object, b As Object) As Boolean
        If a Is Nothing AndAlso b Is Nothing Then
            Return True
        ElseIf a Is Nothing OrElse b Is Nothing Then
            Return False
        End If
        Return a.ToString.Equals(b.ToString, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' 大文字小文字を区別せずに、文字列が含まれているかを返す
    ''' </summary>
    ''' <param name="array">文字列配列</param>
    ''' <param name="src">判定値</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function ContainsIgnoreCase(ByVal array As String(), ByVal src As String) As Boolean
        If src Is Nothing Then
            Return False
        End If
        Return array.Any(Function(value) src.Equals(value, StringComparison.OrdinalIgnoreCase))
    End Function

    ''' <summary>
    ''' 前後のスペースや文字ではない機械コードがあれば除去する
    ''' </summary>
    ''' <param name="value">値</param>
    ''' <returns>除去後の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function Trim(ByVal value As String) As String
        Return TrimStart(TrimEnd(value))
    End Function

    ''' <summary>
    ''' 末尾にスペースや文字ではない機械コードがあれば除去する
    ''' </summary>
    ''' <param name="value">文字列</param>
    ''' <returns>除去後の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function TrimEnd(ByVal value As String) As String
        Return TrimInvalidChar(value, TrimMode.End)
    End Function

    ''' <summary>
    ''' 開始位置にスペースや文字ではない機械コードがあれば除去する
    ''' </summary>
    ''' <param name="value">文字列</param>
    ''' <returns>除去後の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function TrimStart(ByVal value As String) As String
        Return TrimInvalidChar(value, TrimMode.Start)
    End Function

    Private Enum TrimMode
        [End]
        [Start]
    End Enum
    ''' <summary>
    ''' スペースや文字ではない機械コードがあれば除去する
    ''' </summary>
    ''' <param name="value">文字列</param>
    ''' <param name="mode">開始位置か末尾か</param>
    ''' <returns>除去後の文字列</returns>
    ''' <remarks></remarks>
    Private Shared Function TrimInvalidChar(ByVal value As String, ByVal mode As TrimMode) As String
        Const INVALID_CHAR_DEL As Byte = &H7F
        Const VALID_CHAR_START As Byte = &H20    ' 0x20 = " "c = スペース
        If StringUtil.IsEmpty(value) Then
            If value Is Nothing Then
                Return value
            End If
            Return value.TrimEnd
        End If
        Dim chars As Char() = value.ToCharArray
        Dim length As Integer = chars.Length
        If mode = TrimMode.End Then
            For subtrahend As Integer = 1 To length
                Dim bytes As Byte() = System.Text.Encoding.GetEncoding(932).GetBytes(chars(length - subtrahend))
                If bytes.Length = 1 AndAlso (bytes(0) <= VALID_CHAR_START OrElse bytes(0) = INVALID_CHAR_DEL) Then
                    Continue For
                End If
                Return New String(chars, 0, length - subtrahend + 1)
            Next
        Else
            For i As Integer = 0 To length - 1
                Dim bytes As Byte() = System.Text.Encoding.GetEncoding(932).GetBytes(chars(i))
                If bytes.Length = 1 AndAlso (bytes(0) <= VALID_CHAR_START OrElse bytes(0) = INVALID_CHAR_DEL) Then
                    Continue For
                End If
                Return New String(chars, i, length - i)
            Next
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' 小数点以下末尾のゼロを取り除く
    ''' </summary>
    ''' <param name="decimalValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function TrimEndDecimalZero(ByVal decimalValue As String) As String
        If decimalValue Is Nothing Then
            Return Nothing
        End If

        If decimalValue.IndexOf(".") < 0 Then
            Return decimalValue
        End If

        Dim values As String() = Split(decimalValue, ".")
        Dim integerPart As String = values(0)
        Dim decimalPart As String = values(1)

        Dim pointZero As Integer = decimalPart.Length
        For i As Integer = 1 To decimalPart.Length
            If "0"c <> decimalPart.Chars(decimalPart.Length - i) Then
                Exit For
            End If
            pointZero = decimalPart.Length - i
        Next

        If pointZero = 0 Then
            Return integerPart
        End If
        Return integerPart & "." & decimalPart.Substring(0, pointZero)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後にSortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareNullsLast(ByVal x As Object, ByVal y As Object) As Integer
        Return CompareNullsLast(ToString(x), ToString(y))
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後に降順Sortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareDescNullsLast(ByVal x As Object, ByVal y As Object) As Integer
        Return CompareDescNullsLast(ToString(x), ToString(y))
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後にSortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareNullsLast(ByVal x As String, ByVal y As String) As Integer
        Dim result As Integer? = EzUtil.CompareObjectNullsLast(x, y)
        If result.HasValue Then
            Return result.Value
        End If
        Return x.CompareTo(y)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後に降順Sortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareDescNullsLast(ByVal x As String, ByVal y As String) As Integer
        Dim result As Integer? = EzUtil.CompareObjectNullsLast(x, y)
        If result.HasValue Then
            Return result.Value
        End If
        Return y.CompareTo(x)
    End Function

    ''' <summary>
    ''' パラメータをインデックス値に変換する
    ''' </summary>
    ''' <param name="table">変換テーブル</param>
    ''' <param name="value">パラメータ</param>
    ''' <returns>インデックス値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvToIndexFromTable(ByVal table As String, ByVal value As String) As Integer
        Dim chr As Char = Convert.ToChar(value.ToUpper)
        Dim index As Integer = table.IndexOf(chr)
        If index < 0 Then
            Throw New ArgumentException(String.Format("パラメータに、{0}以外が含まれている", table), "value")
        End If
        Return index
    End Function

    ''' <summary>
    ''' ランダムな文字の組み合わせを作成する
    ''' </summary>
    ''' <param name="strCandidate">文字候補</param>
    ''' <param name="figure">桁数</param>
    ''' <returns>ランダムな文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function MakeRandomString(ByVal strCandidate As String, ByVal figure As Integer) As String

        Dim chars As Char() = strCandidate.ToCharArray
        Dim maxValue As Integer = chars.Length

        Dim result As New StringBuilder
        Dim aRandom As New Random
        Dim count As Integer = 0
        While count < figure
            Dim index As Integer = aRandom.Next(maxValue) ' 0以上 maxValue未満 を返す
            result.Append(chars(index))
            count += 1
        End While
        Return result.ToString
    End Function

    Private Const DQ As Char = """"c

    ''' <summary>
    ''' ダブルコーテーションに囲まれた部位を一文字列として切り出す
    ''' </summary>
    ''' <param name="expression">文字列</param>
    ''' <param name="delimiter">区切り文字</param>
    ''' <returns>切り分けられた文字列の配列</returns>
    ''' <remarks></remarks>
    Public Shared Function SplitForEnclosedDQ(ByVal expression As String, ByVal delimiter As Char) As String()
        If String.IsNullOrEmpty(expression) Then
            Return New String() {}
        End If

        Dim result As New List(Of String)
        Dim buf As New StringBuilder
        Dim isEncloseMode As Boolean = False
        Dim innerIsEncloseMode As Boolean = False
        Dim chars As Char() = expression.ToCharArray
        Dim beforeChr As Char = Nothing
        For Each chr As Char In chars
            If isEncloseMode Then
                If DQ.Equals(chr) Then
                    innerIsEncloseMode = Not innerIsEncloseMode

                ElseIf Not innerIsEncloseMode AndAlso DQ.Equals(beforeChr) Then
                    If delimiter.Equals(chr) Then
                        result.Add(RemoveEnclosedDQ(buf.ToString))
                        Clear(buf)
                        beforeChr = chr
                        isEncloseMode = False
                        innerIsEncloseMode = False
                        Continue For

                    ElseIf Char.IsWhiteSpace(chr) Then
                        Continue For
                    Else
                        Dim removedBuf As String = RemoveEnclosedDQ(buf.ToString())
                        Clear(buf)
                        buf.Append(removedBuf)
                    End If
                End If
            Else
                If buf.Length = 0 Then
                    If DQ.Equals(chr) Then
                        isEncloseMode = True
                        innerIsEncloseMode = True
                    ElseIf Char.IsWhiteSpace(chr) Then
                        Continue For
                    End If
                End If
                If delimiter.Equals(chr) Then
                    result.Add(buf.ToString.TrimEnd)
                    Clear(buf)
                    beforeChr = chr
                    Continue For
                End If
            End If
            buf.Append(chr, 1)
            beforeChr = chr
        Next
        If isEncloseMode AndAlso EnclosesInDQ(buf.ToString) Then
            result.Add(RemoveEnclosedDQ(buf.ToString))
        ElseIf Not isEncloseMode Then
            result.Add(buf.ToString.TrimEnd)
        Else
            result.Add(buf.ToString)
        End If

        Return result.ToArray
    End Function

    Private Shared Sub Clear(ByVal sb As StringBuilder)
        ' .NET 2.0 には#Clearメソッドがないみたい
        sb.Length = 0
    End Sub

    ''' <summary>
    ''' 文字列がダブルコーテーションで囲まれているか?
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function EnclosesInDQ(ByVal str As String) As Boolean
        If StringUtil.IsEmpty(str) Then
            Return False
        End If
        Dim chars As Char() = str.ToCharArray
        Return 2 <= chars.Length AndAlso chars(0) = DQ AndAlso chars(chars.Length - 1) = DQ
    End Function

    ''' <summary>
    ''' ダブルコーテーションで囲まれた文字列から本来の文字列を取得する
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns>本来の文字列</returns>
    ''' <remarks></remarks>
    Private Shared Function RemoveEnclosedDQ(ByVal str As String) As String
        Const DQDQ As String = """"""
        str = str.Remove(0, 1)
        str = str.Remove(str.Length - 1, 1)
        str = str.Replace(DQDQ, DQ)
        Return str
    End Function

    ''' <summary>
    ''' 指定の長さで切る
    ''' </summary>
    ''' <param name="value">文字列値</param>
    ''' <param name="length">文字数</param>
    ''' <returns>切った文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function Left(ByVal value As Object, ByVal length As Integer) As String
        Return Left(ToString(value), length)
    End Function

    ''' <summary>
    ''' 指定の長さで切る
    ''' </summary>
    ''' <param name="value">文字列値</param>
    ''' <param name="length">文字数</param>
    ''' <returns>切った文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function Left(ByVal value As String, ByVal length As Integer) As String
        If value Is Nothing Then
            Return Nothing
        End If
        Return If(length < value.Length, value.Substring(0, length), value)
    End Function

    ''' <summary>
    ''' 半角数字か？を返す
    ''' </summary>
    ''' <param name="c">Char</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNumberHalf(ByVal c As Char) As Boolean
        ' Char.IsDigit は 全角数字もtrueになるので、独自実装
        Const ZERO As Char = "0"c
        Const NINE As Char = "9"c
        Return ZERO <= c AndAlso c <= NINE
    End Function

    ''' <summary>
    ''' 半角アルファベット文字か？を返す
    ''' </summary>
    ''' <param name="c">Char</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsAlphabetHalf(ByVal c As Char) As Boolean
        ' Char.IsLetter は 半角カタカナもtrueになるので、独自実装
        Const LOWER_A As Char = "a"c
        Const LOWER_Z As Char = "z"c
        Const UPPER_A As Char = "A"c
        Const UPPER_Z As Char = "Z"c
        Return (LOWER_A <= c AndAlso c <= LOWER_Z) OrElse (UPPER_A <= c AndAlso c <= UPPER_Z)
    End Function

    ''' <summary>
    ''' 半角英数字のみの文字列かを返す
    ''' </summary>
    ''' <param name="str">判定文字列</param>
    ''' <param name="allowChars">英数字以外に許可する文字[]</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsAlphabetOrNumber(ByVal str As String, ByVal ParamArray allowChars As Char()) As Boolean
        If String.IsNullOrEmpty(str) Then
            Return False
        End If
        Dim allowChars2 As New List(Of Char)(allowChars)
        Dim chars As Char() = str.ToCharArray
        Return chars.All(Function(c) IsAlphabetHalf(c) OrElse IsNumberHalf(c) OrElse allowChars2.Contains(c))
    End Function

    ''' <summary>
    ''' 開始～終了に囲まれた文字列を返す
    ''' </summary>
    ''' <param name="message">対象文字列</param>
    ''' <param name="startStr">開始文字列</param>
    ''' <param name="endStr">終了文字列</param>
    ''' <returns>開始～終了に囲まれた文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function ExtractEnclosedString(ByVal message As String, ByVal startStr As String, ByVal endStr As String) As String
        Dim indexOfStart As Integer = message.IndexOf(startStr)
        Dim indexOfEnd As Integer = message.IndexOf(endStr, indexOfStart + 1)
        If indexOfStart < 0 OrElse indexOfEnd < 0 Then
            Return Nothing
        End If
        Return message.Substring(indexOfStart + 1, indexOfEnd - indexOfStart - 1)
    End Function

    ''' <summary>
    ''' 数値の固まりとそれ以外の固まりとに分解する
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SplitNumberAlpha(ByVal str As String) As String()
        Dim chars As Char() = str.ToCharArray
        Dim sb As New StringBuilder
        Dim results As New List(Of String)
        For Each c As Char In chars
            If 0 < sb.Length AndAlso Char.IsDigit(c) <> IsNumeric(sb.ToString) Then
                results.Add(sb.ToString)
                sb.Length = 0
            End If
            sb.Append(c)
        Next
        If 0 < sb.Length Then
            results.Add(sb.ToString)
        End If
        Return results.ToArray
    End Function

    ''' <summary>
    ''' 最大長を超えていたら、末尾を省略して返す
    ''' </summary>
    ''' <param name="value">文字列</param>
    ''' <param name="maxByteLength">最大長(byte)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function OmitIfLengthOver(ByVal value As String, ByVal maxByteLength As Integer) As String
        Const BYTE_LENGTH_OVER_SUFFIX As String = "..."
        Return OmitIfLengthOver(value, maxByteLength, BYTE_LENGTH_OVER_SUFFIX)
    End Function

    ''' <summary>
    ''' 最大長を超えていたら、末尾を省略して返す
    ''' </summary>
    ''' <param name="value">文字列</param>
    ''' <param name="maxByteLength">最大長(byte)</param>
    ''' <param name="suffixIfLengthOver">省略するときの末尾文字列</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function OmitIfLengthOver(ByVal value As String, ByVal maxByteLength As Integer, ByVal suffixIfLengthOver As String) As String
        If value Is Nothing Then
            Return Nothing
        End If
        If StringUtil.GetLengthByte(value) <= maxByteLength Then
            Return value
        End If
        Dim suffix As String = If(maxByteLength < StringUtil.GetLengthByte(suffixIfLengthOver), String.Empty, suffixIfLengthOver)
        Dim aMessage As String = value
        Dim lengthOfSuffix As Integer = StringUtil.GetLengthByte(suffix)
        While maxByteLength - lengthOfSuffix < StringUtil.GetLengthByte(aMessage)
            aMessage = aMessage.Substring(0, aMessage.Length - 1)
        End While
        Return aMessage & suffix
    End Function

    ''' <summary>
    ''' 引数がNULLの場合はNothingを返し、あれば大文字に変える
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns>Nothingか文字列を大文字に変換したもの</returns>
    ''' <remarks></remarks>
    Public Shared Function ToUpper(ByVal str As String) As String
        If str Is Nothing Then
            Return Nothing
        End If
        Return str.ToUpper
    End Function

    ''' <summary>
    ''' 半角文字にする
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ToHankaku(ByVal str As String) As String
        If str Is Nothing Then
            Return Nothing
        End If
        ' SpecFlowから実行すると "ロケーションID(1033)が違うエラー" になる。なので個別指定。
        Return Strings.StrConv(str, VbStrConv.Narrow, LocaleID:=1041)
    End Function

    ''' <summary>
    ''' 全角文字にする
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ToZenkaku(ByVal str As String) As String
        If str Is Nothing Then
            Return Nothing
        End If
        ' SpecFlowから実行すると "ロケーションID(1033)が違うエラー" になる。なので個別指定。
        Return Strings.StrConv(str, VbStrConv.Wide, LocaleID:=1041)
    End Function

    ''' <summary>
    ''' JIS X 0208規格における機種依存文字かどうか？を返す
    ''' </summary>
    ''' <param name="c">文字</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsGaijiForJISX0208(ByVal c As Char) As Boolean
        ' SJIS 0x85xx～0x889E, 0xEBxx～0xFCxx は機種依存文字 http://www5d.biglobe.ne.jp/~noocyte/Programming/CharCode.html
        ' UTF8固有の文字はSJIS変換後'?'になる
        Const QUESTION As Char = "?"c
        Const SJIS_CODE_QUESTION As Integer = &H3F
        Dim sjis As Byte() = Encoding.GetEncoding("SHIFT_JIS").GetBytes(New Char() {c})
        'デバッグ用に残す
        'Dim utf8 As Byte() = Encoding.GetEncoding("UTF-8").GetBytes(New Char() {c})
        'Debug.Print(String.Format("char='{0}' UTF8={1} SJIS={2}", c, BitConverter.ToString(utf8), BitConverter.ToString(sjis)))
        Return (sjis.Length = 2 AndAlso ((&H85 <= sjis(0) AndAlso sjis(0) <= &H87) _
                                         OrElse (&H88 = sjis(0) AndAlso sjis(1) < &H9F) _
                                         OrElse (&HEB <= sjis(0) AndAlso sjis(0) <= &HFC))) _
                OrElse (QUESTION <> c AndAlso sjis.Length = 1 AndAlso sjis(0) = SJIS_CODE_QUESTION)
    End Function

    ''' <summary>
    ''' スペースを取り除く
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns>全角スペース、半角スペースを取り除いた文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function RemoveSpaceCharacters(ByVal str As String) As String
        If str Is Nothing Then
            Return Nothing
        End If
        Const ONE_BYTE_SPACE As Char = " "c
        Const DOUBLE_BYTE_SPACE As Char = "　"c
        Return str.Replace(DOUBLE_BYTE_SPACE, String.Empty).Replace(ONE_BYTE_SPACE, String.Empty)
    End Function

    ''' <summary>
    ''' 桁数で文字列を分割する
    ''' </summary>
    ''' <param name="base">対象文字列</param>
    ''' <param name="size">何桁ずつ分割したいか</param>
    ''' <returns>桁数ごとに分割された文字列[]</returns>
    ''' <remarks></remarks>
    Public Shared Function SplitBySize(base As String, size As Integer) As String()
        If base Is Nothing Then
            base = ""
        End If
        Dim maxLength As Integer = base.Length
        Dim result As IEnumerable(Of String) = Linq.Enumerable.Range(0, maxLength \ size).Select(Function(index) base.Substring(index * size, size))
        Dim remainder As Integer = maxLength Mod size
        If remainder = 0 Then
            Return result.ToArray
        End If
        Return result.Concat({base.Substring(maxLength - remainder)}).ToArray
    End Function

    ''' <summary>
    ''' 数値文字列を抜き出す
    ''' </summary>
    ''' <param name="text">文字列</param>
    ''' <returns>最初に見つかった数値</returns>
    ''' <remarks></remarks>
    Public Shared Function ExtractNumber(text As String) As Integer
        Dim startIndex As Integer = -1
        Dim chars As Char() = If(text Is Nothing, New Char() {}, text.ToCharArray())
        For index As Integer = 0 To chars.Length - 1
            Dim isEffectiveNumber As Boolean = IsHankakuOnly(chars(index)) AndAlso IsNumeric(chars(index))
            If 0 <= startIndex Then
                If Not isEffectiveNumber Then
                    Return CInt(text.Substring(startIndex, index - startIndex))
                End If
            Else
                If isEffectiveNumber Then
                    startIndex = index
                End If
            End If
        Next
        If startIndex < 0 Then
            Throw New ArgumentException(String.Format("値 '{0}' に数値は含まれない", text), "text")
        End If
        Return CInt(text.Substring(startIndex))
    End Function

    ''' <summary>
    ''' バイト数をSi接頭語付き表記に変換する
    ''' </summary>
    ''' <param name="byteNum">バイト数</param>
    ''' <returns>Si接頭語付きバイト数</returns>
    Public Shared Function ConvByteToWithSiPrefix(byteNum As ULong) As String
        If byteNum < 1000 Then
            Return String.Format("{0:#,0}B", byteNum)
        End If
        Dim siPrefixes As String() = {"K", "M", "G", "T"}
        Dim divisionCount As Integer = 0
        Dim result As ULong = byteNum
        Do While 1000000 <= result AndAlso divisionCount < 3
            result = CULng(result / 1024)
            divisionCount += 1
        Loop
        Return String.Format("{0:#.#}{1}B", result / 1024, siPrefixes(divisionCount))
    End Function

End Class
