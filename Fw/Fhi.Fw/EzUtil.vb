Imports System.Text
Imports System.Security.Cryptography
Imports Fhi.Fw.Domain

''' <summary>
''' 簡素な共通処理を集めたユーティリティクラス
''' </summary>
''' <remarks></remarks>
Public Class EzUtil
    ''' <summary>
    ''' C/Java等の i++ を実現する
    ''' </summary>
    ''' <param name="i">i++する値</param>
    ''' <returns>i の値</returns>
    ''' <remarks></remarks>
    Public Shared Function Increment(ByRef i As Integer) As Integer
        Dim result As Integer = i
        i += 1
        Return result
    End Function

    ''' <summary>
    ''' Null値を考慮して同値かを返す
    ''' </summary>
    ''' <param name="a">値a</param>
    ''' <param name="b">値b</param>
    ''' <returns>Null同士、または同値ならtrue</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEqualIfNull(ByVal a As Object, ByVal b As Object) As Boolean
        If a Is Nothing AndAlso b Is Nothing Then
            Return True
        ElseIf a Is Nothing OrElse b Is Nothing Then
            Return False
        End If
        Return a.Equals(b)
    End Function

    ''' <summary>
    ''' 値(null含む)が違うかを返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>違う場合、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEqualIfNull(ByVal x As Object, ByVal y As Object) As Boolean
        Return Not IsEqualIfNull(x, y)
    End Function

    ''' <summary>
    ''' PVOかを無視して、Null値を考慮して同値かを返す
    ''' </summary>
    ''' <param name="a">値a</param>
    ''' <param name="b">値b</param>
    ''' <returns>Null同士、または同値ならtrue</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEqualIgnorePvoIfNull(a As Object, b As Object) As Boolean
        Dim valueA As Object = a
        If a IsNot Nothing AndAlso TypeOf a Is PrimitiveValueObject Then
            valueA = DirectCast(a, PrimitiveValueObject).Value
        End If
        Dim valueB As Object = b
        If b IsNot Nothing AndAlso TypeOf b Is PrimitiveValueObject Then
            valueB = DirectCast(b, PrimitiveValueObject).Value
        End If
        Return IsEqualIfNull(valueA, valueB)
    End Function

    ''' <summary>
    ''' PVOかを無視して、値(null含む)が違うかを返す
    ''' </summary>
    ''' <param name="a">値a</param>
    ''' <param name="b">値b</param>
    ''' <returns>違う場合、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEqualIgnorePvoIfNull(a As Object, b As Object) As Boolean
        Return Not IsEqualIgnorePvoIfNull(a, b)
    End Function

    Private Const ALPHABET_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    ''' <summary>
    ''' 番号(0 start)をAlphabet表記にする
    ''' </summary>
    ''' <param name="index">0から始まる番号</param>
    ''' <returns>"A"とか"BB"とか</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvIndexToAlphabet(ByVal index As Integer) As String
        Dim chars As Char() = ALPHABET_TABLE.ToCharArray
        Dim value As Integer = index
        Dim len As Integer = chars.Length

        Dim result As String = ""

        value += 1
        Do
            Dim amari As Integer = (value - 1) Mod len
            result = chars(amari) & result
            value = Convert.ToInt32(Math.Truncate((value - 1) / len))
        Loop While 0 < value
        Return result
    End Function

    ''' <summary>
    ''' (Excelとかの)Alphabet表記のinedx値を数値(0 start)をにする
    ''' </summary>
    ''' <param name="alphabetIndex">Alphabet表記のinedx値</param>
    ''' <returns>index値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvAlphabetToIndex(ByVal alphabetIndex As String) As Integer
        AssertParameterIsNotNull(alphabetIndex, "alphabetIndex")
        Dim len As Integer = ALPHABET_TABLE.Length

        Dim result As Integer = 0
        For Each index As Integer In From c In alphabetIndex.ToUpper.ToCharArray Select ALPHABET_TABLE.IndexOf(c)
            If index < 0 Then
                Throw New ArgumentException("アルファベット以外が含まれている", "alphabetIndex")
            End If
            result = (result) * len + index + 1
        Next

        Return result - 1
    End Function

    ''' <summary>
    ''' 指定キーの組み合わせで一意になる値を作成する
    ''' </summary>
    ''' <param name="keys">キー()</param>
    ''' <returns>一意のキー</returns>
    ''' <remarks></remarks>
    Public Shared Function MakeKey(ByVal ParamArray keys As Object()) As String
        Const SEPARATOR As String = ";'"":"
        Const SEPARATOR_LENGTH As Integer = 4
        'Return Join(keys, SEPARATOR)
        ' パフォーマンスチューニング
        ' 165922回呼び出しにて404.08(ms)が165.922(ms)に。
        If keys Is Nothing OrElse keys.Length = 0 Then
            Return Nothing
        End If
        Dim result As New StringBuilder
        For Each key As Object In keys
            result.Append(key).Append(SEPARATOR)
        Next
        result.Length -= SEPARATOR_LENGTH
        Return result.ToString
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後にSortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareNullsLast(ByVal x As Object, ByVal y As Object) As Integer
        If x Is Nothing OrElse y Is Nothing Then
            Return CompareObjectNullsLast(x, y).Value
        End If
        If Not x.GetType Is y.GetType() Then
            Throw New ArgumentException(String.Format("型が合っていない. x As {0}, y As {1}", x.GetType.Name, y.GetType.Name))
        End If
        If TypeOf x Is String AndAlso TypeOf y Is String Then
            Return StringUtil.CompareNullsLast(x, y)
        End If
        If TypeOf x Is DateTime AndAlso TypeOf y Is DateTime Then
            Return DateUtil.CompareNullsLast(x, y)
        End If
        If TypeOf x Is IComparable AndAlso Not TypeOf x Is Boolean Then
            Return DirectCast(x, IComparable).CompareTo(DirectCast(y, IComparable))
        End If
        If Not IsNumeric(x) OrElse Not IsNumeric(y) Then
            Throw New ArgumentException(String.Format("未対応の型です. x as {0}, y as {1}", x.GetType, y.GetType))
        End If
        Return CompareNullsLast(CDec(x), CDec(y))
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後に降順Sortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら 1, x &gt; y なら -1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareDescNullsLast(ByVal x As Object, ByVal y As Object) As Integer
        If x Is Nothing OrElse y Is Nothing Then
            Return CompareObjectNullsLast(x, y).Value
        End If
        If Not x.GetType Is y.GetType() Then
            Throw New ArgumentException(String.Format("型が合っていない. x As {0}, y As {1}", x.GetType.Name, y.GetType.Name))
        End If
        If TypeOf x Is String AndAlso TypeOf y Is String Then
            Return StringUtil.CompareDescNullsLast(x, y)
        End If
        If TypeOf x Is DateTime AndAlso TypeOf y Is DateTime Then
            Return DateUtil.CompareDescNullsLast(x, y)
        End If
        If TypeOf x Is IComparable AndAlso Not TypeOf x Is Boolean Then
            Return DirectCast(y, IComparable).CompareTo(DirectCast(x, IComparable))
        End If
        If Not IsNumeric(x) OrElse Not IsNumeric(y) Then
            Throw New ArgumentException(String.Format("未対応の型です. x as {0}, y as {1}", x.GetType, y.GetType))
        End If
        Return CompareDescNullsLast(CDec(x), CDec(y))
    End Function

    Private Delegate Function CompareToCallback(Of T)(ByVal x As t, ByVal y As t) As Integer
    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後にSortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Private Shared Function PerformCompareNullsLast(Of T As Structure)(ByVal callback As CompareToCallback(Of T), ByVal x As T?, ByVal y As T?) As Integer
        Dim result As Integer? = CompareObjectNullsLast(x, y)
        If result.HasValue Then
            Return result.Value
        End If
        Return callback(x.Value, y.Value)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後に降順Sortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら 1, x &gt; y なら -1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Private Shared Function PerformCompareDescNullsLast(Of T As Structure)(ByVal callback As CompareToCallback(Of T), ByVal x As T?, ByVal y As T?) As Integer
        Dim result As Integer? = CompareObjectNullsLast(x, y)
        If result.HasValue Then
            Return result.Value
        End If
        Return callback(y.Value, x.Value)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後にSortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareNullsLast(ByVal x As Decimal?, ByVal y As Decimal?) As Integer
        Return PerformCompareNullsLast(Of Decimal)(Function(x2, y2) x2.CompareTo(y2), x, y)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後に降順Sortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら 1, x &gt; y なら -1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareDescNullsLast(ByVal x As Decimal?, ByVal y As Decimal?) As Integer
        Return PerformCompareDescNullsLast(Of Decimal)(Function(x2, y2) x2.CompareTo(y2), x, y)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後にSortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareNullsLast(ByVal x As Long?, ByVal y As Long?) As Integer
        Return PerformCompareNullsLast(Of Long)(Function(x2, y2) x2.CompareTo(y2), x, y)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後に降順Sortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら 1, x &gt; y なら -1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareDescNullsLast(ByVal x As Long?, ByVal y As Long?) As Integer
        Return PerformCompareDescNullsLast(Of Long)(Function(x2, y2) x2.CompareTo(y2), x, y)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後にSortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら -1, x &gt; y なら 1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareNullsLast(ByVal x As Integer?, ByVal y As Integer?) As Integer
        Return PerformCompareNullsLast(Of Integer)(Function(x2, y2) x2.CompareTo(y2), x, y)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後に降順Sortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>x &lt; y なら 1, x &gt; y なら -1, 等しければ 0</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareDescNullsLast(ByVal x As Integer?, ByVal y As Integer?) As Integer
        Return PerformCompareDescNullsLast(Of Integer)(Function(x2, y2) x2.CompareTo(y2), x, y)
    End Function

    ''' <summary>
    ''' IComparer(Of T)実装用に、Null値だったら最後にSortする値を返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>Null値だったら最後にSortする値. それ以外は、Nothing</returns>
    ''' <remarks></remarks>
    Public Shared Function CompareObjectNullsLast(ByVal x As Object, ByVal y As Object) As Integer?
        If x Is Nothing AndAlso y Is Nothing Then
            Return 0
        ElseIf x Is Nothing Then
            Return 1
        ElseIf y Is Nothing Then
            Return -1
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' パラメータがnullで無い事を保証する(nullの場合、例外を投げる)
    ''' </summary>
    ''' <param name="parameter">引数</param>
    ''' <param name="name">引数名</param>
    ''' <remarks></remarks>
    Public Shared Sub AssertParameterIsNotNull(ByVal parameter As Object, ByVal name As String)
        If parameter Is Nothing Then
            Throw New ArgumentNullException(name)
        End If
    End Sub

    ''' <summary>
    ''' パラメータがemptyで無い事を保証する(emptyの場合、例外を投げる)
    ''' </summary>
    ''' <param name="parameter">引数</param>
    ''' <param name="name">引数名</param>
    ''' <remarks></remarks>
    Public Shared Sub AssertParameterIsNotEmpty(ByVal parameter As Object, ByVal name As String)
        AssertParameterIsNotNull(parameter, name)
        If TypeOf parameter Is String Then
            If StringUtil.IsNotEmpty(parameter) Then
                Return
            End If
        ElseIf TypeOf parameter Is IEnumerable Then
            Dim values As IEnumerable = DirectCast(parameter, IEnumerable)
            If CollectionUtil.IsNotEmpty(values) Then
                Return
            End If
        ElseIf TypeOf parameter Is ICollectionObject Then
            Dim values As ICollectionObject = DirectCast(parameter, ICollectionObject)
            If CollectionUtil.IsNotEmpty(values) Then
                Return
            End If
        ElseIf TypeOf parameter Is PrimitiveValueObject Then
            AssertParameterIsNotEmpty(DirectCast(parameter, PrimitiveValueObject).Value, name)
            Return
        Else
            Throw New NotSupportedException("String型、配列、またはコレクションにのみ対応")
        End If
        Throw New ArgumentException("emptyです.", name)
    End Sub

    ''' <summary>
    ''' NULL値だったら、置き換え値に変換 Null Value Logic
    ''' </summary>
    ''' <typeparam name="T">値の型</typeparam>
    ''' <param name="value">判定値</param>
    ''' <param name="nullValue">置き換え値</param>
    ''' <returns>値</returns>
    ''' <remarks></remarks>
    Public Shared Function Nvl(Of T As Structure)(ByVal value As Nullable(Of T), ByVal nullValue As T) As T
        If value Is Nothing Then
            Return nullValue
        End If
        Return value.Value
    End Function

    ''' <summary>
    ''' NULL値だったら、置き換え値に変換 Null Value Logic
    ''' </summary>
    ''' <param name="value">判定値</param>
    ''' <param name="nullValue">置き換え値</param>
    ''' <returns>値</returns>
    ''' <remarks></remarks>
    Public Shared Function Nvl(ByVal value As Integer?, ByVal nullValue As Integer) As Integer
        Return Nvl(Of Integer)(value, nullValue)
    End Function

    ''' <summary>
    ''' NULL値だったら、置き換え値に変換 Null Value Logic
    ''' </summary>
    ''' <param name="value">判定値</param>
    ''' <param name="nullValue">置き換え値</param>
    ''' <returns>値</returns>
    ''' <remarks></remarks>
    Public Shared Function Nvl(ByVal value As Decimal?, ByVal nullValue As Decimal) As Decimal
        Return Nvl(Of Decimal)(value, nullValue)
    End Function

    ''' <summary>
    ''' NULL値だったら、置き換え値に変換 Null Value Logic
    ''' </summary>
    ''' <typeparam name="T">値の型</typeparam>
    ''' <param name="value">判定値</param>
    ''' <param name="nullValue">置き換え値</param>
    ''' <returns>値</returns>
    ''' <remarks></remarks>
    Public Shared Function Nvl(Of T As PrimitiveValueObject)(value As T, nullValue As T) As T
        If value Is Nothing OrElse value.Value Is Nothing Then
            Return nullValue
        End If
        Return value
    End Function

    ''' <summary>
    ''' 配列から、Listを作成する
    ''' </summary>
    ''' <typeparam name="T">型</typeparam>
    ''' <param name="values">値</param>
    ''' <returns>List値</returns>
    ''' <remarks></remarks>
    Public Shared Function NewList(Of T)(ByVal ParamArray values As T()) As List(Of T)
        Return CollectionUtil.NewList(Of T)(values)
    End Function

    ''' <summary>
    ''' 真かを返す
    ''' </summary>
    ''' <param name="value">null許容のboolean値</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsTrue(ByVal value As Boolean?) As Boolean
        If value Is Nothing Then
            Return False
        End If
        Return value.Value
    End Function

    ''' <summary>
    ''' 真かを返す
    ''' </summary>
    ''' <param name="value">null許容のboolean値</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsTrue(ByVal value As Object) As Boolean
        If value Is Nothing Then
            Return False
        End If
        If TypeOf value Is Boolean Then
            Return DirectCast(value, Boolean)
        End If
        If IsNumeric(value) Then
            Return CInt(value) <> 0
        End If
        Dim str As String = value.ToString.ToLower
        Return "true".Equals(str)
    End Function

    ''' <summary>
    ''' Boolean値そのものか、あるいは"TRUE"/"FALSE"か？を返す
    ''' </summary>
    ''' <param name="value">判定値</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsBooleanValue(ByVal value As Object) As Boolean
        If value Is Nothing Then
            Return False
        End If
        Dim dummy As Boolean
        Return Boolean.TryParse(value.ToString, dummy)
    End Function

    ''' <summary>
    ''' Boolean値へ変換する
    ''' </summary>
    ''' <param name="value">値</param>
    ''' <returns>変換後の値 ※変換できなければnull</returns>
    ''' <remarks></remarks>
    Public Shared Function ParseAsBoolean(ByVal value As Object) As Boolean?
        If value Is Nothing Then
            Return Nothing
        End If
        If TypeOf value Is Boolean? Then
            Return DirectCast(value, Boolean?)
        End If
        If TypeOf value Is Boolean Then
            Return DirectCast(value, Boolean)
        End If
        If IsNumeric(value) Then
            Return CInt(value) <> 0
        End If
        Dim parsedValue As Boolean
        If Boolean.TryParse(value.ToString, parsedValue) Then
            Return parsedValue
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 端末日時を取得する
    ''' </summary>
    ''' <returns>端末日時</returns>
    ''' <remarks></remarks>
    Public Shared Function GetNowDateTime() As DateTime
        Return DateUtil.GetNowDateTime
    End Function

    ''' <summary>
    ''' コンソール出力
    ''' </summary>
    ''' <param name="message">出力メッセージ</param>
    ''' <param name="args">埋め込み引数</param>
    ''' <remarks></remarks>
    Public Shared Sub logDebug(ByVal message As String, ByVal ParamArray args As Object())
        Const TIME_FORMAT As String = "HH:mm:ss,fff"
        Const SPACE As String = " "
        Dim sb As New StringBuilder
        sb.Append(GetNowDateTime.ToString(TIME_FORMAT))
        sb.Append(SPACE)
        If 0 < args.Length Then
            sb.AppendFormat(message, args)
        Else
            sb.Append(message)
        End If
        Debug.Print(sb.ToString)
    End Sub

    ''' <summary>
    ''' リスト(配列)の値(要素)を全て結合した文字列を返す
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="delimiter">区切り文字</param>
    ''' <param name="lists">結合するリスト(配列)</param>
    ''' <returns>結合した文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function JoinAsString(Of T)(ByVal delimiter As String, ByVal ParamArray lists As IEnumerable(Of T)()) As String
        Dim result As New StringBuilder
        For Each collection As IEnumerable(Of T) In lists
            For Each value As T In collection
                result.Append(value.ToString)
                result.Append(delimiter)
            Next
        Next
        If result.Length = 0 Then
            Return Nothing
        End If
        result.Length -= delimiter.Length
        Return result.ToString
    End Function

    ''' <summary>
    ''' Empty値、またはNull値か、を返す
    ''' </summary>
    ''' <param name="str">判定する文字列</param>
    ''' <returns>Empty値かNull値の場合、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmpty(ByVal str As String) As Boolean
        Return StringUtil.IsEmpty(str)
    End Function

    ''' <summary>
    ''' Empty値以外でかつ、Null値以外か、を返す
    ''' </summary>
    ''' <param name="str">判定する文字列</param>
    ''' <returns>Empty値以外でかつ、Null値以外の場合、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEmpty(ByVal str As String) As Boolean
        Return StringUtil.IsNotEmpty(str)
    End Function

    ''' <summary>
    ''' 値を持つコレクションか？を返す
    ''' </summary>
    ''' <param name="collection">コレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEmpty(ByVal collection As IEnumerable) As Boolean
        Return CollectionUtil.IsNotEmpty(collection)
    End Function

    ''' <summary>
    ''' 値を持たないコレクションか？を返す
    ''' </summary>
    ''' <param name="collection">コレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmpty(ByVal collection As IEnumerable) As Boolean
        Return CollectionUtil.IsEmpty(collection)
    End Function

    ''' <summary>
    ''' 値を持つコレクションか？を返す
    ''' </summary>
    ''' <param name="collection">コレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEmpty(ByVal collection As ICollectionObject) As Boolean
        Return CollectionUtil.IsNotEmpty(collection)
    End Function

    ''' <summary>
    ''' 値を持たないコレクションか？を返す
    ''' </summary>
    ''' <param name="collection">コレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmpty(ByVal collection As ICollectionObject) As Boolean
        Return CollectionUtil.IsEmpty(collection)
    End Function

    ''' <summary>
    ''' 値があるか？を返す
    ''' </summary>
    ''' <param name="value">PrimitiveValueObject</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEmpty(value As PrimitiveValueObject) As Boolean
        Return Not IsEmpty(value)
    End Function

    ''' <summary>
    ''' 値がない、もしくは初期値か？を返す
    ''' </summary>
    ''' <param name="value">PrimitiveValueObject</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmpty(value As PrimitiveValueObject) As Boolean
        If value Is Nothing OrElse value.Value Is Nothing Then
            Return True
        End If
        Return value.IsEmpty
    End Function

    ''' <summary>
    ''' 値があるか？を返す
    ''' </summary>
    ''' <param name="value">コレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEmpty(Of T As Structure)(ByVal value As Nullable(Of T)) As Boolean
        Return Not IsEmpty(Of T)(value)
    End Function

    ''' <summary>
    ''' 値がない、もしくは初期値か？を返す
    ''' </summary>
    ''' <param name="value">コレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmpty(Of T As Structure)(ByVal value As Nullable(Of T)) As Boolean
        Return Not value.HasValue OrElse value.Value.Equals(Activator.CreateInstance(GetType(T)))
    End Function

    ''' <summary>
    ''' 複数の値から最大値を返す
    ''' </summary>
    ''' <param name="values">値[]</param>
    ''' <returns>最大値</returns>
    ''' <remarks></remarks>
    Public Shared Function MaxInArray(ByVal ParamArray values As Integer()) As Integer
        Return PerformCompareInArray(Of Integer)(Function(a As Integer, b As Integer) Math.Max(a, b), values)
    End Function

    ''' <summary>
    ''' 複数の値から最大値を返す
    ''' </summary>
    ''' <param name="values">値[]</param>
    ''' <returns>最大値</returns>
    ''' <remarks></remarks>
    Public Shared Function MaxInArray(ByVal ParamArray values As Long()) As Long
        Return PerformCompareInArray(Of Long)(Function(a As Long, b As Long) Math.Max(a, b), values)
    End Function

    ''' <summary>
    ''' 複数の値から最大値を返す
    ''' </summary>
    ''' <param name="values">値[]</param>
    ''' <returns>最大値</returns>
    ''' <remarks></remarks>
    Public Shared Function MaxInArray(ByVal ParamArray values As Single()) As Single
        Return PerformCompareInArray(Of Single)(Function(a As Single, b As Single) Math.Max(a, b), values)
    End Function

    ''' <summary>
    ''' 複数の値から最大値を返す
    ''' </summary>
    ''' <param name="values">値[]</param>
    ''' <returns>最大値</returns>
    ''' <remarks></remarks>
    Public Shared Function MaxInArray(ByVal ParamArray values As Double()) As Double
        Return PerformCompareInArray(Of Double)(Function(a As Double, b As Double) Math.Max(a, b), values)
    End Function

    ''' <summary>
    ''' 複数の値から最大値を返す
    ''' </summary>
    ''' <param name="values">値[]</param>
    ''' <returns>最大値</returns>
    ''' <remarks></remarks>
    Public Shared Function MaxInArray(ByVal ParamArray values As Decimal()) As Decimal
        Return PerformCompareInArray(Of Decimal)(Function(a As Decimal, b As Decimal) Math.Max(a, b), values)
    End Function

    Private Delegate Function CompareCallback(Of T)(ByVal a As t, ByVal B As t) As t

    Private Shared Function PerformCompareInArray(Of T)(ByVal callback As CompareCallback(Of T), ByVal ParamArray values As T()) As T
        Dim result As T = values(0)
        Return values.Aggregate(result, Function(current, value) callback(current, value))
    End Function

    ''' <summary>比較処理</summary>
    Public Delegate Function DelegateJudgeCompare() As Integer
    ''' <summary>
    ''' 比較処理を順次実行し結果を適切に返す
    ''' </summary>
    ''' <param name="compareCallbacks">比較処理[]</param>
    ''' <returns>ソート判定に必要な値</returns>
    ''' <remarks></remarks>
    Public Shared Function JudgeCompare(ByVal ParamArray compareCallbacks As DelegateJudgeCompare()) As Integer
        Return (From compareCallback In compareCallbacks Select compareCallback.Invoke()).FirstOrDefault(Function(result) result <> 0)
    End Function
    ''' <summary>
    ''' 比較処理を順次実行し結果を適切に返す
    ''' </summary>
    ''' <param name="compareCallbacks">比較処理[]</param>
    ''' <returns>ソート判定に必要な値</returns>
    ''' <remarks></remarks>
    Public Shared Function CombineCompare(ByVal ParamArray compareCallbacks As DelegateJudgeCompare()) As Integer
        Return JudgeCompare(compareCallbacks)
    End Function

    ''' <summary>
    ''' (単精度)浮動小数点がゼロか？を返す
    ''' </summary>
    ''' <param name="s">値</param>
    ''' <returns>ゼロなら、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEqualZero(ByVal s As Single) As Boolean
        Return IsEqualIfEpsilon(s, 0.0F)
    End Function

    ''' <summary>
    ''' (単精度)浮動小数点が同じか？を返す
    ''' </summary>
    ''' <param name="a">値A</param>
    ''' <param name="b">値B</param>
    ''' <returns>同じなら、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEqualIfEpsilon(ByVal a As Single, ByVal b As Single) As Boolean
        Const EPSILON As Single = 0.00001F
        Dim size As Integer = Math.Min(GetDecimalSize(a), GetDecimalSize(b))
        Dim factor As Single = CSng(Math.Pow(10, size))
        Dim aa As Single = a * factor
        Dim bb As Single = b * factor
        Return Math.Abs(aa - bb) <= EPSILON
    End Function

    ''' <summary>
    ''' (倍精度)浮動小数点がゼロか？を返す
    ''' </summary>
    ''' <param name="d">値</param>
    ''' <returns>ゼロなら、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEqualZero(ByVal d As Double) As Boolean
        Return IsEqualIfEpsilon(d, 0.0R)
    End Function

    ''' <summary>
    ''' (ゆるい精度)浮動小数点が同じか？を返す
    ''' </summary>
    ''' <param name="a">値A</param>
    ''' <param name="b">値B</param>
    ''' <returns>同じなら、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEqualIfEpsilon(ByVal a As Double, ByVal b As Double) As Boolean
        Const EPSILON As Double = 0.00001R
        Dim size As Integer = Math.Min(GetDecimalSize(a), GetDecimalSize(b))
        Dim factor As Double = Math.Pow(10, size)
        Dim aa As Double = a * factor
        Dim bb As Double = b * factor
        Return Math.Abs(aa - bb) < EPSILON
    End Function

    ''' <summary>
    ''' (高精度)浮動小数点が同じか？を返す
    ''' </summary>
    ''' <param name="a">値A</param>
    ''' <param name="b">値B</param>
    ''' <returns>同じなら、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEqualIfEpsilonStrictly(ByVal a As Double, ByVal b As Double) As Boolean
        Return Math.Abs(a - b) < Double.Epsilon
    End Function

    ''' <summary>
    ''' 小数点以下桁数を返す
    ''' </summary>
    ''' <param name="d">値</param>
    ''' <returns>少数点以下桁数</returns>
    ''' <remarks></remarks>
    Public Shared Function GetDecimalSize(ByVal d As Double) As Integer
        Return PerformGetDecimalSize(d)
    End Function
    ''' <summary>
    ''' 小数点以下桁数を返す
    ''' </summary>
    ''' <param name="d">値</param>
    ''' <returns>少数点以下桁数</returns>
    ''' <remarks></remarks>
    Public Shared Function GetDecimalSize(ByVal d As Single) As Integer
        Return PerformGetDecimalSize(d)
    End Function
    ''' <summary>
    ''' 小数点以下桁数を返す
    ''' </summary>
    ''' <param name="d">値</param>
    ''' <returns>少数点以下桁数</returns>
    ''' <remarks></remarks>
    Public Shared Function GetDecimalSize(ByVal d As Decimal) As Integer
        Return PerformGetDecimalSize(StringUtil.TrimEndDecimalZero(d.ToString))
    End Function
    ''' <summary>
    ''' 小数点以下桁数を返す
    ''' </summary>
    ''' <param name="d">値</param>
    ''' <returns>少数点以下桁数</returns>
    ''' <remarks></remarks>
    Private Shared Function PerformGetDecimalSize(ByVal d As Object) As Integer
        Const PERIOD As Char = "."c
        Dim strD As String = d.ToString
        Dim index As Integer = strD.IndexOf(PERIOD)
        If index < 0 Then
            Return 0
        End If
        Return strD.Length - index - 1
    End Function

    ''' <summary>
    ''' 値をEnum値に変換する
    ''' </summary>
    ''' <typeparam name="T">Enum型</typeparam>
    ''' <param name="value">値</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvValueToEnum(Of T As Structure)(ByVal value As Integer) As T
        Return EnumUtil.ParseByValue(Of T)(value)
    End Function

    ''' <summary>
    ''' 値名をEnum値に変換する
    ''' </summary>
    ''' <typeparam name="T">Enum型</typeparam>
    ''' <param name="valueName">値名</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvNameToEnum(Of T As Structure)(ByVal valueName As String) As T
        Return EnumUtil.ParseByName(Of T)(valueName)
    End Function

    ''' <summary>
    ''' Double値をDecimal値に変換する
    ''' </summary>
    ''' <param name="doubleVal">Double値</param>
    ''' <returns>Decimal値</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvDblToDec(ByVal doubleVal As Double) As Decimal
        Return Decimal.op_Explicit(doubleVal)
    End Function

    ''' <summary>
    ''' べき乗（x^y）を計算する ※Decimal型
    ''' </summary>
    ''' <param name="x">値</param>
    ''' <param name="y">指数</param>
    ''' <returns>べき乗の結果</returns>
    ''' <remarks></remarks>
    Public Shared Function Power(ByVal x As Decimal, ByVal y As Integer) As Decimal
        Dim result As Decimal = 1D
        If 0 < y Then
            For i As Integer = 0 To y - 1
                result *= x
            Next
        ElseIf y < 0 Then
            For i As Integer = 0 To (y * -1) - 1
                result = Decimal.Divide(result, x)
            Next
        End If
        Return result
    End Function

    ''' <summary>引数なしSubメソッドのdelegate</summary>
    ''' <remarks>System.Windows.Formsを参照していないプロジェクトに必要</remarks>
    Public Delegate Sub MethodInvoker()

    ''' <summary>
    ''' Functionとして戻り値なしメソッドを実行する
    ''' </summary>
    ''' <param name="method">戻り値なしメソッド</param>
    ''' <returns>ダミー値（意味をなさない）</returns>
    ''' <remarks></remarks>
    Public Shared Function RunSubAsFunc(ByVal method As MethodInvoker) As Object
        Return RunMethodAsFunc(New MethodInvoker(AddressOf method.Invoke))
    End Function

    ''' <summary>
    ''' Functionとして戻り値なしメソッドを実行する
    ''' </summary>
    ''' <param name="method">戻り値なしメソッド</param>
    ''' <returns>ダミー値（意味をなさない）</returns>
    ''' <remarks></remarks>
    Public Shared Function RunMethodAsFunc(ByVal method As System.Delegate, ByVal ParamArray args As Object()) As Object
        Const DUMMY As Integer = 0
        method.DynamicInvoke(args)
        Return DUMMY
    End Function

    ''' <summary>
    ''' シリアライズ可能なオブジェクトのコピーを生成する
    ''' </summary>
    ''' <typeparam name="T">対象オブジェクトの型</typeparam>
    ''' <param name="obj">対象オブジェクト</param>
    ''' <returns>対象オブジェクトのコピー</returns>
    ''' <remarks></remarks>
    Public Shared Function CopyForSerializable(Of T)(ByVal obj As T) As T
        Dim result As T
        Dim formatter As New Runtime.Serialization.Formatters.Binary.BinaryFormatter

        Using memoryStream As New System.IO.MemoryStream
            formatter.Serialize(memoryStream, obj)
            memoryStream.Position = 0
            result = DirectCast(formatter.Deserialize(memoryStream), T)
        End Using

        Return result
    End Function

    ''' <summary>
    ''' Null値でなければ処理する
    ''' </summary>
    ''' <param name="argIfNull">nullかもしれない値</param>
    ''' <param name="callback">nullでなければ行いたい処理</param>
    ''' <remarks></remarks>
    Public Shared Sub CallIf(argIfNull As Object, callback As Action(Of Object))
        ' Tが解決できないとき用のメソッド
        CallIf(Of Object)(argIfNull, callback)
    End Sub
    ''' <summary>
    ''' Null値でなければ処理する
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="argIfNull">nullかもしれない値</param>
    ''' <param name="callback">nullでなければ行いたい処理</param>
    ''' <remarks></remarks>
    Public Shared Sub CallIf(Of T)(argIfNull As T, callback As Action(Of T))
        If argIfNull Is Nothing Then
            Return
        End If
        callback(argIfNull)
    End Sub
    ''' <summary>
    ''' Null値でなければ処理する
    ''' </summary>
    ''' <param name="argIfNull">nullかもしれない値</param>
    ''' <param name="callback">nullでなければ行いたい処理</param>
    ''' <returns>処理の戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function CallIf(argIfNull As Object, callback As Func(Of Object, Object)) As Object
        ' Tが解決できないとき用のメソッド
        Return CallIf(Of Object, Object)(argIfNull, callback)
    End Function
    ''' <summary>
    ''' Null値でなければ処理する
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <typeparam name="TResult"></typeparam>
    ''' <param name="argIfNull">nullかもしれない値</param>
    ''' <param name="callback">nullでなければ行いたい処理</param>
    ''' <returns>処理の戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function CallIf(Of T, TResult)(argIfNull As T, callback As Func(Of T, TResult)) As TResult
        If argIfNull Is Nothing Then
            Return Nothing
        End If
        Return callback(argIfNull)
    End Function
    ''' <summary>
    ''' Null値でなければ処理する
    ''' </summary>
    ''' <typeparam name="TResult"></typeparam>
    ''' <param name="argIfNull">nullかもしれない値</param>
    ''' <param name="callback">nullでなければ行いたい処理</param>
    ''' <returns>処理の戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function CallIf(Of TResult)(argIfNull As Object, callback As Func(Of TResult)) As TResult
        If argIfNull Is Nothing Then
            Return Nothing
        End If
        Return callback.Invoke
    End Function

    ''' <summary>
    ''' Null値でなければ処理する
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="collectionIfEmpty">nullかもしれない値</param>
    ''' <param name="callback">nullでなければ行いたい処理</param>
    ''' <remarks></remarks>
    Public Shared Sub CallIfEmpty(Of T As ICollection)(collectionIfEmpty As T, callback As Action(Of T))
        CollectionUtil.CallIfEmpty(Of T)(collectionIfEmpty, callback)
    End Sub
    ''' <summary>
    ''' Null値でなければ処理する
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <typeparam name="TResult"></typeparam>
    ''' <param name="collectionIfEmpty">emptyかもしれないコレクション値</param>
    ''' <param name="callback">emptyでなければ行いたい処理</param>
    ''' <returns>処理の戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function CallIfEmpty(Of T, TResult)(collectionIfEmpty As ICollection(Of T), callback As Func(Of TResult)) As TResult
        Return CollectionUtil.CallIfEmpty(Of T, TResult)(collectionIfEmpty, callback)
    End Function
    ''' <summary>
    ''' Null値でなければ処理する
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <typeparam name="TResult"></typeparam>
    ''' <param name="collectionIfEmpty">emptyかもしれないコレクション値</param>
    ''' <param name="callback">emptyでなければ行いたい処理</param>
    ''' <returns>処理の戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function CallIfEmpty(Of T As ICollection, TResult)(collectionIfEmpty As T, callback As Func(Of T, TResult)) As TResult
        Return CollectionUtil.CallIfEmpty(Of T, TResult)(collectionIfEmpty, callback)
    End Function

    ''' <summary>
    ''' ワンデーパスワードを生成する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GenerateOneDayPassword() As String
        Return GenerateHashSHA256(DateTime.Now.ToString("yyyy/MM/dd"))
    End Function

    ''' <summary>
    ''' SHA256のハッシュ値を生成する
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GenerateHashSHA256(ByVal str As String) As String
        Return GenerateHashSHA256(System.Text.Encoding.UTF8.GetBytes(str))
    End Function
    ''' <summary>
    ''' SHA256のハッシュ値を生成する
    ''' </summary>
    ''' <param name="bytes">バイト文字列[]</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GenerateHashSHA256(ByVal bytes As Byte()) As String
        Using aSHA256 = SHA256.Create
            Dim computedHash As Byte() = aSHA256.ComputeHash(bytes)
            Return computedHash.Aggregate(New StringBuilder, Function(sb, sha) sb.Append(sha.ToString("x2"))).ToString
        End Using
    End Function

End Class
