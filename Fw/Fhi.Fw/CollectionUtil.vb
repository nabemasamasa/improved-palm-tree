Imports Fhi.Fw.Domain

''' <summary>
''' コレクション用のユーティリティクラス
''' </summary>
''' <remarks></remarks>
Public Class CollectionUtil
    ''' <summary>値の一意キーを作成する</summary>
    ''' <typeparam name="T">値の型</typeparam>
    ''' <param name="obj">値</param>
    ''' <returns>一意キー</returns>
    ''' <remarks></remarks>
    Public Delegate Function MakeKeyCallback(Of T)(ByVal obj As t) As Object

    ''' <summary>値の一意キーを作成する</summary>
    ''' <typeparam name="T">値の型</typeparam>
    ''' <param name="obj">値</param>
    ''' <returns>一意キー</returns>
    ''' <remarks></remarks>
    Public Delegate Function MakeKeyCallback(Of T, TResult)(ByVal obj As t) As TResult

    ''' <summary>
    ''' コレクション同士の差を取得する ※T がEqualsメソッドで一意判定可能であること
    ''' </summary>
    ''' <typeparam name="T">コレクションの型</typeparam>
    ''' <param name="minuendVos">被減値</param>
    ''' <param name="subtrahendVos">減値</param>
    ''' <returns>差[]</returns>
    ''' <remarks></remarks>
    Public Shared Function SubtractBy(Of T)(ByVal minuendVos As IEnumerable(Of T), ByVal subtrahendVos As IEnumerable(Of T)) As T()
        Return SubtractBy(Of T)(minuendVos, subtrahendVos, Function(obj As T) obj)
    End Function

    ''' <summary>
    ''' コレクション同士の差を取得する
    ''' </summary>
    ''' <typeparam name="T">コレクションの型</typeparam>
    ''' <param name="minuendVos">被減値</param>
    ''' <param name="subtrahendVos">減値</param>
    ''' <param name="aMakeKey">値の一意キーを作成するdelegate</param>
    ''' <returns>差[]</returns>
    ''' <remarks></remarks>
    Public Shared Function SubtractBy(Of T)(ByVal minuendVos As IEnumerable(Of T), ByVal subtrahendVos As IEnumerable(Of T), ByVal aMakeKey As MakeKeyCallback(Of T)) As T()
        Dim map As New Dictionary(Of Object, String)
        For Each vo As T In subtrahendVos
            Dim key As Object = aMakeKey(vo)
            If map.ContainsKey(key) Then
                Continue For
            End If
            map.Add(key, Nothing)
        Next

        Return (From vo In minuendVos Let key = aMakeKey(vo) Where Not map.ContainsKey(key) Select vo).ToArray()
    End Function

    ''' <summary>
    ''' コレクション同士をマージする（同一インスタンスを除外）
    ''' </summary>
    ''' <param name="a">コレクションA</param>
    ''' <param name="b">コレクションB</param>
    ''' <returns>マージした結果配列（コレクション）</returns>
    ''' <remarks></remarks>
    Public Shared Function MergeBy(Of T)(ByVal a As IEnumerable(Of T), ByVal b As IEnumerable(Of T)) As T()
        Return MergeBy(a, b, Function(obj As T) obj)
    End Function

    ''' <summary>
    ''' コレクション同士をマージする
    ''' </summary>
    ''' <param name="a">コレクションA</param>
    ''' <param name="b">コレクションB</param>
    ''' <param name="aMakeKey">値の一意キーを作成するdelegate</param>
    ''' <returns>マージした結果配列（コレクション）</returns>
    ''' <remarks></remarks>
    Public Shared Function MergeBy(Of T)(ByVal a As IEnumerable(Of T), ByVal b As IEnumerable(Of T), ByVal aMakeKey As MakeKeyCallback(Of T)) As T()
        Dim result As New List(Of T)
        Dim existsKeys As New List(Of Object)
        For Each obj As T In Combine(a, b)
            Dim key As Object = aMakeKey(obj)
            If existsKeys.Contains(key) Then
                Continue For
            End If
            existsKeys.Add(key)
            result.Add(obj)
        Next
        Return result.ToArray
    End Function

    ''' <summary>
    ''' リスト（配列）を結合して返す
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="lists">結合するリスト（配列）</param>
    ''' <returns>結合したリスト</returns>
    ''' <remarks></remarks>
    Public Shared Function Combine(Of T)(ByVal ParamArray lists As IEnumerable(Of T)()) As T()
        Dim result As New List(Of T)
        For Each collection As IEnumerable(Of T) In lists
            result.AddRange(collection)
        Next
        Return result.ToArray
    End Function

    ''' <summary>
    ''' 最大数を設定してコレクションを分割する
    ''' </summary>
    ''' <typeparam name="T">コレクションの内部の型</typeparam>
    ''' <param name="collection">コレクション</param>
    ''' <param name="maxCollectionSize">最大数</param>
    ''' <returns>最大数で分割されたコレクションの配列</returns>
    ''' <remarks></remarks>
    Public Shared Function Split(Of T)(ByVal collection As IEnumerable(Of T), ByVal maxCollectionSize As Integer) As T()()
        Dim result As New List(Of T())
        Dim collectionArr As T() = collection.ToArray
        If IsEmpty(collectionArr) Then
            Return result.ToArray
        End If

        Dim splitted As New List(Of T)
        Dim count As Integer = 0
        For Each value As T In collectionArr
            If maxCollectionSize <= count Then
                result.Add(splitted.ToArray)
                splitted.Clear()
                count = 0
            End If
            splitted.Add(value)
            count += 1
        Next
        result.Add(splitted.ToArray)
        Return result.ToArray
    End Function

    ''' <summary>
    ''' 配列から、Listを作成する
    ''' </summary>
    ''' <typeparam name="T">型</typeparam>
    ''' <param name="values">値</param>
    ''' <returns>List値</returns>
    ''' <remarks></remarks>
    Public Shared Function NewList(Of T)(ByVal ParamArray values As T()) As List(Of T)
        Return New List(Of T)(values)
    End Function

    ''' <summary>
    ''' XxxxCollection をList型に変換する
    ''' </summary>
    ''' <typeparam name="T">XxxxCollectionの要素の型</typeparam>
    ''' <param name="enumerable">XxxxCollection</param>
    ''' <returns>List型</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvEnumerableToList(Of T)(ByVal enumerable As IEnumerable) As List(Of T)

        Dim result As New List(Of T)
        For Each obj As Object In enumerable
            If obj IsNot Nothing AndAlso Not TypeOf obj Is T Then
                Throw New InvalidProgramException(String.Format("{0} の中身が {1} ではない. {2}", enumerable.GetType.Name, GetType(T).Name, obj.GetType.Name))
            End If
            result.Add(DirectCast(obj, T))
        Next
        Return result
    End Function

    ''' <summary>
    ''' 値を持つコレクションか？を返す
    ''' </summary>
    ''' <param name="collection">コレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEmpty(ByVal collection As IEnumerable) As Boolean
        If collection Is Nothing Then
            Return False
        End If
        Return collection.Cast(Of Object)().Any()
    End Function

    ''' <summary>
    ''' 値を持つコレクションか？を返す
    ''' </summary>
    ''' <param name="collection">コレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEmpty(ByVal collection As ICollectionObject) As Boolean
        If collection Is Nothing Then
            Return False
        End If
        Return 0 < collection.Count
    End Function

    ''' <summary>
    ''' 値を持たないコレクションか？を返す
    ''' </summary>
    ''' <param name="collection">コレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmpty(ByVal collection As IEnumerable) As Boolean
        Return Not IsNotEmpty(collection)
    End Function

    ''' <summary>
    ''' 値を持たないコレクションか？を返す
    ''' </summary>
    ''' <param name="collection">コレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEmpty(ByVal collection As ICollectionObject) As Boolean
        Return Not IsNotEmpty(collection)
    End Function

    ''' <summary>
    ''' 含まれているかを返す
    ''' </summary>
    ''' <param name="array">母体（複数）</param>
    ''' <param name="src">判定値</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function Contains(ByVal array As Object(), ByVal src As Object) As Boolean
        Return array.Any(Function(value) EzUtil.IsEqualIfNull(value, src))
    End Function

    ''' <summary>
    ''' PVOかを無視して、含まれているかを返す
    ''' </summary>
    ''' <param name="array">母体（複数）</param>
    ''' <param name="src">判定値</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function ContainsIgnorePvo(ByVal array As Object(), ByVal src As Object) As Boolean
        Return array.Any(Function(value) EzUtil.IsEqualIgnorePvoIfNull(value, src))
    End Function

    ''' <summary>
    ''' 含まれているかを返す
    ''' </summary>
    ''' <param name="array">母体（複数）</param>
    ''' <param name="src">判定値</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function Contains(Of T As Structure)(ByVal array As T(), ByVal src As Object) As Boolean
        Return array.Any(Function(value) EzUtil.IsEqualIfNull(value, src))
    End Function

    ''' <summary>
    ''' Empty値を含んでいるかを返す
    ''' </summary>
    ''' <param name="values">値[]</param>
    ''' <returns>含んでいる場合、true</returns>
    ''' <remarks></remarks>
    Public Shared Function ContainsEmpty(ByVal ParamArray values() As Object) As Boolean
        Return values.Any(Function(obj) StringUtil.IsEmpty(obj))
    End Function

    ''' <summary>
    ''' Nullを含んでいるかを返す
    ''' </summary>
    ''' <param name="values">値[]</param>
    ''' <returns>含んでいる場合、true</returns>
    ''' <remarks></remarks>
    Public Shared Function ContainsNull(ByVal ParamArray values As Object()) As Boolean
        Return values.Any(Function(value) value Is Nothing)
    End Function

    ''' <summary>
    ''' 特定のキーごとの一覧にして返す
    ''' </summary>
    ''' <typeparam name="T">コレクションの型</typeparam>
    ''' <param name="vos">コレクション値[]</param>
    ''' <param name="aMakeKey">特定のキーを作成 ※ラムダ式</param>
    ''' <returns>キーごとの一覧</returns>
    ''' <remarks></remarks>
    Public Shared Function MakeVosByKey(Of T, TResult)(ByVal vos As IEnumerable(Of T), ByVal aMakeKey As MakeKeyCallback(Of T, TResult)) As Dictionary(Of TResult, List(Of T))
        Dim result As New Dictionary(Of TResult, List(Of T))
        For Each vo As T In vos
            Dim key As TResult = aMakeKey(vo)
            If Not result.ContainsKey(key) Then
                result.Add(key, New List(Of T))
            End If
            result(key).Add(vo)
        Next
        Return result
    End Function

    ''' <summary>
    ''' 特定のキーごとにして返す ※同一キーの値は、先勝ち
    ''' </summary>
    ''' <typeparam name="T">コレクションの型</typeparam>
    ''' <param name="vos">コレクション値[]</param>
    ''' <param name="aMakeKey">特定のキーを作成 ※ラムダ式</param>
    ''' <returns>キーごとの値</returns>
    ''' <remarks></remarks>
    Public Shared Function MakeVoByKey(Of T, TResult)(ByVal vos As IEnumerable(Of T), ByVal aMakeKey As MakeKeyCallback(Of T, TResult)) As Dictionary(Of TResult, T)
        Dim result As New Dictionary(Of TResult, T)
        For Each vo As T In vos
            Dim key As TResult = aMakeKey(vo)
            If result.ContainsKey(key) Then
                Continue For
            End If
            result.Add(key, vo)
        Next
        Return result
    End Function

    ''' <summary>
    ''' Null値を考慮して同値かを返す
    ''' </summary>
    ''' <param name="a">値a</param>
    ''' <param name="b">値b</param>
    ''' <returns>Null同士、または同値ならtrue</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEqualIfNull(Of T1, T2)(ByVal a As IEnumerable(Of T1), ByVal b As IEnumerable(Of T2)) As Boolean
        If a Is Nothing AndAlso b Is Nothing Then
            Return True
        ElseIf a Is Nothing OrElse b Is Nothing Then
            Return False
        End If
        Dim a2 As List(Of T1) = ConvEnumerableToList(Of T1)(a)
        Dim b2 As List(Of T2) = ConvEnumerableToList(Of T2)(b)
        If a2.Count <> b2.Count Then
            Return False
        End If
        For i As Integer = 0 To a2.Count - 1
            If EzUtil.IsEqualIfNull(a2(i), b2(i)) Then
                Continue For
            End If
            Return False
        Next
        Return True
    End Function

    ''' <summary>
    ''' 値(null含む)が違うかを返す
    ''' </summary>
    ''' <param name="x">値x</param>
    ''' <param name="y">値y</param>
    ''' <returns>違う場合、true</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotEqualIfNull(Of T1, T2)(ByVal x As IEnumerable(Of T1), ByVal y As IEnumerable(Of T2)) As Boolean
        Return Not IsEqualIfNull(x, y)
    End Function

    ''' <summary>
    ''' 配列にして返す
    ''' </summary>
    ''' <typeparam name="T">型</typeparam>
    ''' <param name="collection">コレクション値</param>
    ''' <returns>配列</returns>
    ''' <remarks></remarks>
    Public Shared Function ToArray(Of T)(collection As IEnumerable(Of T)) As T()
        If collection Is Nothing Then
            Return Nothing
        End If
        Return collection.ToArray
    End Function

    ''' <summary>
    ''' List型にして返す
    ''' </summary>
    ''' <typeparam name="T">型</typeparam>
    ''' <param name="collection">コレクション値</param>
    ''' <returns>List型の値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToList(Of T)(collection As IEnumerable(Of T)) As List(Of T)
        If collection Is Nothing Then
            Return Nothing
        End If
        Return collection.ToList
    End Function

    ''' <summary>
    ''' Null値でなれけば処理する
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="collectionIfEmpty">nullかもしれない値</param>
    ''' <param name="callback">nullでなければ行いたい処理</param>
    ''' <remarks></remarks>
    Public Shared Sub CallIfEmpty(Of T As ICollection)(collectionIfEmpty As T, callback As Action(Of T))
        If CollectionUtil.IsEmpty(collectionIfEmpty) Then
            Return
        End If
        callback(collectionIfEmpty)
    End Sub
    ''' <summary>
    ''' Null値でなれけば処理する
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <typeparam name="TResult"></typeparam>
    ''' <param name="collectionIfEmpty">emptyかもしれないコレクション値</param>
    ''' <param name="callback">emptyでなければ行いたい処理</param>
    ''' <returns>処理の戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function CallIfEmpty(Of T, TResult)(collectionIfEmpty As ICollection(Of T), callback As Func(Of TResult)) As TResult
        If CollectionUtil.IsEmpty(collectionIfEmpty) Then
            Return Nothing
        End If
        Return callback()
    End Function
    ''' <summary>
    ''' Null値でなれけば処理する
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <typeparam name="TResult"></typeparam>
    ''' <param name="collectionIfEmpty">emptyかもしれないコレクション値</param>
    ''' <param name="callback">emptyでなければ行いたい処理</param>
    ''' <returns>処理の戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function CallIfEmpty(Of T As ICollection, TResult)(collectionIfEmpty As T, callback As Func(Of T, TResult)) As TResult
        If CollectionUtil.IsEmpty(collectionIfEmpty) Then
            Return Nothing
        End If
        Return callback(collectionIfEmpty)
    End Function

    ''' <summary>
    ''' 2次元配列[,]に変換する
    ''' </summary>
    ''' <param name="data">値[][]</param>
    ''' <returns>2次元配列[,]</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvToTwoDimensionalArray(Of T)(data As IEnumerable(Of T)) As Object(,)
        Return ConvToTwoDimensionalArray(DirectCast(If(TypeOf data Is String OrElse TypeOf data Is Array, data, data.ToArray), Object))
    End Function

    ''' <summary>
    ''' 2次元配列[,]に変換する
    ''' </summary>
    ''' <param name="data">値[][]</param>
    ''' <returns>2次元配列[,]</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvToTwoDimensionalArray(data As Object) As Object(,)
        If data Is Nothing Then
            Dim results(,) As Object = {{}}
            Return results
        ElseIf TypeOf data Is Array Then
            Dim valArray As Array = DirectCast(data, Array)
            If valArray.Rank = 2 Then
                Return DirectCast(data, Object(,))
            ElseIf valArray.Rank = 1 Then
                If data.GetType.Name.EndsWith("[][]") Then
                    Dim rowSize As Integer = valArray.Length
                    Dim columnSize As Integer = Enumerable.Range(0, rowSize).Max(Function(r) DirectCast(valArray.GetValue(r), Array).Length)
                    Dim results(rowSize - 1, columnSize - 1) As Object
                    For row As Integer = 0 To rowSize - 1
                        Dim rowArray As Array = DirectCast(valArray.GetValue(row), Array)
                        For column As Integer = 0 To rowArray.Length - 1
                            results(row, column) = rowArray.GetValue(column)
                        Next
                    Next
                    Return results
                Else
                    Dim columnSize As Integer = valArray.Length
                    Dim results(0, columnSize - 1) As Object
                    For column As Integer = 0 To valArray.Length - 1
                        results(0, column) = valArray.GetValue(column)
                    Next
                    Return results
                End If
            End If
        ElseIf TypeOf data Is String OrElse data.GetType.IsPrimitive Then
            Dim results(,) As Object = {{data}}
            Return results
        End If
        Throw New ArgumentException("配列 data() か 2段階配列 data()() であるべき")
    End Function

    ''' <summary>
    ''' 値を二段階配列にして返す
    ''' </summary>
    ''' <param name="data">値[][]</param>
    ''' <returns>二段階配列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvToTwoJaggedArray(Of T)(data As IEnumerable(Of T)) As Object()()
        Return ConvToTwoJaggedArray(DirectCast(If(TypeOf data Is String OrElse TypeOf data Is Array, data, data.ToArray), Object))
    End Function

    ''' <summary>
    ''' 値を二段階配列にして返す
    ''' </summary>
    ''' <param name="data">値(2次元配列値[,]可)</param>
    ''' <returns>二段階配列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvToTwoJaggedArray(data As Object) As Object()()
        Return PerformConvToTwoJaggedArray(Of Object)(data)
    End Function

    ''' <summary>
    ''' 値を二段階配列にして返す
    ''' </summary>
    ''' <param name="data">値(2次元配列値[,]可)</param>
    ''' <returns>二段階配列</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvToTwoJaggedArrayAsString(data As Object) As String()()
        Return PerformConvToTwoJaggedArray(Of String)(data)
    End Function

    ''' <summary>
    ''' 値を二段階配列にして返す
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="data">値(2次元配列値[,]可)</param>
    ''' <returns>二段階配列</returns>
    ''' <remarks></remarks>
    Private Shared Function PerformConvToTwoJaggedArray(Of T)(data As Object) As T()()
        Dim result As New List(Of T())
        If Data Is Nothing Then
            Return result.ToArray
        End If
        Dim convertInvoker As Func(Of Object, T) = If(GetType(T) Is GetType(String), _
                                                      Function(val As Object) DirectCast(DirectCast(StringUtil.ToString(val), Object), T), _
                                                      Function(val As Object) DirectCast(val, T))
        If TypeOf Data Is Array Then
            Dim matrixArray As Array = DirectCast(Data, Array)
            If matrixArray.Rank = 1 Then
                Dim anArray As Array = matrixArray
                If Not anArray.GetType.Name.EndsWith("[][]") Then
                    anArray = New Object() {anArray}
                End If
                For rowIndex As Integer = anArray.GetLowerBound(0) To anArray.GetLength(0) - 1 + anArray.GetLowerBound(0)
                    Dim rowData As Array = DirectCast(anArray.GetValue(rowIndex), Array)
                    result.Add(Enumerable.Range(rowData.GetLowerBound(0), rowData.GetLength(0)).Select(Function(i) convertInvoker.Invoke(rowData.GetValue(i))).ToArray())
                Next
            ElseIf matrixArray.Rank = 2 Then
                For rowIndex As Integer = matrixArray.GetLowerBound(0) To matrixArray.GetLength(0) - 1 + matrixArray.GetLowerBound(0)
                    Dim row As New List(Of T)
                    For columnIndex As Integer = matrixArray.GetLowerBound(1) To matrixArray.GetLength(1) - 1 + matrixArray.GetLowerBound(1)
                        row.Add(convertInvoker.Invoke(matrixArray.GetValue(rowIndex, columnIndex)))
                    Next
                    result.Add(row.ToArray)
                Next
            Else
                Throw New InvalidProgramException
            End If
        ElseIf TypeOf Data Is String OrElse Data.GetType.IsPrimitive Then
            result.Add(New T() {convertInvoker.Invoke(Data)})
        Else
            Throw New NotSupportedException(String.Format("{0} は未対応. 要PG修正", Data.GetType.Name))
        End If
        Return result.ToArray
    End Function

    ''' <summary>
    ''' 「最近使った値」を利用順に更新する（大文字小文字同一視）
    ''' </summary>
    ''' <param name="newValue">新しい値</param>
    ''' <param name="recentValues">元の値[]</param>
    ''' <param name="recentMaxCount">最大件数</param>
    ''' <returns>更新後の値[]</returns>
    ''' <remarks></remarks>
    Public Shared Function UpdateRecentValues(newValue As String, recentValues As String(), recentMaxCount As Integer) As String()
        If IsEmpty(recentValues) Then
            Return {newValue}
        End If
        Dim values As List(Of String) = recentValues.ToList
        If values.Contains(newValue, StringComparer.OrdinalIgnoreCase) Then
            Dim first As String = values.First(Function(name) String.Compare(name, newValue, True) = 0)
            values.Remove(first)
        End If
        values.Insert(0, newValue)
        While recentMaxCount < values.Count
            values.RemoveAt(values.Count - 1)
        End While
        Return values.ToArray
    End Function

End Class
