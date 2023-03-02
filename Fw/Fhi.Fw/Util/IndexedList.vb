
Namespace Util
    ''' <summary>
    ''' indexと情報を紐付けたクラス
    ''' </summary>
    ''' <typeparam name="T">情報の型</typeparam>
    ''' <remarks></remarks>
    Public Class IndexedList(Of T)
        Implements IIndexedList(Of T)

        Private _Records As New Dictionary(Of Integer, T)
        Private _maxIndex As Integer?

        Private IsGenericValueType As Boolean
        ''' <summary>インスタンスを自動生成するか？を保持</summary>
        Public IsCreateGeneric As Boolean

        Private instanceType As Type

#Region "Nested classes..."
        ''' <summary>
        ''' 数値(Integer)の降順ソート
        ''' </summary>
        ''' <remarks></remarks>
        Public Class IntegerDescComparer : Implements IComparer(Of Integer)
            Public Function Compare(ByVal x As Integer, ByVal y As Integer) As Integer Implements IComparer(Of Integer).Compare
                Return y.CompareTo(x)
            End Function
        End Class
#End Region

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me.New(False)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="values">初期値</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal values As ICollection(Of T))
            Me.New(False, values)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="IsCreateGeneric">インスタンスを自動生成する場合、true</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal IsCreateGeneric As Boolean)
            Me.New(IsCreateGeneric, Nothing, Nothing)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="IsCreateGeneric">インスタンスを自動生成する場合、true</param>
        ''' <param name="values">初期値</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal IsCreateGeneric As Boolean, ByVal values As ICollection(Of T))
            Me.New(IsCreateGeneric, values, Nothing)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="IsCreateGeneric">インスタンスを自動生成する場合、true</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal IsCreateGeneric As Boolean, ByVal instanceType As Type)
            Me.New(IsCreateGeneric, Nothing, instanceType)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="IsCreateGeneric">インスタンスを自動生成する場合、true</param>
        ''' <param name="values">初期値</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal IsCreateGeneric As Boolean, ByVal values As ICollection(Of T), ByVal instanceType As Type)
            Me.IsCreateGeneric = IsCreateGeneric
            Dim genericType As Type = GetType(T)
            IsGenericValueType = genericType.IsValueType

            If instanceType Is Nothing Then
                Me.instanceType = GetType(T)
            Else
                Me.instanceType = instanceType
            End If

            If IsCreateGeneric Then
                If genericType Is GetType(String) Then
                    Throw New NotSupportedException("String型は未対応")
                End If
            End If

            If values IsNot Nothing Then
                AddRange(values)
            End If
        End Sub

        ''' <summary>
        ''' index位置に情報があるかを返す
        ''' </summary>
        ''' <param name="index">index</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Overridable Function HasValue(ByVal index As Integer) As Boolean Implements IIndexedList(Of T).HasValue
            Return _Records.ContainsKey(index)
        End Function

        ''' <summary>情報</summary>
        ''' <param name="index">index</param>
        ''' <value>情報</value>
        ''' <returns>情報</returns>
        ''' <remarks></remarks>
        Default Public Overridable Property Item(ByVal index As Integer) As T Implements IIndexedList(Of T).Item
            Get
                If Not _Records.ContainsKey(index) Then
                    If Not IsCreateGeneric Then
                        Return Nothing
                    End If
                    AssertEffectiveIndex(index)
                    _Records.Add(index, CType(Activator.CreateInstance(instanceType), T))
                End If
                Return _Records(index)
            End Get
            Set(ByVal value As T)
                If _Records.ContainsKey(index) Then
                    _Records.Remove(index)
                End If
                AssertEffectiveIndex(index)
                _Records.Add(index, value)
            End Set
        End Property

        Private Sub AssertEffectiveIndex(ByVal index As Integer)

            If index < 0 Then
                Throw New ArgumentOutOfRangeException("index", index, "添え字に負数はNG")
            End If
        End Sub

        ''' <summary>情報の一覧（添え字順）</summary>
        Public Overridable ReadOnly Property Items() As ICollection(Of T) Implements IIndexedList(Of T).Items
            Get
                Return Me.Indexes.Select(Function(index) _Records(index)).ToList
            End Get
        End Property

        ''' <summary>indexの一覧（昇順）</summary>
        Public Overridable ReadOnly Property Indexes() As ICollection(Of Integer) Implements IIndexedList(Of T).Indexes
            Get
                Dim keys As List(Of Integer) = _Records.Keys.ToList
                keys.Sort()
                Return keys
            End Get
        End Property

        ''' <summary>
        ''' 最大indexを返す
        ''' </summary>
        ''' <returns>最大index</returns>
        ''' <remarks></remarks>
        Public Overridable Function GetMaxIndex() As Integer? Implements IIndexedList(Of T).GetMaxIndex
            If _Records.Keys.Count = 0 Then
                Return Nothing
            End If
            Return _Records.Keys.Max
        End Function

        Private Function GetNewIndex() As Integer
            Return If(GetMaxIndex, -1) + 1
        End Function

        ''' <summary>
        ''' 情報を追加する
        ''' </summary>
        ''' <param name="value">情報</param>
        ''' <remarks></remarks>
        Public Overridable Sub Add(ByVal value As T) Implements IIndexedList(Of T).Add
            _Records.Add(GetNewIndex(), value)
        End Sub

        ''' <summary>
        ''' 情報を追加する
        ''' </summary>
        ''' <param name="values">情報</param>
        ''' <remarks></remarks>
        Public Overridable Sub AddRange(ByVal values As ICollection(Of T)) Implements IIndexedList(Of T).AddRange
            Dim index As Integer = GetNewIndex()
            For Each Item As T In values
                _Records.Add(index, Item)
                index += 1
            Next
        End Sub

        ''' <summary>
        ''' 情報をすべてクリアする
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub Clear() Implements IIndexedList(Of T).Clear
            _Records.Clear()
        End Sub

        ''' <summary>
        ''' 情報をクリアする。後続の情報はそのまま。
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub ClearAt(ByVal index As Integer) Implements IIndexedList(Of T).ClearAt
            If Not _Records.ContainsKey(index) Then
                Return
            End If
            _Records.Remove(index)
        End Sub

        ''' <summary>
        ''' 情報を挿入する
        ''' </summary>
        ''' <param name="index">挿入index</param>
        ''' <param name="count">挿入数</param>
        ''' <remarks></remarks>
        Public Overridable Sub InsertAt(ByVal index As Integer, Optional ByVal count As Integer = 1) Implements IIndexedList(Of T).InsertAt
            Dim newRecords As New Dictionary(Of Integer, T)
            For Each key As Integer In _Records.Keys
                newRecords.Add(If(key < index, key, key + count), _Records(key))
            Next
            _Records = newRecords
        End Sub

        ''' <summary>
        ''' 情報を除去する。後続の情報は前詰め。
        ''' </summary>
        ''' <param name="index">除去index</param>
        ''' <param name="count">除去数</param>
        ''' <remarks></remarks>
        Public Overridable Sub RemoveAt(ByVal index As Integer, Optional ByVal count As Integer = 1) Implements IIndexedList(Of T).RemoveAt
            Dim newRecords As New Dictionary(Of Integer, T)
            For Each key As Integer In _Records.Keys
                If index <= key AndAlso key < index + count Then
                    Continue For
                End If
                newRecords.Add(If(key < index, key, key - count), _Records(key))
            Next
            _Records = newRecords
        End Sub

        ''' <summary>
        ''' 情報を抽出する
        ''' </summary>
        ''' <param name="rowIndex">index</param>
        ''' <param name="count">取得数</param>
        ''' <remarks></remarks>
        Public Overridable Function Extract(ByVal rowIndex As Integer, ByVal count As Integer) As T() Implements IIndexedList(Of T).Extract
            Dim result As New List(Of T)
            For index As Integer = rowIndex To rowIndex + count - 1
                If HasValue(index) Then
                    result.Add(Item(index))
                Else
                    result.Add(Nothing)
                End If
            Next
            Return result.ToArray
        End Function

        ''' <summary>
        ''' 情報を差し替える
        ''' </summary>
        ''' <param name="rowIndex">index</param>
        ''' <param name="recordBags">差し替える情報</param>
        ''' <remarks></remarks>
        Public Overridable Sub Supersede(ByVal rowIndex As Integer, ByVal recordBags As T()) Implements IIndexedList(Of T).Supersede
            If Not TypeOf recordBags Is T() Then
                Throw New ArgumentException("Extractした値ではない", "recordBags")
            End If

            Dim destIndex As Integer = -1
            For Each bag As T In DirectCast(recordBags, T())
                destIndex += 1
                If bag Is Nothing Then
                    Continue For
                End If
                Item(rowIndex + destIndex) = bag
            Next

        End Sub

        ''' <summary>
        ''' 位置indexを返す
        ''' </summary>
        ''' <param name="value">探査する値</param>
        ''' <returns>見つからなければ -1</returns>
        ''' <remarks></remarks>
        Public Overridable Function IndexOf(ByVal value As T) As Integer Implements IIndexedList(Of T).IndexOf
            Dim valObj As Object = value
            Dim obj As Object
            For Each pair As KeyValuePair(Of Integer, T) In _Records
                obj = pair.Value
                If obj Is valObj Then
                    Return pair.Key
                End If
            Next
            Return -1
        End Function

        ''' <summary>
        ''' Indexの歯抜けやNothing値を除去して、前から整列しなおす
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SuppressGap() Implements IIndexedList(Of T).SuppressGap
            Dim indices As New List(Of Integer)(Indexes)
            Dim orderedList As New List(Of T)
            For Each index As Integer In indices
                If _Records(index) Is Nothing Then
                    Continue For
                End If
                orderedList.Add(_Records(index))
            Next
            Clear()
            AddRange(orderedList)
        End Sub

        ''' <summary>
        ''' List末尾のNothing値を除去して、サイズを縮小する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub TrimEnd() Implements IIndexedList(Of T).TrimEnd
            Dim indices As New List(Of Integer)(Indexes)
            indices.Sort(New IntegerDescComparer)
            For Each index As Integer In indices
                If _Records(index) IsNot Nothing Then
                    Exit For
                End If
                RemoveAt(index)
            Next
        End Sub

        ''' <summary>
        ''' IndexedList(T) を反復処理する列挙子を返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function IEnumerator_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return Me.GetEnumerator
        End Function

        ''' <summary>
        ''' IndexedList(T) を反復処理する列挙子を返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
            Return Me.Indexes.Select(Function(key) _Records(key)).GetEnumerator
        End Function

    End Class
End Namespace