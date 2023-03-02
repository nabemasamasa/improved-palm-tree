Imports System.Reflection

Namespace Domain
    ''' <summary>
    ''' 不変のコレクションオブジェクトを担うクラス
    ''' </summary>
    ''' <typeparam name="T">要素</typeparam>
    ''' <remarks></remarks>
    Public Class CollectionObject(Of T) : Implements ICollectionObject

        ''' <summary>要素を持つ内部List</summary>
        Protected ReadOnly InternalList As List(Of T)

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me.New(DirectCast(Nothing, IEnumerable(Of T)))
        End Sub
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="src"></param>
        ''' <remarks>※SubClassは当コンストラクタ定義を再実装しないと#Clone()が動作しない</remarks>
        Public Sub New(src As CollectionObject(Of T))
            Me.New(If(src Is Nothing, Nothing, src.InternalList))
        End Sub
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="initialList">初期値となる要素[]</param>
        ''' <remarks></remarks>
        Public Sub New(initialList As IEnumerable(Of T))
            Me.InternalList = New List(Of T)
            If initialList IsNot Nothing Then
                Me.InternalList.AddRange(initialList)
            End If
        End Sub

        ''' <summary>
        ''' 複製する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overridable Function Clone() As CollectionObject(Of T)
            Return Clone(Of CollectionObject(Of T))()
        End Function
        ''' <summary>
        ''' 複製する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overridable Function Clone(Of TResult As CollectionObject(Of T))() As TResult
            Dim parameters As Object() = New Object() {Me}
            Dim types As Type() = New Type() {Me.GetType}
            Dim constructor As ConstructorInfo = Me.GetType.GetConstructor(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance,
                                                                           Nothing, types, Nothing)
            If constructor Is Nothing Then
                Throw New InvalidProgramException(String.Format("型 {0} に.ctor({1}) が無いと動作しない",
                                                                Me.GetType.FullName, Join(types.Select(Function(t) t.Name).ToArray, ",")))
            End If
            Return DirectCast(constructor.Invoke(parameters), TResult)
        End Function

        ''' <summary>
        ''' 各要素に対してActionを実行する
        ''' </summary>
        ''' <param name="action">Action</param>
        ''' <remarks></remarks>
        <Obsolete("自身にメソッド化するべき")> Public Sub ForEach(action As Action(Of T))
            EzUtil.AssertParameterIsNotNull(action, "action")
            For Each item As T In InternalList
                action.Invoke(item)
            Next
        End Sub

        ''' <summary>
        ''' 各要素に対してActionを実行する
        ''' </summary>
        ''' <param name="action">Action</param>
        ''' <remarks></remarks>
        <Obsolete("自身にメソッド化するべき")> Public Sub ForEach(action As Action(Of T, Integer))
            EzUtil.AssertParameterIsNotNull(action, "action")
            Dim i As Integer = 0
            For Each item As T In InternalList
                action.Invoke(item, i)
                i += 1
            Next
        End Sub

        ''' <summary>要素の数</summary>
        Public ReadOnly Property Count() As Integer Implements ICollectionObject.Count
            Get
                Return InternalList.Count
            End Get
        End Property

        ''' <summary>要素（値オブジェクト）</summary>
        Private ReadOnly Property ICollectionObject_Item(index As Integer) As Object Implements ICollectionObject.Item
            Get
                Return Me.Item(index)
            End Get
        End Property
        ''' <summary>要素（値オブジェクト）</summary>
        Default Public ReadOnly Property Item(index As Integer) As T
            Get
                Return InternalList.Item(index)
            End Get
        End Property

        ''' <summary>
        ''' 要素を更新する
        ''' </summary>
        ''' <param name="index">位置index</param>
        ''' <param name="updateCallback">指定位置要素更新Callback</param>
        ''' <returns>更新後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function UpdateItem(index As Integer, updateCallback As Func(Of T, T)) As CollectionObject(Of T)
            Return UpdateItem(index, updateCallback.Invoke(InternalList(index)))
        End Function
        ''' <summary>
        ''' 要素を更新する
        ''' </summary>
        ''' <param name="index">位置index</param>
        ''' <param name="updateCallback">指定位置要素更新Callback</param>
        ''' <returns>更新後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function UpdateItem(Of TResult As CollectionObject(Of T))(index As Integer, updateCallback As Func(Of T, T)) As TResult
            Return UpdateItem(Of TResult)(index, updateCallback.Invoke(InternalList(index)))
        End Function

        ''' <summary>
        ''' 要素を更新する
        ''' </summary>
        ''' <param name="index">位置index</param>
        ''' <param name="item">要素</param>
        ''' <returns>更新後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function UpdateItem(index As Integer, item As T) As CollectionObject(Of T)
            Return UpdateItem(Of CollectionObject(Of T))(index, item)
        End Function
        ''' <summary>
        ''' 要素を更新する
        ''' </summary>
        ''' <param name="index">位置index</param>
        ''' <param name="item">要素</param>
        ''' <returns>更新後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function UpdateItem(Of TResult As CollectionObject(Of T))(index As Integer, item As T) As TResult
            Dim result As CollectionObject(Of T) = Clone()
            result.InternalList(index) = item
            Return DirectCast(result, TResult)
        End Function

        ''' <summary>
        ''' 要素を追加する
        ''' </summary>
        ''' <param name="item">要素</param>
        ''' <returns>追加後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Add(item As T) As CollectionObject(Of T)
            Return Add(Of CollectionObject(Of T))(item)
        End Function
        ''' <summary>
        ''' 要素を追加する
        ''' </summary>
        ''' <param name="item">要素</param>
        ''' <returns>追加後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Add(Of TResult As CollectionObject(Of T))(item As T) As TResult
            Dim result As CollectionObject(Of T) = Clone()
            result.InternalList.Add(item)
            Return DirectCast(result, TResult)
        End Function

        ''' <summary>
        ''' 要素を追加する
        ''' </summary>
        ''' <param name="items">要素[]</param>
        ''' <returns>追加後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function AddRange(items As IEnumerable(Of T)) As CollectionObject(Of T)
            Return AddRange(Of CollectionObject(Of T))(items)
        End Function
        ''' <summary>
        ''' 要素を追加する
        ''' </summary>
        ''' <param name="items">要素[]</param>
        ''' <returns>追加後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function AddRange(Of TResult As CollectionObject(Of T))(items As IEnumerable(Of T)) As TResult
            Dim result As CollectionObject(Of T) = Clone()
            result.InternalList.AddRange(items)
            Return DirectCast(result, TResult)
        End Function

        ''' <summary>
        ''' 要素が含まれるか？を返す
        ''' </summary>
        ''' <param name="item">要素</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Contains(ByVal item As T) As Boolean
            Return InternalList.Contains(item)
        End Function

        ''' <summary>
        ''' 要素のある位置indexを返す
        ''' </summary>
        ''' <param name="item">要素</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IndexOf(ByVal item As T) As Integer
            Return InternalList.IndexOf(item)
        End Function

        ''' <summary>
        ''' 要素を挿入する
        ''' </summary>
        ''' <param name="index">挿入位置idnex</param>
        ''' <param name="item">要素</param>
        ''' <returns>挿入後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Insert(ByVal index As Integer, ByVal item As T) As CollectionObject(Of T)
            Return Insert(Of CollectionObject(Of T))(index, item)
        End Function
        ''' <summary>
        ''' 要素を挿入する
        ''' </summary>
        ''' <param name="index">挿入位置idnex</param>
        ''' <param name="item">要素</param>
        ''' <returns>挿入後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Insert(Of TResult As CollectionObject(Of T))(ByVal index As Integer, ByVal item As T) As TResult
            Dim result As CollectionObject(Of T) = Clone()
            result.InternalList.Insert(index, item)
            Return DirectCast(result, TResult)
        End Function

        ''' <summary>
        ''' 要素を除去する
        ''' </summary>
        ''' <param name="item">要素</param>
        ''' <returns>除去後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Remove(ByVal item As T) As CollectionObject(Of T)
            Return Remove(Of CollectionObject(Of T))(item)
        End Function
        ''' <summary>
        ''' 要素を除去する
        ''' </summary>
        ''' <param name="item">要素</param>
        ''' <returns>除去後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Remove(Of TResult As CollectionObject(Of T))(ByVal item As T) As TResult
            Dim result As CollectionObject(Of T) = Clone()
            result.InternalList.Remove(item)
            Return DirectCast(result, TResult)
        End Function

        ''' <summary>
        ''' 位置indexの要素を除去する
        ''' </summary>
        ''' <param name="index">除去する要素の位置index</param>
        ''' <returns>除去後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function RemoveAt(ByVal index As Integer) As CollectionObject(Of T)
            Return RemoveAt(Of CollectionObject(Of T))(index)
        End Function
        ''' <summary>
        ''' 位置indexの要素を除去する
        ''' </summary>
        ''' <param name="index">除去する要素の位置index</param>
        ''' <returns>除去後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function RemoveAt(Of TResult As CollectionObject(Of T))(ByVal index As Integer) As TResult
            Dim result As CollectionObject(Of T) = Clone()
            result.InternalList.RemoveAt(index)
            Return DirectCast(result, TResult)
        End Function

        ''' <summary>
        ''' 並び替えする
        ''' </summary>
        ''' <param name="comparer">並び替え条件</param>
        ''' <returns>並び替え後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Sort(ByVal comparer As IComparer(Of T)) As CollectionObject(Of T)
            Return Sort(Of CollectionObject(Of T))(comparer)
        End Function
        ''' <summary>
        ''' 並び替えする
        ''' </summary>
        ''' <param name="comparer">並び替え条件</param>
        ''' <returns>並び替え後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Sort(Of TResult As CollectionObject(Of T))(ByVal comparer As IComparer(Of T)) As TResult
            Dim result As CollectionObject(Of T) = Clone()
            result.InternalList.Sort(comparer)
            Return DirectCast(result, TResult)
        End Function

        ''' <summary>
        ''' 並び替えする
        ''' </summary>
        ''' <param name="comparison">並び替え条件</param>
        ''' <returns>並び替え後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Sort(ByVal comparison As Comparison(Of T)) As CollectionObject(Of T)
            Return Sort(Of CollectionObject(Of T))(comparison)
        End Function
        ''' <summary>
        ''' 並び替えする
        ''' </summary>
        ''' <param name="comparison">並び替え条件</param>
        ''' <returns>並び替え後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Sort(Of TResult As CollectionObject(Of T))(ByVal comparison As Comparison(Of T)) As TResult
            Dim result As CollectionObject(Of T) = Clone()
            result.InternalList.Sort(comparison)
            Return DirectCast(result, TResult)
        End Function

        ''' <summary>
        ''' 絞り込む
        ''' </summary>
        ''' <param name="predicate">絞り込み条件</param>
        ''' <returns>絞込み後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Where(ByVal predicate As Func(Of T, Boolean)) As CollectionObject(Of T)
            Return Where(Of CollectionObject(Of T))(predicate)
        End Function
        ''' <summary>
        ''' 絞り込む
        ''' </summary>
        ''' <param name="predicate">絞り込み条件</param>
        ''' <returns>絞込み後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Where(Of TResult As CollectionObject(Of T))(ByVal predicate As Func(Of T, Boolean)) As TResult
            Dim result As CollectionObject(Of T) = Clone()
            result.InternalList.Clear()
            result.InternalList.AddRange(Me.InternalList.Where(predicate))
            Return DirectCast(result, TResult)
        End Function

        ''' <summary>
        ''' 入れ替える
        ''' </summary>
        ''' <param name="x">入れ替えるindex1</param>
        ''' <param name="y">入れ替えるindex2</param>
        ''' <returns>入れ替え後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Exchange(x As Integer, y As Integer) As CollectionObject(Of T)
            Return Exchange(Of CollectionObject(Of T))(x, y)
        End Function
        ''' <summary>
        ''' 入れ替える
        ''' </summary>
        ''' <param name="x">入れ替えるindex1</param>
        ''' <param name="y">入れ替えるindex2</param>
        ''' <returns>入れ替え後のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Protected Overridable Function Exchange(Of TResult As CollectionObject(Of T))(x As Integer, y As Integer) As TResult
            Dim result As CollectionObject(Of T) = Clone()
            Dim itemX As T = result.InternalList(x)
            result.InternalList(x) = result.InternalList(y)
            result.InternalList(y) = itemX
            Return DirectCast(result, TResult)
        End Function

    End Class
End Namespace