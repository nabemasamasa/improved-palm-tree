Namespace Util
    Public Interface IIndexedList(Of T) : Inherits IEnumerable(Of T)

        ''' <summary>
        ''' index位置に情報があるかを返す
        ''' </summary>
        ''' <param name="index">index</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Function HasValue(ByVal index As Integer) As Boolean

        ''' <summary>情報</summary>
        ''' <param name="index">index</param>
        ''' <value>情報</value>
        ''' <returns>情報</returns>
        ''' <remarks></remarks>
        Default Property Item(ByVal index As Integer) As T

        ''' <summary>情報の一覧（順不同）</summary>
        ReadOnly Property Items() As ICollection(Of T)

        ''' <summary>indexの一覧（順不同）</summary>
        ReadOnly Property Indexes() As ICollection(Of Integer)

        ''' <summary>
        ''' 最大indexを返す
        ''' </summary>
        ''' <returns>最大index</returns>
        ''' <remarks></remarks>
        Function GetMaxIndex() As Integer?

        ''' <summary>
        ''' 情報を追加する
        ''' </summary>
        ''' <param name="value">情報</param>
        ''' <remarks></remarks>
        Sub Add(ByVal value As T)

        ''' <summary>
        ''' 情報を追加する
        ''' </summary>
        ''' <param name="values">情報</param>
        ''' <remarks></remarks>
        Sub AddRange(ByVal values As ICollection(Of T))

        ''' <summary>
        ''' 情報をすべてクリアする
        ''' </summary>
        ''' <remarks></remarks>
        Sub Clear()

        ''' <summary>
        ''' 情報をクリアする。後続の情報はそのまま。
        ''' </summary>
        ''' <remarks></remarks>
        Sub ClearAt(ByVal index As Integer)

        ''' <summary>
        ''' 情報を挿入する
        ''' </summary>
        ''' <param name="index">挿入index</param>
        ''' <param name="count">挿入数</param>
        ''' <remarks></remarks>
        Sub InsertAt(ByVal index As Integer, Optional ByVal count As Integer = 1)

        ''' <summary>
        ''' 情報を除去する。後続の情報は前詰め。
        ''' </summary>
        ''' <param name="index">除去index</param>
        ''' <param name="count">除去数</param>
        ''' <remarks></remarks>
        Sub RemoveAt(ByVal index As Integer, Optional ByVal count As Integer = 1)

        ''' <summary>
        ''' 情報を抽出する
        ''' </summary>
        ''' <param name="rowIndex">index</param>
        ''' <param name="count">取得数</param>
        ''' <remarks></remarks>
        Function Extract(ByVal rowIndex As Integer, ByVal count As Integer) As T()

        ''' <summary>
        ''' 情報を差し替える
        ''' </summary>
        ''' <param name="rowIndex">index</param>
        ''' <param name="recordBags">差し替える情報</param>
        ''' <remarks></remarks>
        Sub Supersede(ByVal rowIndex As Integer, ByVal recordBags As T())

        ''' <summary>
        ''' 位置indexを返す
        ''' </summary>
        ''' <param name="value">探査する値</param>
        ''' <returns>見つからなければ -1</returns>
        ''' <remarks></remarks>
        Function IndexOf(ByVal value As T) As Integer

        ''' <summary>
        ''' Indexの歯抜けやNothing値を除去して、前から整列しなおす
        ''' </summary>
        ''' <remarks></remarks>
        Sub SuppressGap()

        ''' <summary>
        ''' List末尾のNothing値を除去して、サイズを縮小する
        ''' </summary>
        ''' <remarks></remarks>
        Sub TrimEnd()
    End Interface
End Namespace