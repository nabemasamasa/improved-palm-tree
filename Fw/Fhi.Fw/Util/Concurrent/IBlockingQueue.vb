Namespace Util.Concurrent
    Public Interface IBlockingQueue(Of T)
        ': Inherits ICollection(Of T), IEnumerable(Of T)
        ''' <summary>
        ''' 指定要素をこのキューに挿入する
        ''' </summary>
        ''' <param name="value">要素</param>
        ''' <remarks></remarks>
        Sub Put(ByVal value As T)

        ''' <summary>
        ''' このキューの先頭を取得して削除する
        ''' </summary>
        ''' <returns>キューの先頭</returns>
        ''' <remarks></remarks>
        Function Take() As T
    End Interface
End Namespace