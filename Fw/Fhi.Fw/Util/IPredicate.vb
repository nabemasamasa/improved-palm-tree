Namespace Util
    ''' <summary>
    ''' List#Find()用の Predicate(Of T) デリゲートのInterface
    ''' </summary>
    ''' <typeparam name="T">Listの型、判定する型</typeparam>
    ''' <remarks></remarks>
    Public Interface IPredicate(Of T)
        ''' <summary>
        ''' 一致するか？を判定する
        ''' </summary>
        ''' <param name="element">判定する値</param>
        ''' <returns>一致する場合、true</returns>
        ''' <remarks></remarks>
        Function Judge(ByVal element As T) As Boolean
    End Interface
End Namespace