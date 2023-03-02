Namespace Domain
    ''' <summary>
    ''' 不変のコレクションオブジェクトのObject型Interface
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface ICollectionObject
        ''' <summary>要素の数</summary>
        ReadOnly Property Count() As Integer

        ''' <summary>要素（値オブジェクト）</summary>
        Default ReadOnly Property Item(index As Integer) As Object
    End Interface
End Namespace