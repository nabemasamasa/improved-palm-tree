Namespace App.Xls
    ''' <summary>
    ''' セル範囲定義
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class CellRangeDefinition
        ''' <summary>Voの型</summary>
        Public Property ObjectType As Type

        ''' <summary>開始行 ※1はじまり</summary>
        Public Property StartRow As Integer

        ''' <summary>開始列 ※1はじまり</summary>
        Public Property StartCol As Integer

        ''' <summary>変換ルール</summary>
        Public Property MutableRules() As IEnumerable(Of XlVoPropertyRule)
    End Class
End Namespace
