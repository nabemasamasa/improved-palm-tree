Namespace Db.Sql
    Public interface ISqlExpressionEvaluator
        ''' <summary>
        ''' 数式を評価して結果を返す
        ''' </summary>
        ''' <returns>結果</returns>
        ''' <remarks></remarks>
        Function Evaluate(ByVal expression As String, ByVal param As Object) As Boolean
    end interface
End NameSpace