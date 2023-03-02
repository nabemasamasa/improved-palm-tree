Namespace Db
    ''' <summary>
    ''' SQLの検索条件を設定するインターフェース
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface CriteriaBinder

        ''' <summary>
        ''' 検索条件を取得する
        ''' </summary>
        ''' <returns>検索条件[]</returns>
        ''' <remarks></remarks>
        Function GetParameters() As CriteriaParameter()

        ''' <summary>
        ''' valueと等しいか
        ''' field = value ／ field IS NULL
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Equal(field As Object, value As Object) As CriteriaBinder

        ''' <summary>
        ''' valuesのいずれかに等しいか
        ''' field IN (values)
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="values"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Any(field As Object, values As IEnumerable) As CriteriaBinder

        ''' <summary>
        ''' ワイルドカード検索
        ''' field LIKE value
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function [Like](field As Object, value As Object) As CriteriaBinder

        ''' <summary>
        ''' valueより大きいか
        ''' value &lt; field
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GreaterThan(field As Object, value As Object) As CriteriaBinder

        ''' <summary>
        ''' valueより小さいか
        ''' field &lt; value
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function LessThan(field As Object, value As Object) As CriteriaBinder

        ''' <summary>
        ''' value以上か
        ''' value &lt;= field
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GreaterEqual(field As Object, value As Object) As CriteriaBinder

        ''' <summary>
        ''' value以下か
        ''' field &lt;= value
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function LessEqual(field As Object, value As Object) As CriteriaBinder

        ''' <summary>
        ''' 識別名から対応した値を取得する
        ''' </summary>
        ''' <param name="identifyName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetValueByIdentifyName(identifyName As String) As Object

        ''' <summary>
        ''' かつ
        ''' AND (callback)
        ''' </summary>
        ''' <param name="callback"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function [And](callback As Action) As CriteriaBinder

        ''' <summary>
        ''' または
        ''' OR
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function [Or]() As CriteriaBinder

        ''' <summary>
        ''' または
        ''' OR (callback)
        ''' </summary>
        ''' <param name="callback"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function [Or](callback As Action) As CriteriaBinder

        ''' <summary>
        ''' ではない(否)
        ''' NOT
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function [Not]() As CriteriaBinder

        ''' <summary>
        ''' ではない(否)
        ''' NOT (callback)
        ''' </summary>
        ''' <param name="callback"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function [Not](callback As Action) As CriteriaBinder

        ''' <summary>
        ''' 入れ子にする
        ''' 「AND (」や「OR (」の直後に括弧でくくりたいとき用
        ''' </summary>
        ''' <param name="callback"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Nest(callback As Action) As CriteriaBinder
    End Interface
End Namespace