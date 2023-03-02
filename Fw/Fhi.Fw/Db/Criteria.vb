Imports Fhi.Fw.Util

Namespace Db
    ''' <summary>
    ''' SQLの検索条件を設定するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Criteria(Of T) : Implements CriteriaBinder

        Private ReadOnly voMarker As New VoPropertyMarker
        Private ReadOnly _params As New List(Of CriteriaParameter)

        Public Sub New(tableVo As T)
            voMarker.MarkVo(tableVo)
        End Sub

        ''' <summary>
        ''' 検索条件を取得する
        ''' </summary>
        ''' <returns>検索条件[]</returns>
        ''' <remarks></remarks>
        Public Function GetParameters() As CriteriaParameter() Implements CriteriaBinder.GetParameters
            Return _params.ToArray()
        End Function

        Private Sub AddCondition(field As Object, condition As CriteriaParameter.SearchCondition, value As Object)
            AddAndConditionIfNecessary()
            _params.Add(New CriteriaParameter(voMarker.GetMemberName(field), GetIdentifyName(field), condition, value))
        End Sub

        Private Sub AddAndConditionIfNecessary()
            If CollectionUtil.IsNotEmpty(_params) AndAlso (_params.Last().PropertyName IsNot Nothing _
                                                           OrElse _params.Last().Condition = CriteriaParameter.SearchCondition.CloseBracket) Then
                AddOperator(CriteriaParameter.SearchCondition.And)
            End If
        End Sub

        Private Sub AddOperator(condition As CriteriaParameter.SearchCondition)
            _params.Add(New CriteriaParameter(Nothing, Nothing, condition, Nothing))
        End Sub

        Private identifyNumber As Integer = 0
        Private Function GetIdentifyName(field As Object) As String
            Return String.Format("{0}{1}", voMarker.GetMemberName(field), EzUtil.Increment(identifyNumber))
        End Function

        ''' <summary>
        ''' valueと等しいか
        ''' field = value ／ field IS NULL
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Equal(field As Object, value As Object) As CriteriaBinder Implements CriteriaBinder.Equal
            AddCondition(field, CriteriaParameter.SearchCondition.Equal, value)
            Return Me
        End Function

        ''' <summary>
        ''' valuesのいずれかに等しいか
        ''' field IN (values)
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="values"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Any(field As Object, values As IEnumerable) As CriteriaBinder Implements CriteriaBinder.Any
            AddCondition(field, CriteriaParameter.SearchCondition.Any, values.Cast(Of Object)().ToArray())
            Return Me
        End Function

        ''' <summary>
        ''' ワイルドカード検索
        ''' field LIKE value
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function [Like](field As Object, value As Object) As CriteriaBinder Implements CriteriaBinder.[Like]
            AddCondition(field, CriteriaParameter.SearchCondition.[Like], value)
            Return Me
        End Function

        ''' <summary>
        ''' valueより大きいか
        ''' value &lt; field
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GreaterThan(field As Object, value As Object) As CriteriaBinder Implements CriteriaBinder.GreaterThan
            AddCondition(field, CriteriaParameter.SearchCondition.GreaterThan, value)
            Return Me
        End Function

        ''' <summary>
        ''' valueより小さいか
        ''' field &lt; value
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LessThan(field As Object, value As Object) As CriteriaBinder Implements CriteriaBinder.LessThan
            AddCondition(field, CriteriaParameter.SearchCondition.LessThan, value)
            Return Me
        End Function

        ''' <summary>
        ''' value以上か
        ''' value &lt;= field
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GreaterEqual(field As Object, value As Object) As CriteriaBinder Implements CriteriaBinder.GreaterEqual
            AddCondition(field, CriteriaParameter.SearchCondition.GreaterEqual, value)
            Return Me
        End Function

        ''' <summary>
        ''' value以下か
        ''' field &lt;= value
        ''' </summary>
        ''' <param name="field"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LessEqual(field As Object, value As Object) As CriteriaBinder Implements CriteriaBinder.LessEqual
            AddCondition(field, CriteriaParameter.SearchCondition.LessEqual, value)
            Return Me
        End Function

        ''' <summary>
        ''' かつ
        ''' AND (callback)
        ''' </summary>
        ''' <param name="callback"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function [And](callback As Action) As CriteriaBinder Implements CriteriaBinder.[And]
            AddOperator(CriteriaParameter.SearchCondition.AndBracket)
            callback.Invoke()
            AddOperator(CriteriaParameter.SearchCondition.CloseBracket)
            Return Me
        End Function

        ''' <summary>
        ''' または
        ''' OR
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function [Or]() As CriteriaBinder Implements CriteriaBinder.[Or]
            AddOperator(CriteriaParameter.SearchCondition.Or)
            Return Me
        End Function

        ''' <summary>
        ''' または
        ''' OR (callback)
        ''' </summary>
        ''' <param name="callback"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function [Or](callback As Action) As CriteriaBinder Implements CriteriaBinder.[Or]
            AddOperator(CriteriaParameter.SearchCondition.OrBracket)
            callback.Invoke()
            AddOperator(CriteriaParameter.SearchCondition.CloseBracket)
            Return Me
        End Function

        ''' <summary>
        ''' ではない/否
        ''' NOT
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function [Not]() As CriteriaBinder Implements CriteriaBinder.[Not]
            AddAndConditionIfNecessary()
            AddOperator(CriteriaParameter.SearchCondition.Not)
            Return Me
        End Function

        ''' <summary>
        ''' ではない/否
        ''' NOT (callback)
        ''' </summary>
        ''' <param name="callback"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function [Not](callback As Action) As CriteriaBinder Implements CriteriaBinder.[Not]
            AddAndConditionIfNecessary()
            AddOperator(CriteriaParameter.SearchCondition.NotBracket)
            callback.Invoke()
            AddOperator(CriteriaParameter.SearchCondition.CloseBracket)
            Return Me
        End Function

        ''' <summary>
        ''' 入れ子にする
        ''' 「AND (」や「OR (」の直後に括弧でくくりたいとき用
        ''' </summary>
        ''' <param name="callback"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Nest(callback As Action) As CriteriaBinder Implements CriteriaBinder.Nest
            AddOperator(CriteriaParameter.SearchCondition.Bracket)
            callback.Invoke()
            AddOperator(CriteriaParameter.SearchCondition.CloseBracket)
            Return Me
        End Function

        ''' <summary>
        ''' 識別名から対応した値を取得する
        ''' </summary>
        ''' <param name="identifyName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetValueByIdentifyName(identifyName As String) As Object Implements CriteriaBinder.GetValueByIdentifyName
            Const OPEN_BRACKET As Char = "("c
            Const CLOSE_BRACKET As Char = ")"c
            If Not identifyName.Contains(OPEN_BRACKET) Then
                Return GetParameters.First(Function(binder) identifyName.Equals(binder.IdentifyName)).Value
            End If
            Dim identPropertyName As String = StringUtil.Left(identifyName, identifyName.IndexOf(OPEN_BRACKET))
            Dim startIndex As Integer = identifyName.IndexOf(OPEN_BRACKET) + 1
            Dim endIndex As Integer = identifyName.IndexOf(CLOSE_BRACKET)
            Dim collectionIndex As Integer = CInt(identifyName.Substring(startIndex, endIndex - startIndex))
            Dim values As Object = GetParameters.First(Function(binder) identPropertyName.Equals(binder.IdentifyName)).Value
            Return DirectCast(values, Object())(collectionIndex)
        End Function

    End Class
End Namespace