Namespace Db.Sql

    ''' <summary>
    ''' 数式を担うクラス（数式を評価する）
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ExpressionOld
        Private Class EvaluateFailure
        End Class

        Private Shared ReadOnly BIND_PARAM_VALUE_FAILURE As New EvaluateFailure

        Private Const EQUAL As String = "=="
        Private Const EQUAL2 As String = "="
        Private Const NOT_EQUAL As String = "!="
        Private Const GREATER_THAN As String = ">"
        Private Const GREATER_EQUAL As String = ">="
        Private Const LESS_THAN As String = "<"
        Private Const LESS_EQUAL As String = "<="

        Private Const LOGICAL_AND_LOWER As String = "and"
        Private Const LOGICAL_OR_LOWER As String = "or"

        ' 埋め込みパラメータ値
        Private param As Object
        ' 数式の1要素  (ex. x+y=z だったら "x" / "+" / "y" / "=" / "z" のどれか)
        Private term As String

        Private expressionA As ExpressionOld
        Private expressionB As ExpressionOld
        Private expressionC As ExpressionOld

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="param">埋め込みパラメータ値</param>
        ''' <param name="term">数式の1要素</param>
        ''' <remarks></remarks>
        Private Sub New(ByVal param As Object, ByVal term As String)
            Me.term = term
            Me.param = param
        End Sub
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="value1"></param>
        ''' <param name="value2"></param>
        ''' <param name="value3"></param>
        ''' <remarks></remarks>
        Private Sub New(ByVal value1 As ExpressionOld, ByVal value2 As ExpressionOld, ByVal value3 As ExpressionOld)
            Me.expressionA = value1
            Me.expressionB = value2
            Me.expressionC = value3
            If value2.CanEvaluate Then
                Throw New ArgumentException(String.Format("'{0} {1} {2}' の式は評価出来ません", value1.term, value2.term, value3.term))
            End If
        End Sub

        ''' <summary>
        ''' 新しいインスタンスを生成する
        ''' </summary>
        ''' <param name="param">埋め込みパラメータ値</param>
        ''' <param name="terms">数式の1要素たち</param>
        ''' <returns>新しいインスタンス</returns>
        ''' <remarks></remarks>
        Public Shared Function NewInstance(ByVal param As Object, ByVal terms As List(Of String)) As ExpressionOld
            Dim sikis As New List(Of ExpressionOld)
            For Each term As String In terms
                sikis.Add(New ExpressionOld(param, term))
            Next
            Return CollapseSiki(sikis)
        End Function

        Private Shared Function CollapseSiki(ByVal sikis As List(Of ExpressionOld)) As ExpressionOld
            If sikis.Count = 1 Then
                Return sikis(0)
            End If
            Dim argList As New List(Of ExpressionOld)
            Dim start As Integer = 0

            If IsLogical(sikis(start + 1)) Then
                If start + 3 < sikis.Count AndAlso Not IsLogical(sikis(start + 3)) Then
                    argList.Add(sikis(start))
                    argList.Add(sikis(start + 1))
                    start += 2
                End If
            End If
            If start + 2 < sikis.Count Then
                argList.Add(New ExpressionOld(sikis(start), sikis(start + 1), sikis(start + 2)))
                start += 3
            Else
                Throw New ArgumentException()
            End If
            For i As Integer = start To sikis.Count - 1
                argList.Add(sikis(i))
            Next
            Return CollapseSiki(argList)
        End Function
        ''' <summary>
        ''' 論理演算子かを返す
        ''' </summary>
        ''' <param name="aExpression">演算子</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Private Shared Function IsLogical(ByVal aExpression As ExpressionOld) As Boolean
            Dim term As String = aExpression.term.ToLower
            Return LOGICAL_AND_LOWER.Equals(term) OrElse LOGICAL_OR_LOWER.Equals(term)
        End Function
        ''' <summary>
        ''' 評価出来るインスタンスかを返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Private Function CanEvaluate() As Boolean
            Return term = Nothing
        End Function
        Private ReadOnly Property Value() As Object
            Get
                If CanEvaluate() Then
                    Return Evaluate()
                ElseIf term.Chars(0) = "@"c Then
                    If SqlBindAnalyzer.VALUE_BIND_PARAM_NAME.Equals(term) Then
                        If param Is Nothing OrElse param.GetType.IsValueType OrElse TypeOf param Is String Then
                            Return param
                        End If
                    End If
                    Return GetValue(param, term.Substring(1))
                End If
                Dim termLower As String = term.ToLower
                If "true".Equals(termLower) Then
                    Return True
                ElseIf "false".Equals(termLower) Then
                    Return False
                ElseIf "null".Equals(termLower) Then
                    Return Nothing
                End If
                Return term
            End Get
        End Property

        ''' <summary>
        ''' 値を取得する
        ''' </summary>
        ''' <param name="obj">値object</param>
        ''' <param name="name">プロパティ名</param>
        ''' <returns>値</returns>
        ''' <remarks>括弧"XxxList(3)"を使用しcollection要素の参照可。"."経由で更にプロパティ値の参照可。</remarks>
        Public Shared Function GetValue(ByVal obj As Object, ByVal name As String) As Object
            Return VoUtil.GetValue(obj, name, BIND_PARAM_VALUE_FAILURE)
        End Function

        ''' <summary>
        ''' 数式(論理演算式)を評価する
        ''' </summary>
        ''' <returns>結果</returns>
        ''' <remarks></remarks>
        Public Function Evaluate() As Boolean
            If Not CanEvaluate() Then
                Dim result As Object = Value
                If TypeOf result Is Boolean Then
                    Return Convert.ToBoolean(result)
                Else
                    Throw New ArgumentException(term & " は評価出来ません.")
                End If
            End If

            Dim value2 As String = expressionB.Value.ToString
            If LOGICAL_AND_LOWER.Equals(value2.ToLower) Then
                ' A and C の評価
                Dim aValue As Boolean = Convert.ToBoolean(expressionA.Value)
                If aValue = False Then
                    Return False
                End If
                Return Convert.ToBoolean(expressionC.Value)
            ElseIf LOGICAL_OR_LOWER.Equals(value2.ToLower) Then
                ' A or C の評価
                Dim aValue As Boolean = Convert.ToBoolean(expressionA.Value)
                If aValue = True Then
                    Return True
                End If
                Return Convert.ToBoolean(expressionC.Value)
            End If

            If Not NOT_EQUAL.Equals(value2) AndAlso Not EQUAL.Equals(value2) AndAlso Not EQUAL2.Equals(value2) AndAlso Not GREATER_THAN.Equals(value2) AndAlso Not GREATER_EQUAL.Equals(value2) AndAlso Not LESS_THAN.Equals(value2) AndAlso Not LESS_EQUAL.Equals(value2) Then
                Throw New NotSupportedException("演算子 " & value2 & " は未対応.")
            End If
            Dim leftValue As Object = expressionA.Value
            If leftValue Is BIND_PARAM_VALUE_FAILURE Then
                Return False
            End If
            Dim rightValue As Object = expressionC.Value
            If rightValue Is BIND_PARAM_VALUE_FAILURE Then
                Return False
            End If

            Dim values() As Object = {leftValue, rightValue}
            For i As Integer = 0 To values.Length - 1
                If values(i) Is Nothing Then
                    If NOT_EQUAL.Equals(value2) Then
                        Return values(1 - i) IsNot Nothing
                    ElseIf EQUAL.Equals(value2) OrElse EQUAL2.Equals(value2) Then
                        Return values(1 - i) Is Nothing
                    End If
                    Return False
                End If
            Next
            If NOT_EQUAL.Equals(value2) Then
                Return Not leftValue.ToString.Equals(rightValue.ToString)
            ElseIf EQUAL.Equals(value2) OrElse EQUAL2.Equals(value2) Then
                Return leftValue.ToString.Equals(rightValue.ToString)
            ElseIf LESS_THAN.Equals(value2) Then
                If IsNumeric(leftValue) AndAlso IsNumeric(rightValue) Then
                    Return Decimal.Parse(leftValue.ToString) < Decimal.Parse(rightValue.ToString)
                Else
                    Return leftValue.ToString.CompareTo(rightValue.ToString) < 0
                End If
            ElseIf LESS_EQUAL.Equals(value2) Then
                If IsNumeric(leftValue) AndAlso IsNumeric(rightValue) Then
                    Return Decimal.Parse(leftValue.ToString) <= Decimal.Parse(rightValue.ToString)
                Else
                    Return leftValue.ToString.CompareTo(rightValue.ToString) <= 0
                End If
            ElseIf GREATER_THAN.Equals(value2) Then
                If IsNumeric(leftValue) AndAlso IsNumeric(rightValue) Then
                    Return Decimal.Parse(leftValue.ToString) > Decimal.Parse(rightValue.ToString)
                Else
                    Return leftValue.ToString.CompareTo(rightValue.ToString) > 0
                End If
            ElseIf GREATER_EQUAL.Equals(value2) Then
                If IsNumeric(leftValue) AndAlso IsNumeric(rightValue) Then
                    Return Decimal.Parse(leftValue.ToString) >= Decimal.Parse(rightValue.ToString)
                Else
                    Return leftValue.ToString.CompareTo(rightValue.ToString) >= 0
                End If
            End If
            Throw New NotSupportedException(String.Format("'{0} {1} {2}' この式は評価出来ません.", leftValue, value2, rightValue))
        End Function
    End Class
End Namespace