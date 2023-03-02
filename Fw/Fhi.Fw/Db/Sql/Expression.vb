Namespace Db.Sql

    ''' <summary>
    ''' 数式を担うクラス（数式を評価する）
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Expression
        ''' <summary>評価が不正である旨の値</summary>
        Private Shared ReadOnly EVALUATE_FAILURE As New EvaluateFailure

        ' 埋め込みパラメータ値
        Private ReadOnly param As Object

        Private ReadOnly formula As String
        Private ReadOnly terms As List(Of String)

#Region "Nested classes..."
        ''' <summary>評価が不正である旨のクラス</summary>
        Private Class EvaluateFailure
        End Class

        ''' <summary>
        ''' 四則演算が適切に処理できない例外
        ''' </summary>
        ''' <remarks></remarks>
        Public Class IllegalArithmeticException : Inherits ArgumentException

            Public Sub New(ByVal leftVal As Object, ByVal [operator] As String, ByVal rightVal As Object, ByVal formula As String)
                MyBase.New(String.Format("四則演算 {0} {1} {2} を処理できません. f='{3}'", leftVal, [operator], rightVal, formula), "formula")
            End Sub
        End Class
#End Region
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="formula">数式</param>
        ''' <param name="param">埋め込みパラメータ値</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal formula As String, ByVal param As Object)
            Me.formula = formula
            Me.param = param
        End Sub

        Private Sub New(ByVal param As Object, ByVal terms As List(Of String))
            Me.terms = terms
            Me.param = param
        End Sub

        ''' <summary>
        ''' 新しいインスタンスを生成する
        ''' </summary>
        ''' <param name="param">埋め込みパラメータ値</param>
        ''' <param name="terms">数式の1要素たち</param>
        ''' <returns>新しいインスタンス</returns>
        ''' <remarks></remarks>
        Public Shared Function NewInstance(ByVal param As Object, ByVal terms As List(Of String)) As Expression
            Return New Expression(param, terms)
        End Function

        ''' <summary>
        ''' 値を取得する
        ''' </summary>
        ''' <param name="obj">値object</param>
        ''' <param name="name">プロパティ名</param>
        ''' <returns>値</returns>
        ''' <remarks>括弧"XxxList(3)"を使用しcollection要素の参照可。"."経由で更にプロパティ値の参照可。</remarks>
        Public Shared Function GetValue(ByVal obj As Object, ByVal name As String) As Object
            Return VoUtil.GetValue(obj, name, EVALUATE_FAILURE)
        End Function

        Private Shared Function PerformGetValue(ByVal obj As Object, ByVal atName As String) As Object
            Return PerformGetValue(obj, atName, Nothing)
        End Function

        Private Shared Function PerformGetValue(ByVal obj As Object, ByVal atName As String, ByVal index As String) As Object
            If SqlBindAnalyzer.VALUE_BIND_PARAM_NAME.Equals(atName) Then
                If obj Is Nothing OrElse obj.GetType.IsValueType OrElse TypeOf obj Is String Then
                    Return obj
                End If
            End If
            If String.IsNullOrEmpty(index) Then
                Return GetValue(obj, atName.Substring(1))
            End If
            Return GetValue(obj, atName.Substring(1) & "(" & index & ")")
        End Function

        ''' <summary>
        ''' 計算する
        ''' </summary>
        ''' <returns>結果</returns>
        ''' <remarks></remarks>
        Public Function Calculate() As Object
            Const AT_MARK As Char = "@"c
            Dim postfixTerms As List(Of String) = If(formula IsNot Nothing, PostfixNotation.MakePostfix(formula), PostfixNotation.MakePostfix(terms))
            Dim skipIndex As Integer = -1
            Dim stack As New Stack
            For i As Integer = 0 To postfixTerms.Count - 1
                If i <= skipIndex Then
                    Continue For
                End If
                Dim term As String = postfixTerms(i)
                Dim termLower As String = term.ToLower
                Select Case termLower
                    Case "and"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If Not Convert.ToBoolean(leftVal) Then
                            stack.Push(False)
                        Else
                            stack.Push(Convert.ToBoolean(rightVal))
                        End If

                    Case "or"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If Convert.ToBoolean(leftVal) Then
                            stack.Push(True)
                        Else
                            stack.Push(Convert.ToBoolean(rightVal))
                        End If

                    Case "!"
                        Dim pop As Object = stack.Pop()
                        stack.Push(Not EzUtil.IsTrue(pop))

                    Case "=", "=="
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If leftVal Is EVALUATE_FAILURE OrElse rightVal Is EVALUATE_FAILURE Then
                            Return False
                        End If
                        stack.Push(IsEqualByString(leftVal, rightVal))

                    Case "!=", "<>"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If leftVal Is EVALUATE_FAILURE OrElse rightVal Is EVALUATE_FAILURE Then
                            Return False
                        End If
                        stack.Push(Not IsEqualByString(leftVal, rightVal))

                    Case "()"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If leftVal.ToString.Chars(0) = AT_MARK Then
                            Throw New NotSupportedException("ここ通ることがあるのか？")
                            'If SqlBindAnalyzer.VALUE_BIND_PARAM_NAME.Equals(leftVal.ToString) Then
                            '    If param Is Nothing OrElse param.GetType.IsValueType OrElse TypeOf param Is String Then
                            '        stack.Push(param)
                            '        Continue For
                            '    End If
                            'End If
                            'stack.Push(GetValue(param, term.Substring(1) & "(" & rightVal.ToString & ")"))
                        Else
                            stack.Push(GetValue(leftVal, "Value(" & rightVal.ToString & ")"))
                        End If

                    Case "."
                        Dim rightVal As String = stack.Pop().ToString
                        Dim leftVal As Object = stack.Pop()
                        stack.Push(GetValue(leftVal, rightVal))

                    Case ","
                        Throw New NotSupportedException(String.Format("'{0}' は未サポート", term))

                    Case "<"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If leftVal Is Nothing OrElse rightVal Is Nothing OrElse leftVal Is EVALUATE_FAILURE OrElse rightVal Is EVALUATE_FAILURE Then
                            stack.Push(False)
                        ElseIf IsNumeric(leftVal) AndAlso IsNumeric(rightVal) Then
                            stack.Push(CDec(leftVal) < CDec(rightVal))
                        Else
                            stack.Push(leftVal.ToString.CompareTo(rightVal.ToString) < 0)
                        End If

                    Case "<="
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If leftVal Is Nothing OrElse rightVal Is Nothing OrElse leftVal Is EVALUATE_FAILURE OrElse rightVal Is EVALUATE_FAILURE Then
                            stack.Push(False)
                        ElseIf IsNumeric(leftVal) AndAlso IsNumeric(rightVal) Then
                            stack.Push(CDec(leftVal) <= CDec(rightVal))
                        Else
                            stack.Push(leftVal.ToString.CompareTo(rightVal.ToString) <= 0)
                        End If

                    Case ">"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If leftVal Is Nothing OrElse rightVal Is Nothing OrElse leftVal Is EVALUATE_FAILURE OrElse rightVal Is EVALUATE_FAILURE Then
                            stack.Push(False)
                        ElseIf IsNumeric(leftVal) AndAlso IsNumeric(rightVal) Then
                            stack.Push(CDec(leftVal) > CDec(rightVal))
                        Else
                            stack.Push(leftVal.ToString.CompareTo(rightVal.ToString) > 0)
                        End If

                    Case ">="
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If leftVal Is Nothing OrElse rightVal Is Nothing OrElse leftVal Is EVALUATE_FAILURE OrElse rightVal Is EVALUATE_FAILURE Then
                            stack.Push(False)
                        ElseIf IsNumeric(leftVal) AndAlso IsNumeric(rightVal) Then
                            stack.Push(CDec(leftVal) >= CDec(rightVal))
                        Else
                            stack.Push(leftVal.ToString.CompareTo(rightVal.ToString) >= 0)
                        End If

                    Case "+"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If leftVal Is Nothing OrElse rightVal Is Nothing OrElse leftVal Is EVALUATE_FAILURE OrElse rightVal Is EVALUATE_FAILURE Then
                            Throw New IllegalArithmeticException(leftVal, "+", rightVal, formula)
                        End If
                        If IsNumeric(leftVal) AndAlso IsNumeric(rightVal) Then
                            stack.Push(CDec(leftVal) + CDec(rightVal))
                        ElseIf TypeOf leftVal Is String OrElse TypeOf rightVal Is String Then
                            stack.Push(If(leftVal IsNot Nothing, leftVal.ToString, "") & If(rightVal IsNot Nothing, rightVal.ToString, ""))
                        Else
                            Throw New IllegalArithmeticException(leftVal, "+", rightVal, formula)
                        End If

                    Case "-"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If Not IsNumeric(leftVal) OrElse Not IsNumeric(rightVal) Then
                            Throw New IllegalArithmeticException(leftVal, "-", rightVal, formula)
                        End If
                        stack.Push(CDec(leftVal) - CDec(rightVal))

                    Case "*"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If Not IsNumeric(leftVal) OrElse Not IsNumeric(rightVal) Then
                            Throw New IllegalArithmeticException(leftVal, "*", rightVal, formula)
                        End If
                        stack.Push(CDec(leftVal) * CDec(rightVal))

                    Case "/"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If Not IsNumeric(leftVal) OrElse Not IsNumeric(rightVal) Then
                            Throw New IllegalArithmeticException(leftVal, "/", rightVal, formula)
                        End If
                        stack.Push(CDec(leftVal) / CDec(rightVal))

                    Case "%"
                        Dim rightVal As Object = stack.Pop()
                        Dim leftVal As Object = stack.Pop()
                        If Not IsNumeric(leftVal) OrElse Not IsNumeric(rightVal) Then
                            Throw New IllegalArithmeticException(leftVal, "%", rightVal, formula)
                        End If
                        stack.Push(CDec(leftVal) Mod CDec(rightVal))

                    Case "true"
                        stack.Push(True)

                    Case "false"
                        stack.Push(False)

                    Case "null"
                        stack.Push(Nothing)

                    Case Else
                        If term.Chars(0) = AT_MARK Then
                            stack.Push(PerformGetValue(param, atName:=term))
                        Else
                            stack.Push(term)
                        End If
                End Select
                If TypeOf stack.Peek Is Boolean Then
                    skipIndex = PostfixNotation.DetectSkipIndexFromPostfix(DirectCast(stack.Peek, Boolean), postfixTerms, i + 1)
                End If
            Next
            Return stack.Peek
        End Function

        ''' <summary>
        ''' 数式(論理演算式)を評価する
        ''' </summary>
        ''' <returns>結果</returns>
        ''' <remarks></remarks>
        Public Function Evaluate() As Boolean
            Return Evaluate(Calculate)
        End Function

        ''' <summary>
        ''' 数式(論理演算式)を評価する
        ''' </summary>
        ''' <param name="calculatedValue">計算結果値</param>
        ''' <returns>結果</returns>
        ''' <remarks></remarks>
        Public Shared Function Evaluate(ByVal calculatedValue As Object) As Boolean
            If calculatedValue Is EVALUATE_FAILURE Then
                Return False
            End If
            Return Convert.ToBoolean(calculatedValue)
        End Function

        ''' <summary>
        ''' 文字列化した値で同値評価する
        ''' </summary>
        ''' <param name="leftVal"></param>
        ''' <param name="rightVal"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function IsEqualByString(ByVal leftVal As Object, ByVal rightVal As Object) As Boolean

            If leftVal Is Nothing AndAlso rightVal Is Nothing Then
                Return True
            ElseIf leftVal Is Nothing OrElse rightVal Is Nothing Then
                Return False
            End If
            Return leftVal.ToString.Equals(rightVal.ToString)
        End Function
    End Class
End Namespace