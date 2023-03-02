Imports System.Text

Namespace Db.Sql
    ''' <summary>
    ''' 式を後置記法にすることを担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class PostfixNotation

        ' 演算子優先順位
        '() .
        '!
        '* / %
        '+ -
        '< <= > >=
        '= != == <>
        'and or

        ''' <summary>演算子の優先順位[]</summary>
        Private Shared ReadOnly OPERATORS_BY_ORDER As List(Of String)() = {New List(Of String)(New String() {"()", "."}), _
                                                                           New List(Of String)(New String() {"!"}), _
                                                                           New List(Of String)(New String() {"*", "/", "%"}), _
                                                                           New List(Of String)(New String() {"+", "-"}), _
                                                                           New List(Of String)(New String() {"<", "<=", ">", ">="}), _
                                                                           New List(Of String)(New String() {"=", "==", "!=", "<>"}), _
                                                                           New List(Of String)(New String() {"and", "or"}), _
                                                                           New List(Of String)(New String() {","})}
        '''' <summary>単項目演算子[]</summary>
        'Private Shared ReadOnly UNARY_OPERATORS As New List(Of String)(New String() {"!"})
        ''' <summary>括弧[]</summary>
        Private Shared ReadOnly BRACKET_PAIR_OPERATORS As New List(Of String)(New String() {"()"})
        ''' <summary>すべての演算子[]</summary>
        Private Shared ReadOnly ALL_OPERATORS As List(Of String)

        Shared Sub New()
            ALL_OPERATORS = New List(Of String)
            For Each ops As List(Of String) In OPERATORS_BY_ORDER
                ALL_OPERATORS.AddRange(ops)
            Next
            For Each bracketPair As String In BRACKET_PAIR_OPERATORS
                ALL_OPERATORS.Remove(bracketPair)
                ALL_OPERATORS.Add(bracketPair.Substring(0, 1))
                ALL_OPERATORS.Add(bracketPair.Substring(1, 1))
            Next
        End Sub

        Private ReadOnly formula As String
        Private ReadOnly infixTerms As IEnumerable(Of String)

#Region "Nested classes..."
        Public Class InvalidFormulaException : Inherits ArgumentException

            Public Sub New(ByVal formula As String)
                MyBase.New("数式を正しく評価できません. '" & formula & "'", "formula")
            End Sub
        End Class

        Public Class IllegalFormulaTermException : Inherits ArgumentException

            Public Sub New(ByVal formula As String, ByVal term As String)
                MyBase.New("項 '" & term & "'が不明です. f='" & formula & "'", "formula")
            End Sub
        End Class

        Public Class IllegalFormulaEndException : Inherits ArgumentException

            Public Sub New(ByVal formula As String)
                MyBase.New("式が途中で終わってます. f='" & formula & "'", "formula")
            End Sub
        End Class

        Public Class InvalidBracketException : Inherits ArgumentException

            Public Sub New(ByVal formula As String, ByVal bracket As String)
                MyBase.New("対応する括弧 '" & bracket & "' がありません. f='" & formula & "'", "formula")
            End Sub
        End Class

#End Region
#Region "Public properties..."
        Private _postfixTerms As List(Of String)
        ''' <summary>後置記法の結果[]</summary>
        Public ReadOnly Property ResultPostfix() As List(Of String)
            Get
                Return _postfixTerms
            End Get
        End Property
#End Region

        ''' <summary>
        ''' 式を後置記法（逆ポーランド記法）にして返す
        ''' </summary>
        ''' <param name="formula">数式</param>
        ''' <returns>逆ポーランド記法にした演算子とオペランド[]</returns>
        ''' <remarks></remarks>
        Public Shared Function MakePostfix(ByVal formula As String) As List(Of String)
            Dim postfix As New PostfixNotation(formula)
            postfix.Make()
            Return postfix.ResultPostfix
        End Function

        ''' <summary>
        ''' 式を後置記法（逆ポーランド記法）にして返す
        ''' </summary>
        ''' <param name="infixTerms">中置記法</param>
        ''' <returns>逆ポーランド記法にした演算子とオペランド[]</returns>
        ''' <remarks></remarks>
        Friend Shared Function MakePostfix(ByVal infixTerms As IEnumerable(Of String)) As List(Of String)
            Dim infix As List(Of String) = New List(Of String)(infixTerms)
            Dim postfix As New PostfixNotation(Join(infix.ToArray, " "), infix)
            postfix.Make()
            Return postfix.ResultPostfix
        End Function

        Public Sub New(ByVal formula As String)
            Me.New(formula, Nothing)
        End Sub

        Public Sub New(ByVal formula As String, ByVal infixTerms As IEnumerable(Of String))
            Me.formula = formula
            Me.infixTerms = infixTerms
        End Sub

        ''' <summary>
        ''' 式を後置記法（逆ポーランド記法）にする
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Make()
            Dim postfix As New List(Of String)
            Dim infixTerms As List(Of String) = DirectCast(If(Me.infixTerms, SplitFormula(formula)), List(Of String))
            MakePostfixOperator(infixTerms, postfix, operatorOrderDescIndex:=0)
            If 0 < infixTerms.Count Then
                Throw New IllegalFormulaTermException(formula, infixTerms(0))
            End If
            _postfixTerms = postfix
        End Sub

        ''' <summary>
        ''' 演算子優先順位での後置記法を作成する
        ''' </summary>
        ''' <param name="infixTerms">中置記法のコレクション</param>
        ''' <param name="resultPostfixTerms">後置記法にした結果</param>
        ''' <param name="operatorOrderDescIndex">演算子優先順位の降順index</param>
        ''' <remarks></remarks>
        Private Sub MakePostfixOperator(ByVal infixTerms As List(Of String), ByVal resultPostfixTerms As List(Of String), _
                                        ByVal operatorOrderDescIndex As Integer)
            If OPERATORS_BY_ORDER.Length <= operatorOrderDescIndex Then
                Return
            End If
            Const BRACKET_PAIR As String = "()"
            Const SQUARE_BRACKET_PAIR As String = "[]"
            Dim wasSetOperand As Boolean = PerformMakePostfixAndIsSetOperand(infixTerms, resultPostfixTerms, operatorOrderDescIndex + 1)
            While 0 < infixTerms.Count
                If OPERATORS_BY_ORDER(OPERATORS_BY_ORDER.Length - 1 - operatorOrderDescIndex).Contains(infixTerms(0)) Then
                    Dim thisTerm As String = infixTerms(0)
                    infixTerms.RemoveAt(0)
                    wasSetOperand = PerformMakePostfixAndIsSetOperand(infixTerms, resultPostfixTerms, operatorOrderDescIndex + 1)
                    resultPostfixTerms.Add(thisTerm)

                ElseIf "(".Equals(infixTerms(0)) AndAlso OPERATORS_BY_ORDER(OPERATORS_BY_ORDER.Length - 1 - operatorOrderDescIndex).Contains(BRACKET_PAIR) Then
                    MakePostfixBracket(infixTerms, resultPostfixTerms, BRACKET_PAIR, wasSetOperand)

                ElseIf "[".Equals(infixTerms(0)) AndAlso OPERATORS_BY_ORDER(OPERATORS_BY_ORDER.Length - 1 - operatorOrderDescIndex).Contains(SQUARE_BRACKET_PAIR) Then
                    MakePostfixBracket(infixTerms, resultPostfixTerms, SQUARE_BRACKET_PAIR, wasSetOperand)

                Else
                    Return
                End If
            End While
        End Sub

        ''' <summary>
        ''' 括弧内の後置記法を作成する
        ''' </summary>
        ''' <param name="infixTerms"></param>
        ''' <param name="resultPostfixTerms"></param>
        ''' <param name="bracketPair"></param>
        ''' <param name="wasSetOperand">オペランドを処理した直後なら、true</param>
        ''' <remarks></remarks>
        Private Sub MakePostfixBracket(ByVal infixTerms As List(Of String), ByVal resultPostfixTerms As List(Of String), ByVal bracketPair As String, ByVal wasSetOperand As Boolean)
            If Not bracketPair.Substring(0, 1).Equals(infixTerms(0)) Then
                Throw New InvalidOperationException("")
            End If
            infixTerms.RemoveAt(0)
            Dim count As Integer = resultPostfixTerms.Count
            MakePostfixOperator(infixTerms, resultPostfixTerms, operatorOrderDescIndex:=0)
            If 0 < infixTerms.Count AndAlso bracketPair.Substring(1).Equals(infixTerms(0)) Then
                infixTerms.RemoveAt(0)
                If wasSetOperand Then
                    If count = resultPostfixTerms.Count Then
                        resultPostfixTerms.Add("\")
                    End If
                    resultPostfixTerms.Add(bracketPair)
                End If
            Else
                If infixTerms.Count = 0 Then
                    Throw New InvalidBracketException(formula, bracketPair.Substring(1))
                Else
                    Throw New IllegalFormulaTermException(formula, infixTerms(0))
                End If
            End If
        End Sub

        ''' <summary>
        ''' 演算子優先順位の後置記法を作成する
        ''' </summary>
        ''' <param name="infixTerms">中置記法のコレクション</param>
        ''' <param name="resultPostfixTerms">後置記法にした結果（途中）</param>
        ''' <param name="operatorOrderDescIndex">演算子優先順位の降順index</param>
        ''' <returns>オペランドを設定した直後なら、true</returns>
        ''' <remarks></remarks>
        Private Function PerformMakePostfixAndIsSetOperand(ByVal infixTerms As List(Of String), ByVal resultPostfixTerms As List(Of String), ByVal operatorOrderDescIndex As Integer) As Boolean

            If operatorOrderDescIndex < OPERATORS_BY_ORDER.Length Then
                MakePostfixOperator(infixTerms, resultPostfixTerms, operatorOrderDescIndex)
                Return False
            End If
            Return MakePostfixOperand(infixTerms, resultPostfixTerms)
        End Function

        ''' <summary>
        ''' オペランドなら後置記法を追記する
        ''' </summary>
        ''' <param name="infixTerms">中置記法のコレクション</param>
        ''' <param name="resultPostfixTerms">後置記法にした結果（途中）</param>
        ''' <returns>オペランドだった場合、true</returns>
        ''' <remarks></remarks>
        Private Function MakePostfixOperand(ByVal infixTerms As List(Of String), ByVal resultPostfixTerms As List(Of String)) As Boolean
            If infixTerms.Count = 0 Then
                Throw New IllegalFormulaEndException(formula)
            End If
            If ALL_OPERATORS.Contains(infixTerms(0)) Then
                Return False
            End If
            resultPostfixTerms.Add(infixTerms(0))
            infixTerms.RemoveAt(0)
            Return True
        End Function

        Private Const WHITE_SPACE As String = " " & vbCrLf & vbTab

        ''' <summary>
        ''' Whitespaceかを返す
        ''' </summary>
        ''' <param name="c">判定文字</param>
        ''' <returns>Whitespaceの場合、true</returns>
        ''' <remarks></remarks>
        Private Shared Function IsWhitespace(ByVal c As Char) As Boolean
            Return 0 <= WHITE_SPACE.IndexOf(c)
        End Function

        ''' <summary>
        ''' 数式を変数名や等号不等号で区切った List にして返す
        ''' </summary>
        ''' <param name="formula">数式</param>
        ''' <returns>区切った List</returns>
        ''' <remarks></remarks>
        Public Shared Function SplitFormula(ByVal formula As String) As List(Of String)
            Const EQ As Char = "="c
            Const LT As Char = "<"c
            Const GT As Char = ">"c
            Const DOT As Char = "."c
            Dim drafts As New List(Of String)
            Dim literal As New StringBuilder
            Dim expressionChars As Char() = (formula & " ").ToCharArray
            For i As Integer = 0 To expressionChars.Length - 1
                Dim c As Char = expressionChars(i)
                If 0 <= "!=<>()+-*/%.,".IndexOf(c) Then
                    If 0 < literal.Length Then
                        If c = DOT AndAlso IsNumeric(literal.ToString) Then
                            literal.Append(c)
                            Continue For
                        End If
                        drafts.Add(literal.ToString)
                        literal.Length = 0 ' StringBuilderのクリア
                    End If
                    Dim draftsCount As Integer = drafts.Count
                    If 0 < draftsCount AndAlso ((EQ = c AndAlso 0 <= "!=<>".IndexOf(drafts(draftsCount - 1).Chars(0))) _
                                                OrElse (GT = c AndAlso LT = drafts(draftsCount - 1).Chars(0))) AndAlso drafts(draftsCount - 1).Length = 1 Then
                        ' != == <= >= <> は演算子と認識
                        drafts(draftsCount - 1) &= c
                    Else
                        drafts.Add(c)
                    End If
                ElseIf IsWhitespace(c) Then
                    If 0 < literal.Length Then
                        drafts.Add(literal.ToString)
                        literal.Length = 0 ' StringBuilderのクリア
                    End If
                Else
                    literal.Append(c)
                End If
            Next
            Return drafts
        End Function

        ''' <summary>
        ''' 'false and hoge' もしくは 'true or hoge' の hoge 終了時点のindexを探索して返す
        ''' </summary>
        ''' <param name="baseValue">基準値</param>
        ''' <param name="postfixTerms">後置記法[]</param>
        ''' <param name="startIndex">開始位置index</param>
        ''' <returns>見つかれば後置記法の位置index. なければ-1</returns>
        ''' <remarks></remarks>
        Public Shared Function DetectSkipIndexFromPostfix(ByVal baseValue As Boolean, ByVal postfixTerms As List(Of String), ByVal startIndex As Integer) As Integer
            Dim stack As New Stack
            stack.Push(baseValue)
            Dim result As Integer = -1
            For i As Integer = startIndex To postfixTerms.Count - 1
                Dim term As String = postfixTerms(i).ToLower
                Select Case term
                    Case "and", "or"
                        If stack.Count < 2 Then
                            ' 後置記法 [ true, false, and, hoge, ...] で startIndex=1 のとき、
                            ' 呼び出し元が [ false, hoge, ...] にまで評価すべきだから return
                            Return result
                        End If
                        stack.Pop()
                        If stack.Count = 1 Then
                            Dim leftVal As Boolean = Convert.ToBoolean(stack.Pop())
                            If "and".Equals(term) Then
                                If leftVal Then
                                    ' A and B にて A が true なら B を評価すべきだから return
                                    Return result
                                End If
                            Else
                                If Not leftVal Then
                                    ' A or B にて A が false なら B を評価すべきだから return
                                    Return result
                                End If
                            End If
                            result = i
                            stack.Push(leftVal)
                        End If

                    Case "!"
                        ' 単項目処理
                        Dim value As Object = stack.Pop()
                        If TypeOf value Is Boolean Then
                            stack.Push(Not Convert.ToBoolean(value))
                        Else
                            stack.Push(value)
                        End If

                    Case "=", "==", "!=", "<>", "()", ".", ",", "<", "<=", ">", ">=", "+", "-", "*", "/", "%"
                        ' 2項処理して、値を返すから、結果1項減る
                        stack.Pop()
                        If stack.Count = 0 OrElse TypeOf stack.Peek Is Boolean Then
                            Return result
                        End If

                    Case Else
                        stack.Push(term)

                End Select
            Next
            If stack.Count = 0 OrElse Not baseValue.Equals(stack.Peek) Then
                Return -1
            End If
            Return result
        End Function
    End Class
End Namespace