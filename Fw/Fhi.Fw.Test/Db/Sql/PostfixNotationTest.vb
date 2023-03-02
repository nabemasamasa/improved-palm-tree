Imports NUnit.Framework

Namespace Db.Sql
    Public MustInherit Class PostfixNotationTest

        <SetUp()> Public Overridable Sub SetUp()

        End Sub

        Public Class SplitFormula_数式を演算子とオペランドに分割する : Inherits PostfixNotationTest

            <Test()> Public Sub 認識している演算子(<Values("=", "!=", "==", ">", ">=", "<", "<=", "<>")> ByVal [operator] As String)
                Dim actuals As List(Of String) = PostfixNotation.SplitFormula("@Id" & [operator] & "null")
                Assert.That(actuals(0), [Is].EqualTo("@Id"))
                Assert.That(actuals(1), [Is].EqualTo([operator]))
                Assert.That(actuals(2), [Is].EqualTo("null"))
                Assert.AreEqual(3, actuals.Count)
            End Sub

            <Test()> Public Sub 認めない演算子の組み合わせ(<Values("><", "=>", "=<", "<<", ">>", "=!", "!!", ")(", "(=", ")=")> ByVal [operator] As String)
                Dim actuals As List(Of String) = PostfixNotation.SplitFormula("@Id" & [operator] & "null")
                Assert.That(actuals(0), [Is].EqualTo("@Id"))
                Assert.That(actuals(1), [Is].EqualTo([operator].Substring(0, 1)), "別々の演算子として認識する")
                Assert.That(actuals(2), [Is].EqualTo([operator].Substring(1, 1)), "別々の演算子として認識する")
                Assert.That(actuals(3), [Is].EqualTo("null"))
                Assert.AreEqual(4, actuals.Count)
            End Sub

            <Test()> Public Sub 小数点のドットは演算子扱いしない(<Values("12.34", "1.0", "33.", "0.5")> ByVal value As String)
                Dim results As List(Of String) = PostfixNotation.SplitFormula("@Id<" & value)
                Assert.That(results(0), [Is].EqualTo("@Id"))
                Assert.That(results(1), [Is].EqualTo("<"))
                Assert.That(results(2), [Is].EqualTo(value))
                Assert.That(results.Count, [Is].EqualTo(3))
            End Sub

            <Test()> Public Sub スペース区切り()
                Dim results As List(Of String) = PostfixNotation.SplitFormula("@Id != null")
                Assert.AreEqual(3, results.Count)
                Assert.AreEqual("@Id", results(0))
                Assert.AreEqual("!=", results(1))
                Assert.AreEqual("null", results(2))
            End Sub

            <Test()> Public Sub スペース区切り_Andで二つの条件()
                Dim actuals As List(Of String) = PostfixNotation.SplitFormula("@Id != null and @Name != null")
                Assert.That(actuals(0), [Is].EqualTo("@Id"))
                Assert.That(actuals(1), [Is].EqualTo("!="))
                Assert.That(actuals(2), [Is].EqualTo("null"))
                Assert.That(actuals(3), [Is].EqualTo("and"))
                Assert.That(actuals(4), [Is].EqualTo("@Name"))
                Assert.That(actuals(5), [Is].EqualTo("!="))
                Assert.That(actuals(6), [Is].EqualTo("null"))
                Assert.That(actuals.Count, [Is].EqualTo(7))
            End Sub

            <Test()> Public Sub スペース区切り_少し複雑()
                Dim actuals As List(Of String) = PostfixNotation.SplitFormula("@FugaArray(0).FugaArray(0).StrArray != null and 1 < @FugaArray(0).FugaArray(0).StrArray.Length")
                Assert.That(Join(actuals.ToArray, " "), [Is].EqualTo("@FugaArray ( 0 ) . FugaArray ( 0 ) . StrArray != null and 1 < @FugaArray ( 0 ) . FugaArray ( 0 ) . StrArray . Length"))
            End Sub

            <Test()> Public Sub Bunkai_スペース区切り()
                Dim results As List(Of String) = SqlExpressionEvaluator.Bunkai("@FugaArray(0).FugaArray(0).StrArray != null and 1 < @FugaArray(0).FugaArray(0).StrArray.Length")
                Assert.AreEqual("@FugaArray(0).FugaArray(0).StrArray", results(0))
                Assert.AreEqual("!=", results(1))
                Assert.AreEqual("null", results(2))
                Assert.AreEqual("and", results(3))
                Assert.AreEqual("1", results(4))
                Assert.AreEqual("<", results(5))
                Assert.AreEqual("@FugaArray(0).FugaArray(0).StrArray.Length", results(6))
                Assert.AreEqual(7, results.Count)
            End Sub

        End Class

        Public Class MakePostfix_ : Inherits PostfixNotationTest

            <Test(), Sequential()> Public Sub 四則演算( _
                    <Values("a+b", "a+b+c", "a+b*c", "(a+b)*c", "a*b+c", "a*(b+c)")> ByVal formula As String, _
                    <Values("a b +", "a b + c +", "a b c * +", "a b + c *", "a b * c +", "a b c + *")> ByVal expected As String)
                Dim actuals As List(Of String) = PostfixNotation.MakePostfix(formula)
                Assert.That(Join(actuals.ToArray, " "), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub 論理演算( _
                    <Values("a=b", "a=b and c", "a and b=c", "a=(b and c)", "a and (b=c)", "!a", "!(a and b) and c", "a and !(b and c)")> ByVal formula As String, _
                    <Values("a b =", "a b = c and", "a b c = and", "a b c and =", "a b c = and", "a !", "a b and ! c and", "a b c and ! and")> ByVal expected As String)
                Dim actuals As List(Of String) = PostfixNotation.MakePostfix(formula)
                Assert.That(Join(actuals.ToArray, " "), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub 関数呼出し( _
                    <Values("a(b)", "a (b)", "a()", "a ()", "a(b(c))", "a(b)(c)")> ByVal formula As String, _
                    <Values("a b ()", "a b ()", "a \ ()", "a \ ()", "a b c () ()", "a b () c ()")> ByVal expected As String)
                ' a (b) なら a b にしたいけど、未対応
                ' a ()  ならエラーにしたいけど、未対応
                Dim actuals As List(Of String) = PostfixNotation.MakePostfix(formula)
                Assert.That(Join(actuals.ToArray, " "), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub メンバへアクセス( _
                    <Values("a.b", "a .b", "a. b", "a . b")> ByVal formula As String, _
                    <Values("a b .", "a b .", "a b .", "a b .")> ByVal expected As String)
                ' a .b ならエラーにしたいけど、未対応
                ' a. b ならエラーにしたいけど、未対応
                Dim actuals As List(Of String) = PostfixNotation.MakePostfix(formula)
                Assert.That(Join(actuals.ToArray, " "), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub 関数とメンバは同順( _
                    <Values("a(b).c(d)", "a(b).c(d).e")> ByVal formula As String, _
                    <Values("a b () c . d ()", "a b () c . d () e .")> ByVal expected As String)
                Dim actuals As List(Of String) = PostfixNotation.MakePostfix(formula)
                Assert.That(Join(actuals.ToArray, " "), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub 否定は関数とメンバの次順位( _
                    <Values("!a.b", "!a(b)", "!a(b).c", "!a(b).c(d)")> ByVal formula As String, _
                    <Values("a b . !", "a b () !", "a b () c . !", "a b () c . d () !")> ByVal expected As String)
                Dim actuals As List(Of String) = PostfixNotation.MakePostfix(formula)
                Assert.That(Join(actuals.ToArray, " "), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub コンマ演算子( _
                    <Values("a(b,c)", "a(b+c,d)")> ByVal formula As String, _
                    <Values("a b c , ()", "a b c + d , ()")> ByVal expected As String)
                Dim actuals As List(Of String) = PostfixNotation.MakePostfix(formula)
                Assert.That(Join(actuals.ToArray, " "), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub 構文エラーなら例外( _
                    <Values("(a=b", "a(=b", "a=(b", "a=b(", ")a=b", "a)=b", "a=)b", "a=b)", _
                            "a(b)c", "a(b))")> ByVal formula As String)
                Try
                    Dim postfixTerms As List(Of String) = PostfixNotation.MakePostfix(formula)
                    Assert.Fail("構文エラーなのに例外にならず " & Join(postfixTerms.ToArray, " "))
                Catch ex As PostfixNotation.IllegalFormulaTermException
                    Assert.IsTrue(True)
                Catch ex As PostfixNotation.IllegalFormulaEndException
                    CollectionAssert.Contains(New String() {"a=b("}, formula)
                Catch ex As PostfixNotation.InvalidBracketException
                    CollectionAssert.Contains(New String() {"(a=b", "a(=b", "a=(b"}, formula)
                End Try
            End Sub

        End Class

    End Class
End Namespace
