Imports NUnit.Framework

Namespace Db.Sql

    <TestFixture()> Public MustInherit Class SqlExpressionEvaluatorTest

        Public Class [Default] : Inherits SqlExpressionEvaluatorTest

            <Test()> Public Sub paramがnullならNE_Null条件でも評価出来ないからfalse()
                Dim evaluater As New SqlExpressionEvaluator()
                Assert.IsFalse(evaluater.Evaluate("@Id != null", Nothing))
            End Sub

            <Test()> Public Sub paramがnullならE_Null条件でも評価出来ないからfalse()
                Dim evaluater As New SqlExpressionEvaluator()
                Assert.IsFalse(evaluater.Evaluate("@Id = null", Nothing))
            End Sub

            <Test()> Public Sub idがnullでNE_Null条件はfalse()
                Dim vo As New HogeVo
                Dim evaluator As New SqlExpressionEvaluator()
                Assert.IsFalse(evaluator.Evaluate("@Id != null", vo))
            End Sub

            <Test()> Public Sub idが123でNE_Null条件はtrue()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim evaluator As New SqlExpressionEvaluator()
                Assert.IsTrue(evaluator.Evaluate("@Id != null", vo))
            End Sub

            <Test()> Public Sub idがnullでE_Null条件はtrue()
                Dim vo As New HogeVo
                Dim evaluator As New SqlExpressionEvaluator()
                Assert.IsTrue(evaluator.Evaluate("@Id = null", vo))
            End Sub

            <Test()> Public Sub idが123でE_Null条件はfalse()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim evaluator As New SqlExpressionEvaluator()
                Assert.IsFalse(evaluator.Evaluate("@Id = null", vo))
            End Sub

            <Test()> Public Sub String型の値Valueが123でE_Null条件はfalse()
                Dim evaluator As New SqlExpressionEvaluator()
                Dim val As String = "123"
                Assert.IsFalse(evaluator.Evaluate("@Value = null", val))
            End Sub

            <Test()> Public Sub String型の値Valueが123でNE_Null条件はtrue()
                Dim evaluator As New SqlExpressionEvaluator()
                Dim val As String = "123"
                Assert.IsTrue(evaluator.Evaluate("@Value != null", val))
            End Sub

            <Test()> Public Sub andで二つの条件()
                Dim vo As New HogeVo
                vo.Id = 123
                vo.Name = "asd"
                Dim evaluator As New SqlExpressionEvaluator()
                Assert.IsTrue(evaluator.Evaluate("@Id != null and @Name != null", vo))
            End Sub

            <Test()> Public Sub 等符号LessThan()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim evaluator As New SqlExpressionEvaluator()
                Assert.IsTrue(evaluator.Evaluate("@Id > 99", vo))
            End Sub

        End Class

    End Class
End Namespace
