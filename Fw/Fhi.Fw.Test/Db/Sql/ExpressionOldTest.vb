Imports System
Imports NUnit.Framework

Namespace Db.Sql

    <TestFixture()> Public MustInherit Class ExpressionOldTest

        Public Class [Default] : Inherits ExpressionOldTest

            <Test()> Public Sub true文字列だけならtrue()
                Dim param As Object = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("true"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueAndTrueならtrue()
                Dim param As Object = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("true", "and", "true"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueAndFalseならfalse()
                Dim param As Object = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("true", "and", "false"))
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueOrFalseならtrue()
                Dim param As Object = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("true", "or", "false"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueOrFalseAndFalseならfalse()
                Dim param As Object = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("true", "or", "false", "and", "false"))
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueAndFalseOrTrueならtrue()
                Dim param As Object = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("true", "and", "false", "or", "true"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueAndTrueOrFalseならtrue()
                Dim param As Object = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("true", "and", "true", "or", "false"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueAndTrueOrXxxxのXxxxはエラーでも結果はtrue()
                Dim param As Object = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("true", "and", "true", "or", "@dummy"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub propertyNeNullで値あればtrue()
                Dim param As New HogeVo
                param.Id = 123
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Id", "!=", "null"))
                Assert.IsTrue(exp.Evaluate)
            End Sub
            <Test()> Public Sub propertyNeNullで値nullだからfalse()
                Dim param As New HogeVo
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Id", "!=", "null"))
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub valueNeNullでString値123だからtrue()
                Dim param As String = "123"
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Value", "!=", "null"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub valueENullでString値123だからfalse(<Values("=", "==")> ByVal equalSign As String)
                Dim param As String = "123"
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Value", equalSign, "null"))
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub valueNeNullでString値nullだからfalse()
                Dim param As String = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Value", "!=", "null"))
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub valueENullでString値nullだからtrue(<Values("=", "==")> ByVal equalSign As String)
                Dim param As String = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Value", equalSign, "null"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub valueNeNullでInt値123だからtrue()
                Dim param As Integer = 123
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Value", "!=", "null"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub propertyNeNullAndPropertyNeNullで二つとも値あればtrue()
                Dim param As New HogeVo
                param.Id = 123
                param.Name = "asd"
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Id", "!=", "null", "and", "@Name", "!=", "null"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub propertyNeNullAndPropertyNeNullで片方値が無ければfalse()
                Dim param As New HogeVo
                param.Name = "asd"
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Id", "!=", "null", "and", "@Name", "!=", "null"))
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub propertyのパラメータ値がnullだと評価しないでfalse固定()
                Dim param As HogeVo = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Name", "!=", "null"))
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub パラメータ名に相当するプロパティをパラメータ値が持たないと例外()
                Dim param As New HogeVo
                '' 小文字 id だから例外になる
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("@Nanika", "!=", "null"))
                Try
                    exp.Evaluate()
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub And演算子の左辺がfalseだから右辺は評価しないでfalse()
                Dim param As HogeVo = Nothing
                ' nothing だから右辺は評価しようにも出来ない
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("false", "and", "@Name", "!=", "null"))
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub Or演算子の左辺がtrueだから右辺は評価しないでtrue()
                Dim param As HogeVo = Nothing
                ' nothing だから右辺は評価しようにも出来ない
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(param, StringUtil.ToList("true", "or", "@Name", "!=", "null"))
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub idがnullでNE_Null条件はfalseになる右辺左辺入替に対応()
                Dim vo As New HogeVo
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("null", "!=", "@Id"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub idが123でNE_Null条件はtrueになる右辺左辺入替に対応()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("null", "!=", "@Id"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 同型プロパティ同士の比較で同値なのでtrue(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.ExName7 = "123"
                vo.Name = "123"
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@ExName7", equalSign, "@Name"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 同型プロパティ同士の比較で違う値なのでfalse(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.ExName7 = "12"
                vo.Name = "123"
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@ExName7", equalSign, "@Name"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 同型プロパティ同士の比較で片方がnothingなのでfalse(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.ExName7 = Nothing
                vo.Name = "123"
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@ExName7", equalSign, "@Name"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 同型プロパティ同士の比較で両方ともnothingなのでtrue(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.ExName7 = Nothing
                vo.Name = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@ExName7", equalSign, "@Name"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 型違いプロパティ同士の比較(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.Id = 123
                vo.Name = "123"
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", equalSign, "@Name"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_等号(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.Id = 123
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", equalSign, "123"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_等号_違う値(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.Id = 123
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", equalSign, "12"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_等号_自分がNull(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.Id = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", equalSign, "123"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_等号_相手がNull(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.Id = 123
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", equalSign, "null"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_LT()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", ">", "11"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_LT_境界false()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", ">", "12"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_LE()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", ">=", "12"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_LE_境界false()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", ">=", "12.1"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_GT()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", "<", "12.1"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_GT_境界false()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", "<", "12"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_GE()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", "<=", "12"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_GE_境界false()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", "<=", "11.9"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_文字列比較ではないことの確認()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", ">", "9"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_自分がNull()
                Dim vo As New HogeVo
                vo.Id = Nothing
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", ">", "123"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_相手がNull()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("@Id", ">", "null"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub Null値のプロパティを参照したらFalseになる()
                Dim vo As New FugaVo
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("0", "<", "@Name.Length", "or", "0", "<", "@StrArray.Length"))
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub Null値のプロパティを参照したらFalseになる_論理演算なら_一つでも評価できればtrueになる()
                Dim vo As New FugaVo With {.StrArray = New String() {"Z"}}
                Dim exp As ExpressionOld = ExpressionOld.NewInstance(vo, StringUtil.ToList("0", "<", "@Name.Length", "or", "0", "<", "@StrArray.Length"))
                Assert.IsTrue(exp.Evaluate())
            End Sub

        End Class

    End Class
End Namespace
