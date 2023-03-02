Imports System
Imports NUnit.Framework

Namespace Db.Sql

    <TestFixture()> Public MustInherit Class ExpressionTest

        Public Class [Default] : Inherits ExpressionTest

            <Test()> Public Sub true文字列だけならtrue()
                Dim param As Object = Nothing
                Dim exp As New Expression("true", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueAndTrueならtrue()
                Dim param As Object = Nothing
                Dim exp As New Expression("true and true", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueAndFalseならfalse()
                Dim param As Object = Nothing
                Dim exp As New Expression("true and false", param)
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueOrFalseならtrue()
                Dim param As Object = Nothing
                Dim exp As New Expression("true or false", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueOrFalseAndFalseならfalse()
                Dim param As Object = Nothing
                Dim exp As New Expression("true or false and false", param)
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueAndFalseOrTrueならtrue()
                Dim param As Object = Nothing
                Dim exp As New Expression("true and false or true", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueAndTrueOrFalseならtrue()
                Dim param As Object = Nothing
                Dim exp As New Expression("true and true or false", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub TrueAndTrueOrXxxxのXxxxはエラーでも結果はtrue()
                Dim param As Object = Nothing
                Dim exp As New Expression("true and true or @dummy", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub Booleanとの比較もできる(<Values(True, False)> ByVal value As Boolean)
                Dim exp As New Expression("@Value = False", value)
                Assert.That(exp.Evaluate, [Is].EqualTo(Not value))
            End Sub

            <Test()> Public Sub propertyNeNullで値あればtrue()
                Dim param As New HogeVo
                param.Id = 123
                Dim exp As New Expression("@Id!=null", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub
            <Test()> Public Sub propertyNeNullで値nullだからfalse()
                Dim param As New HogeVo
                Dim exp As New Expression("@Id!=null", param)
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub valueNeNullでString値123だからtrue()
                Dim param As String = "123"
                Dim exp As New Expression("@Value!=null", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub valueENullでString値123だからfalse(<Values("=", "==")> ByVal equalSign As String)
                Dim param As String = "123"
                Dim exp As New Expression("@Value" & equalSign & "null", param)
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub valueNeNullでString値nullだからfalse()
                Dim param As String = Nothing
                Dim exp As New Expression("@Value!=null", param)
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub valueENullでString値nullだからtrue(<Values("=", "==")> ByVal equalSign As String)
                Dim param As String = Nothing
                Dim exp As New Expression("@Value" & equalSign & "null", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub valueNeNullでInt値123だからtrue()
                Dim param As Integer = 123
                Dim exp As New Expression("@Value!=null", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub propertyNeNullAndPropertyNeNullで二つとも値あればtrue()
                Dim param As New HogeVo
                param.Id = 123
                param.Name = "asd"
                Dim exp As New Expression("@Id!=null and @Name!=null", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub propertyNeNullAndPropertyNeNullで片方値が無ければfalse()
                Dim param As New HogeVo
                param.Name = "asd"
                Dim exp As New Expression("@Id!=null and @Name!=null", param)
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub propertyのパラメータ値がnullだと評価しないでfalse固定()
                Dim param As HogeVo = Nothing
                Dim exp As New Expression("@Name!=null", param)
                Assert.IsFalse(exp.Evaluate)
            End Sub

            <Test()> Public Sub パラメータ名に相当するプロパティをパラメータ値が持たないと例外()
                Dim param As New HogeVo
                '' 小文字 id だから例外になる
                Dim exp As New Expression("@Nanika!=null", param)
                Try
                    exp.Evaluate()
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub idがnullでNE_Null条件はfalseになる右辺左辺入替に対応()
                Dim vo As New HogeVo
                Dim exp As New Expression("null!=@Id", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub idが123でNE_Null条件はtrueになる右辺左辺入替に対応()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim exp As New Expression("null!=@Id", vo)
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 同型プロパティ同士の比較で同値なのでtrue(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.ExName7 = "123"
                vo.Name = "123"
                Dim exp As New Expression("@ExName7" & equalSign & "@Name", vo)
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 同型プロパティ同士の比較で違う値なのでfalse(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.ExName7 = "12"
                vo.Name = "123"
                Dim exp As New Expression("@ExName7" & equalSign & "@Name", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 同型プロパティ同士の比較で片方がnothingなのでfalse(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.ExName7 = Nothing
                vo.Name = "123"
                Dim exp As New Expression("@ExName7" & equalSign & "@Name", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 同型プロパティ同士の比較で両方ともnothingなのでtrue(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.ExName7 = Nothing
                vo.Name = Nothing
                Dim exp As New Expression("@ExName7" & equalSign & "@Name", vo)
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 型違いプロパティ同士の比較(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.Id = 123
                vo.Name = "123"
                Dim exp As New Expression("@Id" & equalSign & "@Name", vo)
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_等号(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.Id = 123
                Dim exp As New Expression("@Id" & equalSign & "123", vo)
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_等号_違う値(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.Id = 123
                Dim exp As New Expression("@Id" & equalSign & "12", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_等号_自分がNull(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.Id = Nothing
                Dim exp As New Expression("@Id" & equalSign & "123", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_等号_相手がNull(<Values("=", "==")> ByVal equalSign As String)
                Dim vo As New HogeVo
                vo.Id = 123
                Dim exp As New Expression("@Id" & equalSign & "null", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_LT()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As New Expression("@Id>11", vo)
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_LT_境界false()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As New Expression("@Id>12", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_LE()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As New Expression("@Id>=12", vo)
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_LE_境界false()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As New Expression("@Id>=12.1", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_GT()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As New Expression("@Id<12.1", vo)
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_GT_境界false()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As New Expression("@Id<12", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_GE()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As New Expression("@Id<=12", vo)
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_GE_境界false()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As New Expression("@Id<=11.9", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_文字列比較ではないことの確認()
                Dim vo As New HogeVo
                vo.Id = 12
                Dim exp As New Expression("@Id>9", vo)
                Assert.IsTrue(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_自分がNull()
                Dim vo As New HogeVo
                vo.Id = Nothing
                Dim exp As New Expression("@Id>123", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub 数の比較_不等号_相手がNull()
                Dim vo As New HogeVo
                vo.Id = 123
                Dim exp As New Expression("@Id>null", vo)
                Assert.IsFalse(exp.Evaluate())
            End Sub

            <Test()> Public Sub Null値のプロパティを参照したらFalseになる()
                Dim vo As New FugaVo
                Dim exp As New Expression("0 < @Name.Length or 0 < @StrArray.Length", vo)
                Assert.That(exp.Evaluate, [Is].False)
            End Sub

            <Test()> Public Sub Null値のプロパティを参照したらFalseになる_論理演算なら_一つでも評価できればtrueになる()
                Dim vo As New FugaVo With {.StrArray = New String() {"Z"}}
                Dim exp As New Expression("0 < @Name.Length or 0 < @StrArray.Length", vo)
                Assert.That(exp.Evaluate, [Is].True)
            End Sub

        End Class

        Public Class エラー制御Test : Inherits ExpressionTest

            <Test()> Public Sub 四則演算と剰余は2項とも数値以外はエラー(<Values("*", "/", "%", "-")> ByVal [operator] As String)
                Dim exp As New Expression("10" & [operator] & "A", Nothing)
                Try
                    exp.Calculate()
                    Assert.Fail()
                Catch expected As Expression.IllegalArithmeticException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub 四則演算と剰余は2項とも数値以外はエラー_加算だけは文字列として結合動作する()
                Dim exp As New Expression("10+A", Nothing)
                Assert.That(exp.Calculate, [Is].EqualTo("10A"), "文字列として結合する")
            End Sub

            <Test()> Public Sub Null値のプロパティを参照したら強制的にFalseになる_左辺(<Values("=", "!=", "<>", "<", "<=", ">", ">=")> ByVal [operator] As String)
                Dim vo As New FugaVo
                Dim exp As New Expression("@Name.Length " & [operator] & " null", vo)
                Assert.That(exp.Calculate(), [Is].False)
            End Sub

            <Test()> Public Sub Null値のプロパティを参照したら強制的にFalseになる_右辺(<Values("=", "!=", "<>", "<", "<=", ">", ">=")> ByVal [operator] As String)
                Dim vo As New FugaVo
                Dim exp As New Expression("null " & [operator] & " @Name.Length", vo)
                Assert.That(exp.Calculate(), [Is].False)
            End Sub

            <Test()> Public Sub 四則演算でNull値のプロパティを参照したらエラー_左辺(<Values("+", "-", "*", "/", "%")> ByVal [operator] As String)
                Dim vo As New FugaVo
                Dim exp As New Expression("@Name.Length " & [operator] & " 234", vo)
                Try
                    exp.Calculate()
                    Assert.Fail()
                Catch expected As Expression.IllegalArithmeticException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub 四則演算でNull値のプロパティを参照したらエラー_右辺(<Values("+", "-", "*", "/", "%")> ByVal [operator] As String)
                Dim vo As New FugaVo
                Dim exp As New Expression("123 " & [operator] & " @Name.Length", vo)
                Try
                    exp.Calculate()
                    Assert.Fail()
                Catch expected As Expression.IllegalArithmeticException
                    Assert.IsTrue(True)
                End Try
            End Sub

            <Test()> Public Sub 四則演算でNull値を参照したらエラー(<Values("+", "-", "*", "/", "%")> ByVal [operator] As String)
                Dim vo As New FugaVo
                Dim exp As New Expression("@Name " & [operator] & " 234", vo)
                Try
                    exp.Calculate()
                    Assert.Fail()
                Catch expected As Expression.IllegalArithmeticException
                    Assert.IsTrue(True)
                End Try
            End Sub

        End Class

        Public Class 論理演算子_And_Or_の左辺値によって右辺の評価をSkipするTest : Inherits ExpressionTest

            <Test()> Public Sub And演算子の左辺がfalseだから右辺は評価しないでfalse()
                Dim param As String = "A"
                Dim exp As New Expression("false and @Hoge!=null", param)
                Assert.IsFalse(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub And演算子の左辺がfalseだから右辺は評価しないでfalse_右辺にorを追加()
                Dim param As String = "A"
                Dim exp As New Expression("false and @Hoge!=null or true", param)
                Assert.IsTrue(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub And演算子の左辺がfalseだから右辺は評価しないでfalse_左辺にandを追加()
                Dim param As String = "A"
                Dim exp As New Expression("true and false and @Hoge!=null", param)
                Assert.IsFalse(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub And演算子の左辺がfalseだから右辺は評価しないでfalse_右辺にandを追加()
                Dim param As String = "A"
                Dim exp As New Expression("false and @Hoge and @Fuga", param)
                Assert.IsFalse(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub And演算子の左辺がfalseだから右辺は評価しないでfalse_右辺は不正なプロパティ()
                Dim param As String = "A"
                Dim exp As New Expression("true and (false and @Hoge)", param)
                Assert.IsFalse(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub And演算子の左辺がfalseだから右辺は評価しないでfalse_右辺は不正なプロパティ2()
                Dim param As String = "A"
                Dim exp As New Expression("false and (@Hoge and @Fuga)", param)
                Assert.IsFalse(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub And演算子の左辺がfalseだから右辺は評価しないでfalse_式が複雑()
                Dim param As String = "A"
                Dim exp As New Expression("false and !@Hoge(0)=1*(2+3)", param)
                Assert.IsFalse(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub And演算子の左辺がfalseだから右辺は評価しないでfalse_左辺埋め込み変数()
                Dim param As Boolean = False
                Dim exp As New Expression("@Value and @Hoge!=null", param)
                Assert.IsFalse(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub And演算子の左辺がNotTrueだから右辺は評価しないでfalse()
                Dim param As String = "A"
                Dim exp As New Expression("!true and @Hoge!=null", param)
                Assert.IsFalse(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub And演算子の左辺がNotTrueだから右辺は評価しないでfalse_左辺にandを追加()
                Dim param As String = "A"
                Dim exp As New Expression("true and !true and @Hoge!=null", param)
                Assert.IsFalse(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub And演算子の左辺がNotFalseなら右辺は評価する()
                Dim param As Integer = 12
                Dim exp As New Expression("!false and 12=@Value", param)
                Assert.IsTrue(exp.Evaluate)
            End Sub

            <Test()> Public Sub Or演算子の左辺がtrueだから右辺は評価しないでtrue()
                Dim param As String = "A"
                Dim exp As New Expression("true or @Hoge!=null", param)
                Assert.IsTrue(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub Or演算子の左辺がtrueだから右辺は評価しないでtrue_右辺にandを追加()
                Dim param As String = "A"
                Dim exp As New Expression("true or @Hoge!=null and false", param)
                Assert.IsFalse(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub Or演算子の左辺がtrueだから右辺は評価しないでtrue_式が複雑()
                Dim param As String = "A"
                Dim exp As New Expression("true or !@Hoge(0)==!Fuga(2).Piyo(4)", param)
                Assert.IsTrue(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub Or演算子の左辺がtrueだから右辺は評価しないでtrue_左辺埋め込み変数()
                Dim param As Boolean = True
                Dim exp As New Expression("@Value or @Hoge!=null", param)
                Assert.IsTrue(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub Or演算子の左辺がNotFalseだから右辺は評価しないでtrue()
                Dim param As String = "A"
                Dim exp As New Expression("!false or @Hoge!=null", param)
                Assert.IsTrue(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub Or演算子の左辺がNotFalseだから右辺は評価しないでtrue_左辺にorを追加()
                Dim param As String = "A"
                Dim exp As New Expression("false or !false or @Hoge!=null", param)
                Assert.IsTrue(exp.Evaluate, "String型にHogeというプロパティはないから評価しようとするとエラー")
            End Sub

            <Test()> Public Sub Or演算子の左辺がNotTrueだから右辺は評価する()
                Dim param As Integer = 10
                Dim exp As New Expression("!true or 12+3=@Value", param)
                Assert.IsFalse(exp.Evaluate)
            End Sub

            Private Class HogeVo
                Private _isHoge As Boolean?
                Private _isFuga As Boolean?

                Public Property IsHoge() As Boolean?
                    Get
                        Return _isHoge
                    End Get
                    Set(ByVal value As Boolean?)
                        _isHoge = value
                    End Set
                End Property

                Public Property IsFuga() As Boolean?
                    Get
                        Return _isFuga
                    End Get
                    Set(ByVal value As Boolean?)
                        _isFuga = value
                    End Set
                End Property
            End Class

            <Test()> Public Sub AandBの_Aが正しく評価されてtrueだから_Bの評価が結果となる(<Values(True, False)> ByVal isFuga As Boolean)
                Dim param As New HogeVo With {.IsHoge = False, .IsFuga = isFuga}
                Dim exp As New Expression("@IsHoge = false and @IsFuga", param)
                Assert.That(exp.Evaluate, [Is].EqualTo(isFuga))
            End Sub

            <Test()> Public Sub AorBの_Aが評価されてfalseだから_Bの評価が結果となる(<Values(True, False)> ByVal isFuga As Boolean)
                Dim param As New HogeVo With {.IsHoge = True, .IsFuga = isFuga}
                Dim exp As New Expression("@IsHoge = false or @IsFuga", param)
                Assert.That(exp.Evaluate, [Is].EqualTo(isFuga))
            End Sub

            <Test()> Public Sub And演算子の左辺がfalseだから右辺は評価しないでfalse_右辺は不正なプロパティ3()
                Dim param As New HogeVo With {.IsHoge = True}
                Dim exp As New Expression("@IsHoge = false and @IsPiyo", param)
                Assert.That(exp.Evaluate, [Is].False)
            End Sub

            <Test()> Public Sub Or演算子の左辺がtrueだから右辺は評価しないでtrue_右辺は不正なプロパティ()
                Dim param As New HogeVo With {.IsHoge = False}
                Dim exp As New Expression("@IsHoge = false or @IsPiyo", param)
                Assert.That(exp.Evaluate, [Is].True)
            End Sub

        End Class

        Public Class 関数とメンバの組み合わせで評価するTest : Inherits ExpressionTest

            <Test()> Public Sub 文字列配列の要素を指定して評価できる(<Values(0, 1, 2)> ByVal index As Integer)
                Dim names As String() = {"a", "b", "c"}
                Dim exp As New Expression("@Value(" & index & ")=" & names(index), names)
                Assert.That(exp.Evaluate, [Is].True)
            End Sub

            <Test()> Public Sub 文字列配列の要素を指定して評価できる_不一致(<Values(0, 1, 2)> ByVal index As Integer)
                Dim names As String() = {"a", "b", "c"}
                Dim exp As New Expression("@Value(" & index & ")=" & names((index + 1) Mod 3), names)
                Assert.That(exp.Evaluate, [Is].False)
            End Sub

            <Test()> Public Sub VoのInt配列の要素を指定して評価できる(<Values(0, 1, 2)> ByVal index As Integer)
                Dim vo As New FugaVo With {.IntArray = New Integer() {3, 5, 8}}
                Dim exp As New Expression("@IntArray(" & index & ")=" & vo.IntArray(index), vo)
                Assert.That(exp.Evaluate, [Is].True)
            End Sub

            <Test()> Public Sub VoのHogeVo配列の要素を指定して評価できる(<Values(0, 1, 2)> ByVal index As Integer)
                Dim vo As New FugaVo With {.HogeArray = New HogeVo() {New HogeVo With {.Id = 2}, New HogeVo With {.Id = 3}, New HogeVo With {.Id = 5}}}
                Dim exp As New Expression("@HogeArray(" & index & ").Id=" & vo.HogeArray(index).Id, vo)
                Assert.That(exp.Evaluate, [Is].True)
            End Sub

            <Test()> Public Sub Voの中のVo配列の要素を指定して評価できる(<Values(0, 1, 2)> ByVal index As Integer)
                Dim vo As New FugaVo With {.FugaArray = New FugaVo() {New FugaVo() With {.HogeArray = New HogeVo() {New HogeVo With {.Id = 2}, New HogeVo With {.Id = 3}, New HogeVo With {.Id = 5}}}}}
                Dim exp As New Expression("@FugaArray(0).HogeArray(" & index & ").Id=" & vo.FugaArray(0).HogeArray(index).Id, vo)
                Assert.That(exp.Evaluate, [Is].True)
            End Sub

            <Test()> Public Sub 指定要素がnullなのにプロパティアクセスしてるから常にfalse()
                Dim vo As New FugaVo With {.HogeArray = New HogeVo() {Nothing}}
                Dim exp As New Expression("@HogeArray(0).Id=null", vo)
                Assert.That(exp.Evaluate, [Is].False)
            End Sub

        End Class

    End Class
End Namespace
