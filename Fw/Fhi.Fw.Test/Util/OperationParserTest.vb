Imports NUnit.Framework

Namespace Util
    Public MustInherit Class OperationParserTest

        <SetUp()> Public Overridable Sub SetUp()

        End Sub

        Public Class ParseTest : Inherits OperationParserTest

            <Test()> Public Sub 名称のみなら_Parent無しで_Memberになる(<Values("foo", " foo", "foo ", " foo ")> ByVal operation As String)
                Dim actual As OperationParser.Operation = OperationParser.Parse(operation)
                Assert.That(actual.Parent, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Args, [Is].Null)
            End Sub

            <Test()> Public Sub ピリオド区切りで_ParentとMemberになる()
                Dim actual As OperationParser.Operation = OperationParser.Parse("foo.bar")
                Assert.That(actual.Parent, [Is].Not.Null)
                Assert.That(actual.Parent.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Parent.Args, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("bar"))
                Assert.That(actual.Args, [Is].Null)
            End Sub

            <Test()> Public Sub マルチピリオドも_入れ子で_ParentとMemberになる()
                Dim actual As OperationParser.Operation = OperationParser.Parse("foo.bar.baz")
                Assert.That(actual.Parent, [Is].Not.Null)
                Assert.That(actual.Parent.Parent, [Is].Not.Null)
                Assert.That(actual.Parent.Parent.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Parent.Parent.Args, [Is].Null)
                Assert.That(actual.Parent.Member, [Is].EqualTo("bar"))
                Assert.That(actual.Parent.Args, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("baz"))
                Assert.That(actual.Args, [Is].Null)
            End Sub

            <Test()> Public Sub 空の括弧があると_引数は長さ0の配列になる(<Values("foo()", "foo( )")> ByVal operation As String)
                Dim actual As OperationParser.Operation = OperationParser.Parse(operation)
                Assert.That(actual.Parent, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Args, [Is].Not.Null)
                Assert.That(actual.Args.Length, [Is].EqualTo(0), "空の括弧があると長さ0の配列になる")
            End Sub

            <Test()> Public Sub 引数も_Memberになる(<Values("foo(bar)", "foo( bar)", "foo(bar )", "foo( bar )")> ByVal operation As String)
                Dim actual As OperationParser.Operation = OperationParser.Parse(operation)
                Assert.That(actual.Parent, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Args, [Is].Not.Empty)
                Assert.That(actual.Args(0).Parent, [Is].Null)
                Assert.That(actual.Args(0).Member, [Is].EqualTo("bar"))
                Assert.That(actual.Args(0).Args, [Is].Null)
                Assert.That(actual.Args.Length, [Is].EqualTo(1))
            End Sub

            <Test()> Public Sub 複数の引数も_Memberになる(<Values("foo(bar,baz)", "foo(bar ,baz)", "foo(bar, baz)", "foo(bar , baz)")> ByVal operation As String)
                Dim actual As OperationParser.Operation = OperationParser.Parse(operation)
                Assert.That(actual.Parent, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Args, [Is].Not.Empty)
                Assert.That(actual.Args(0).Parent, [Is].Null)
                Assert.That(actual.Args(0).Member, [Is].EqualTo("bar"))
                Assert.That(actual.Args(0).Args, [Is].Null)
                Assert.That(actual.Args(1).Parent, [Is].Null)
                Assert.That(actual.Args(1).Member, [Is].EqualTo("baz"))
                Assert.That(actual.Args(1).Args, [Is].Null)
                Assert.That(actual.Args.Length, [Is].EqualTo(2))
            End Sub

            <Test()> Public Sub 複数の引数も_Memberになる_引数3つ()
                Dim actual As OperationParser.Operation = OperationParser.Parse("foo( bar, baz, qux)")
                Assert.That(actual.Parent, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Args, [Is].Not.Empty)
                Assert.That(actual.Args(0).Parent, [Is].Null)
                Assert.That(actual.Args(0).Member, [Is].EqualTo("bar"))
                Assert.That(actual.Args(0).Args, [Is].Null)
                Assert.That(actual.Args(1).Parent, [Is].Null)
                Assert.That(actual.Args(1).Member, [Is].EqualTo("baz"))
                Assert.That(actual.Args(1).Args, [Is].Null)
                Assert.That(actual.Args(2).Parent, [Is].Null)
                Assert.That(actual.Args(2).Member, [Is].EqualTo("qux"))
                Assert.That(actual.Args(2).Args, [Is].Null)
                Assert.That(actual.Args.Length, [Is].EqualTo(3))
            End Sub

            <Test()> Public Sub 引数はMemberのみだけど_呼び出し元はParentとMemberになる()
                Dim actual As OperationParser.Operation = OperationParser.Parse("foo.bar(baz)")
                Assert.That(actual.Parent, [Is].Not.Null)
                Assert.That(actual.Parent.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Parent.Args, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("bar"))
                Assert.That(actual.Args, [Is].Not.Empty)
                Assert.That(actual.Args(0).Parent, [Is].Null)
                Assert.That(actual.Args(0).Member, [Is].EqualTo("baz"))
                Assert.That(actual.Args(0).Args, [Is].Null)
            End Sub

            <Test()> Public Sub 複数の引数はMemberのみだけど_呼び出し元はParentとMemberになる()
                Dim actual As OperationParser.Operation = OperationParser.Parse("foo.bar(baz,qux)")
                Assert.That(actual.Parent, [Is].Not.Null)
                Assert.That(actual.Parent.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Parent.Args, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("bar"))
                Assert.That(actual.Args, [Is].Not.Empty)
                Assert.That(actual.Args(0).Parent, [Is].Null)
                Assert.That(actual.Args(0).Member, [Is].EqualTo("baz"))
                Assert.That(actual.Args(0).Args, [Is].Null)
                Assert.That(actual.Args(1).Parent, [Is].Null)
                Assert.That(actual.Args(1).Member, [Is].EqualTo("qux"))
                Assert.That(actual.Args(1).Args, [Is].Null)
                Assert.That(actual.Args.Length, [Is].EqualTo(2))
            End Sub

            <Test()> Public Sub 引数も入れ子で_ParentとMemberになる()
                Dim actual As OperationParser.Operation = OperationParser.Parse("foo.bar(baz.qux)")
                Assert.That(actual.Parent, [Is].Not.Null)
                Assert.That(actual.Parent.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Parent.Args, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("bar"))
                Assert.That(actual.Args(0).Parent, [Is].Not.Null)
                Assert.That(actual.Args(0).Parent.Member, [Is].EqualTo("baz"))
                Assert.That(actual.Args(0).Parent.Args, [Is].Null)
                Assert.That(actual.Args(0).Member, [Is].EqualTo("qux"))
                Assert.That(actual.Args(0).Args, [Is].Null)
                Assert.That(actual.Args.Length, [Is].EqualTo(1))
            End Sub

            <Test()> Public Sub ParentにMemberだけの引数があって_それとMemberになる(<Values("foo(bar).baz", "foo( bar ).baz")> ByVal operation As String)
                Dim actual As OperationParser.Operation = OperationParser.Parse(operation)
                Assert.That(actual.Parent, [Is].Not.Null)
                Assert.That(actual.Parent.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Parent.Args, [Is].Not.Null)
                Assert.That(actual.Parent.Args(0).Parent, [Is].Null)
                Assert.That(actual.Parent.Args(0).Member, [Is].EqualTo("bar"))
                Assert.That(actual.Parent.Args(0).Args, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("baz"))
                Assert.That(actual.Args, [Is].Null)
            End Sub

            <Test()> Public Sub 引数が_メソッド呼出の入れ子なら_入れ子で表現する(<Values("foo(bar(baz(qux)))", "foo( bar( baz( qux ) ) )")> ByVal operation As String)
                Dim actual As OperationParser.Operation = OperationParser.Parse(operation)
                Assert.That(actual.Parent, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Args, [Is].Not.Empty)
                Assert.That(actual.Args(0).Parent, [Is].Null)
                Assert.That(actual.Args(0).Member, [Is].EqualTo("bar"))
                Assert.That(actual.Args(0).Args, [Is].Not.Empty)
                Assert.That(actual.Args(0).Args(0).Parent, [Is].Null)
                Assert.That(actual.Args(0).Args(0).Member, [Is].EqualTo("baz"))
                Assert.That(actual.Args(0).Args(0).Args, [Is].Not.Empty)
                Assert.That(actual.Args(0).Args(0).Args(0).Parent, [Is].Null)
                Assert.That(actual.Args(0).Args(0).Args(0).Member, [Is].EqualTo("qux"))
                Assert.That(actual.Args(0).Args(0).Args(0).Args, [Is].Null)
            End Sub

            <Test()> Public Sub 引数が_メソッド呼出と_Memberなら_そうなる(<Values("foo(bar(baz),qux)", "foo( bar( baz), qux)")> ByVal operation As String)
                Dim actual As OperationParser.Operation = OperationParser.Parse(operation)
                Assert.That(actual.Parent, [Is].Null)
                Assert.That(actual.Member, [Is].EqualTo("foo"))
                Assert.That(actual.Args, [Is].Not.Empty)
                Assert.That(actual.Args(0).Parent, [Is].Null)
                Assert.That(actual.Args(0).Member, [Is].EqualTo("bar"))
                Assert.That(actual.Args(0).Args, [Is].Not.Empty)
                Assert.That(actual.Args(0).Args(0).Parent, [Is].Null)
                Assert.That(actual.Args(0).Args(0).Member, [Is].EqualTo("baz"))
                Assert.That(actual.Args(0).Args(0).Args, [Is].Null)
                Assert.That(actual.Args(1).Parent, [Is].Null)
                Assert.That(actual.Args(1).Member, [Is].EqualTo("qux"))
                Assert.That(actual.Args(1).Args, [Is].Null)
            End Sub

            <Test()> Public Sub Member無しの括弧始まりは_例外になる(<Values("(foo)", "foo((bar))")> ByVal operation As String)
                Try
                    OperationParser.Parse(operation)
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].StringStarting("Member無しの括弧始まりは処理しない"))
                End Try
            End Sub

            <Test()> Public Sub ピリオドの不正で例外になる(<Values(".foo", "foo..bar")> ByVal operation As String)
                Try
                    OperationParser.Parse(operation)
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].StringStarting("ピリオドの位置が不正です"))
                End Try
            End Sub

            <Test()> Public Sub 括弧が開きっぱなしなら_例外になる(<Values("foo(", "foo("")""", "foo(()")> ByVal operation As String)
                Try
                    OperationParser.Parse(operation)
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].StringStarting("括弧を開きっぱなし"))
                End Try
            End Sub

            <Test()> Public Sub 括弧を開いて無いのに_閉じたら_例外になる(<Values("foo)", "foo(""(""))", "foo)(")> ByVal operation As String)
                Try
                    OperationParser.Parse(operation)
                    Assert.Fail()
                Catch expected As ArgumentException
                    Assert.That(expected.Message, [Is].StringStarting("括弧を開いて無いのに閉じている"))
                End Try
            End Sub

        End Class

    End Class
End Namespace