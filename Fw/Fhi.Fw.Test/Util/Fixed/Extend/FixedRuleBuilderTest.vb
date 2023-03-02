Imports NUnit.Framework

Namespace Util.Fixed.Extend
    Public Class FixedRuleBuilderTest

#Region "Testing Fixed Classes"

        Private Class Parent
            Private _num As String
            Private _children As Child()
            Private _cost As Decimal?

            Public Property Num() As String
                Get
                    Return _num
                End Get
                Set(ByVal value As String)
                    _num = value
                End Set
            End Property

            Public Property Children() As Child()
                Get
                    Return _children
                End Get
                Set(ByVal value As Child())
                    _children = value
                End Set
            End Property

            Public Property Cost() As Decimal?
                Get
                    Return _cost
                End Get
                Set(ByVal value As Decimal?)
                    _cost = value
                End Set
            End Property
        End Class

        Private Class ParentOneChild
            Private _num As String
            Private _child As Child
            Private _cost As Decimal?

            Public Property Num() As String
                Get
                    Return _num
                End Get
                Set (ByVal value As String)
                    _num = value
                End Set
            End Property

            Public Property Child() As Child
                Get
                    Return _child
                End Get
                Set (ByVal value As Child)
                    _child = value
                End Set
            End Property

            Public Property Cost() As Decimal?
                Get
                    Return _cost
                End Get
                Set (ByVal value As Decimal?)
                    _cost = value
                End Set
            End Property
        End Class

        Private Class Child
            Private _name As String
            Private _children As GrandChild()
            Private _second As String

            Public Property Name() As String
                Get
                    Return _name
                End Get
                Set(ByVal value As String)
                    _name = value
                End Set
            End Property

            Public Property Children() As GrandChild()
                Get
                    Return _children
                End Get
                Set(ByVal value As GrandChild())
                    _children = value
                End Set
            End Property

            Public Property Second() As String
                Get
                    Return _second
                End Get
                Set(ByVal value As String)
                    _second = value
                End Set
            End Property
        End Class

        Private Class Child2
            Private _first As String
            Private _second As String
            Private _third As String

            Public Property First() As String
                Get
                    Return _first
                End Get
                Set(ByVal value As String)
                    _first = value
                End Set
            End Property

            Public Property Second() As String
                Get
                    Return _second
                End Get
                Set(ByVal value As String)
                    _second = value
                End Set
            End Property

            Public Property Third() As String
                Get
                    Return _third
                End Get
                Set(ByVal value As String)
                    _third = value
                End Set
            End Property
        End Class

        Private Class GrandChild
            Private _weight As Decimal?

            Public Property Weight() As Decimal?
                Get
                    Return _weight
                End Get
                Set(ByVal value As Decimal?)
                    _weight = value
                End Set
            End Property
        End Class
#End Region

        Private Class ParentRuleImpl : Implements IFixedRule(Of Parent)

            Public Sub Configure(ByVal defineBy As IFixedRuleLocator, ByVal vo As Parent) Implements IFixedRule(Of Parent).Configure
                defineBy.Hankaku(vo.Num, 3).GroupRepeat(vo.Children, New ChildRuleImpl, 3).Number(vo.Cost, 8, 0)
            End Sub
        End Class

        Private Class ChildRuleImpl : Implements IFixedRule(Of Child)

            Public Sub Configure(ByVal defineBy As IFixedRuleLocator, ByVal vo As Child) Implements IFixedRule(Of Child).Configure
                defineBy.Zenkaku(vo.Name, 8).GroupRepeat(vo.Children, New GrandChildRuleImpl, 2).Zenkaku(vo.Second, 4)
            End Sub
        End Class

        Private Class GrandChildRuleImpl : Implements IFixedRule(Of GrandChild)

            Public Sub Configure(ByVal defineBy As IFixedRuleLocator, ByVal vo As GrandChild) Implements IFixedRule(Of GrandChild).Configure
                defineBy.Number(vo.Weight, 10, 3)
            End Sub
        End Class

        Private Class FakeRuleImpl(Of T) : Implements IFixedRule(Of T)

            Private rule As Action(Of IFixedRuleLocator, T)

            Public Sub SetConfigure(ByVal rule As Action(Of IFixedRuleLocator, T))
                Me.rule = rule
            End Sub

            Public Sub Configure(ByVal defineBy As IFixedRuleLocator, ByVal vo As T) Implements IFixedRule(Of T).Configure
                rule.Invoke(defineBy, vo)
            End Sub
        End Class

        <Test()> Public Sub インスタンス生成直後に_固定長列設定が済んでいる_一階層()
            Dim builder As New FixedRuleBuilder(Of GrandChild)(New GrandChildRuleImpl)
            Dim root As FixedGroup = builder.MakeGroup(Nothing, 1)

            Assert.IsTrue(root.ContainsName("WEIGHT"))
            With root.GetChlid("WEIGHT")
                Assert.AreEqual(10, .Length)
                Assert.AreEqual(0, .Offset)
            End With
        End Sub

        <Test()> Public Sub インスタンス生成直後に_固定長列設定が済んでいる_二階層()
            Dim builder As New FixedRuleBuilder(Of Child)(New ChildRuleImpl)
            Dim root As FixedGroup = builder.MakeGroup(Nothing, 1)

            Assert.IsTrue(root.ContainsName("name"))
            Assert.IsTrue(root.ContainsName("children"))
            Assert.IsTrue(root.ContainsName("second"))
            With root.GetChlid("name")
                Assert.AreEqual(8, .Length)
                Assert.AreEqual(0, .Offset)
            End With
            With root.GetChlid("children")
                Assert.AreEqual(10, .Length)
                Assert.AreEqual(8, .Offset)
            End With
            With root.GetChlid("second")
                Assert.AreEqual(4, .Length)
                Assert.AreEqual(28, .Offset)
            End With
        End Sub

        <Test()> Public Sub インスタンス生成直後に_固定長列設定が済んでいる_三階層()
            Dim builder As New FixedRuleBuilder(Of Parent)(New ParentRuleImpl)
            Dim root As FixedGroup = builder.MakeGroup(Nothing, 1)

            Assert.IsTrue(root.ContainsName("num"))
            Assert.IsTrue(root.ContainsName("children"))
            Assert.IsTrue(root.ContainsName("cost"))
            With root.GetChlid("num")
                Assert.AreEqual(3, .Length)
                Assert.AreEqual(0, .Offset)
            End With
            With root.GetChlid("children")
                Assert.AreEqual(32, .Length)
                Assert.AreEqual(3, .Offset)
            End With
            With root.GetChlid("cost")
                Assert.AreEqual(8, .Length)
                Assert.AreEqual(99, .Offset)
            End With
        End Sub

        <Test()> Public Sub 右端まで切り取る設定をした後にさらにルールを追加すると_例外になる()
            Dim ruleImpl As New FakeRuleImpl(Of Child2)
            ruleImpl.SetConfigure(Sub(defineBy, vo)
                                      defineBy.Zenkaku(vo.First, 8).Hankaku(vo.Second).Hankaku(vo.Third, 4)
                                  End Sub)

            Try
                Dim ignore As New FixedRuleBuilder(Of Child2)(ruleImpl)
                Assert.Fail()
            Catch expected As InvalidOperationException
                Assert.That(expected.Message, [Is].StringContaining("固定長文字列の最後まで処理されています"))
            End Try
        End Sub

    End Class
End Namespace