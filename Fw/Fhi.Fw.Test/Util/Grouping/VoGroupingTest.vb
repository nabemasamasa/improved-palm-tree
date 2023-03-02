Imports System.Collections.Generic
Imports Fhi.Fw.Util.Grouping
Imports NUnit.Framework

Namespace Util.Grouping
    Public MustInherit Class VoGroupingTest

        Private Function NewVo(ByVal eventCode As String, ByVal eventName As String, ByVal hyojijunNo As Integer?) As FakeEventVo
            Dim vo As New FakeEventVo
            vo.ShisakuEventCode = eventCode
            vo.ShisakuEventName = eventName
            vo.HyojijunNo = hyojijunNo
            Return vo
        End Function

        Private Class RuleMaxHyojijunGroupByEventCode : Implements IVoGroupingRule(Of FakeEventVo)
            Public Sub Configure(ByVal group As IVoGroupingLocator, ByVal vo As FakeEventVo) Implements IVoGroupingRule(Of FakeEventVo).Configure
                group.By(vo.ShisakuEventCode).Max(vo.HyojijunNo)
            End Sub
        End Class

        Private Class RuleMaxEventNameGroupByEventCode : Implements IVoGroupingRule(Of FakeEventVo)
            Public Sub Configure(ByVal group As IVoGroupingLocator, ByVal vo As FakeEventVo) Implements IVoGroupingRule(Of FakeEventVo).Configure
                group.By(vo.ShisakuEventCode).Max(vo.ShisakuEventName)
            End Sub
        End Class

        Private Class RuleTopHyojijunGroupByEventCode : Implements IVoGroupingRule(Of FakeEventVo)
            Public Sub Configure(ByVal group As IVoGroupingLocator, ByVal vo As FakeEventVo) Implements IVoGroupingRule(Of FakeEventVo).Configure
                group.By(vo.ShisakuEventCode).Top(vo.HyojijunNo)
            End Sub
        End Class

        Private Class Grouping : Implements IVoGroupingRule(Of FakeEventVo)
            Public Sub Configure(ByVal group As IVoGroupingLocator, ByVal vo As FakeEventVo) Implements IVoGroupingRule(Of FakeEventVo).Configure
                group.By(vo.ShisakuEventName)
            End Sub
        End Class

        Public Class 元々のテストたち : Inherits VoGroupingTest

            <Test()> Public Sub 一項目でGrouping()
                Dim grouping As New VoGrouping(Of FakeEventVo)(New Grouping)

                Dim vo1 As FakeEventVo = NewVo("123", "hoge", 1)
                Dim vo2 As FakeEventVo = NewVo("234", "hoge", 2)
                Dim vo3 As FakeEventVo = NewVo("345", "fuga", 2)

                Dim results As List(Of FakeEventVo) = grouping.Group(EzUtil.NewList(vo1, vo2, vo3))

                Assert.AreEqual(2, results.Count)
                Assert.AreEqual("hoge", results(0).ShisakuEventName)
                Assert.AreEqual("fuga", results(1).ShisakuEventName)
                Assert.IsNull(results(0).ShisakuEventCode)
                Assert.IsNull(results(1).ShisakuEventCode)
            End Sub

            <Test()> Public Sub 一項目でGrouping_ラムダ式()
                Dim grouping As New VoGrouping(Of FakeEventVo)(Function(group As IVoGroupingLocator, vo As FakeEventVo) group.By(vo.ShisakuEventName))

                Dim vo1 As FakeEventVo = NewVo("123", "hoge", 1)
                Dim vo2 As FakeEventVo = NewVo("234", "hoge", 2)
                Dim vo3 As FakeEventVo = NewVo("345", "fuga", 2)

                Dim results As List(Of FakeEventVo) = grouping.Group(EzUtil.NewList(vo1, vo2, vo3))

                Assert.AreEqual(2, results.Count)
                Assert.AreEqual("hoge", results(0).ShisakuEventName)
                Assert.AreEqual("fuga", results(1).ShisakuEventName)
                Assert.IsNull(results(0).ShisakuEventCode)
                Assert.IsNull(results(1).ShisakuEventCode)
            End Sub

            Private Class GroupingTwo : Implements IVoGroupingRule(Of FakeEventVo)
                Public Sub Configure(ByVal group As IVoGroupingLocator, ByVal vo As FakeEventVo) Implements IVoGroupingRule(Of FakeEventVo).Configure
                    group.By(vo.ShisakuEventName).By(vo.ShisakuEventPhaseName)
                End Sub
            End Class

            <Test()> Public Sub 二項目でGrouping_一つもグループ化されない()
                Dim grouping As New VoGrouping(Of FakeEventVo)(New GroupingTwo)

                Dim vo1 As New FakeEventVo
                vo1.ShisakuEventCode = "123"
                vo1.ShisakuEventName = "hoge"
                vo1.ShisakuEventPhaseName = "a"
                Dim vo2 As New FakeEventVo
                vo2.ShisakuEventCode = "234"
                vo2.ShisakuEventName = "hoge"
                vo2.ShisakuEventPhaseName = "b"
                Dim vo3 As New FakeEventVo
                vo3.ShisakuEventCode = "345"
                vo3.ShisakuEventName = "fuga"
                vo3.ShisakuEventPhaseName = "b"
                Dim aList As New List(Of FakeEventVo)(New FakeEventVo() {vo1, vo2, vo3})

                Dim results As List(Of FakeEventVo) = grouping.Group(aList)

                Assert.AreEqual(3, results.Count)
                Assert.AreEqual("hoge", results(0).ShisakuEventName)
                Assert.AreEqual("a", results(0).ShisakuEventPhaseName)
                Assert.AreEqual("hoge", results(1).ShisakuEventName)
                Assert.AreEqual("b", results(1).ShisakuEventPhaseName)
                Assert.AreEqual("fuga", results(2).ShisakuEventName)
                Assert.AreEqual("b", results(2).ShisakuEventPhaseName)
                Assert.IsNull(results(0).ShisakuEventCode)
                Assert.IsNull(results(1).ShisakuEventCode)
                Assert.IsNull(results(2).ShisakuEventCode)
            End Sub

            <Test()> Public Sub ResultBreakdown_グループ化した内容の内訳()
                Dim grouping As New VoGrouping(Of FakeEventVo)(New GroupingTwo)

                Dim vo1 As New FakeEventVo
                vo1.ShisakuEventCode = "123"
                vo1.ShisakuEventName = "hoge"
                vo1.ShisakuEventPhaseName = "a"
                Dim vo2 As New FakeEventVo
                vo2.ShisakuEventCode = "234"
                vo2.ShisakuEventName = "hoge"
                vo2.ShisakuEventPhaseName = "a"
                Dim vo3 As New FakeEventVo
                vo3.ShisakuEventCode = "345"
                vo3.ShisakuEventName = "fuga"
                vo3.ShisakuEventPhaseName = "b"
                Dim aList As New List(Of FakeEventVo)(New FakeEventVo() {vo1, vo2, vo3})

                Dim results As List(Of FakeEventVo) = grouping.Group(aList)

                Assert.AreEqual(2, results.Count)

                Dim actual As List(Of FakeEventVo)() = grouping.ResultBreakdown

                Assert.AreEqual(2, actual.Length)

                Assert.AreEqual(2, actual(0).Count)
                Assert.AreEqual(1, actual(1).Count)

                Assert.AreEqual("123", actual(0)(0).ShisakuEventCode)
                Assert.AreEqual("234", actual(0)(1).ShisakuEventCode)
                Assert.AreEqual("345", actual(1)(0).ShisakuEventCode)
            End Sub

            <Test()> Public Sub Max_最大値を取得()

                Dim vo1 As FakeEventVo = NewVo("123", "a", 10)
                Dim vo2 As FakeEventVo = NewVo("123", "b", 20)
                Dim vo3 As FakeEventVo = NewVo("123", "c", 101)
                Dim vo4 As FakeEventVo = NewVo("223", "c", 1001)

                Dim grouping As New VoGrouping(Of FakeEventVo)(New RuleMaxHyojijunGroupByEventCode)
                grouping.Group(EzUtil.NewList(vo1, vo2, vo3, vo4))

                Assert.AreEqual(2, grouping.Result.Count)
                With grouping.Result(0)
                    Assert.AreEqual("123", .ShisakuEventCode)
                    Assert.IsNull(.ShisakuEventName)
                    Assert.AreEqual(101, .HyojijunNo)
                End With
                With grouping.Result(1)
                    Assert.AreEqual(1001, .HyojijunNo)
                End With
            End Sub

            <Test()> Public Sub Max_Nullが含まれていても最大値を取得()

                Dim vo1 As FakeEventVo = NewVo("123", "a", 10)
                Dim vo2 As FakeEventVo = NewVo("123", "b", Nothing)
                Dim vo3 As FakeEventVo = NewVo("123", "c", 101)

                Dim grouping As New VoGrouping(Of FakeEventVo)(New RuleMaxHyojijunGroupByEventCode)
                grouping.Group(EzUtil.NewList(vo1, vo2, vo3))

                Assert.AreEqual(1, grouping.Result.Count)
                With grouping.Result(0)
                    Assert.AreEqual("123", .ShisakuEventCode)
                    Assert.IsNull(.ShisakuEventName)
                    Assert.AreEqual(101, .HyojijunNo)
                End With
            End Sub

            <Test()> Public Sub Max_Null値しかなければNull()

                Dim vo1 As FakeEventVo = NewVo("123", "a", Nothing)
                Dim vo2 As FakeEventVo = NewVo("223", "c", 10)

                Dim grouping As New VoGrouping(Of FakeEventVo)(New RuleMaxHyojijunGroupByEventCode)
                grouping.Group(EzUtil.NewList(vo1, vo2))

                Assert.AreEqual(2, grouping.Result.Count)
                With grouping.Result(0)
                    Assert.IsNull(.HyojijunNo)
                End With
                With grouping.Result(1)
                    Assert.AreEqual(10, .HyojijunNo)
                End With
            End Sub

            <Test()> Public Sub Max_最大値を取得_文字列()

                Dim vo1 As FakeEventVo = NewVo("123", "aaaaa", 10)
                Dim vo2 As FakeEventVo = NewVo("123", "bbb", 20)
                Dim vo3 As FakeEventVo = NewVo("123", "cc", 101)
                Dim vo4 As FakeEventVo = NewVo("223", "zz", 1001)

                Dim grouping As New VoGrouping(Of FakeEventVo)(New RuleMaxEventNameGroupByEventCode)
                grouping.Group(EzUtil.NewList(vo1, vo2, vo3, vo4))

                Assert.AreEqual(2, grouping.Result.Count)
                With grouping.Result(0)
                    Assert.AreEqual("123", .ShisakuEventCode)
                    Assert.AreEqual("cc", .ShisakuEventName)
                    Assert.IsNull(.HyojijunNo)
                End With
                With grouping.Result(1)
                    Assert.AreEqual("zz", .ShisakuEventName)
                End With
            End Sub

            <Test()> Public Sub Top_先頭値を取得()

                Dim vo1 As FakeEventVo = NewVo("123", "a", 20)
                Dim vo2 As FakeEventVo = NewVo("123", "b", 10)
                Dim vo3 As FakeEventVo = NewVo("123", "c", 101)
                Dim vo4 As FakeEventVo = NewVo("223", "c", 1001)

                Dim grouping As New VoGrouping(Of FakeEventVo)(New RuleTopHyojijunGroupByEventCode)
                grouping.Group(EzUtil.NewList(vo1, vo2, vo3, vo4))

                Assert.AreEqual(2, grouping.Result.Count)
                With grouping.Result(0)
                    Assert.AreEqual("123", .ShisakuEventCode)
                    Assert.IsNull(.ShisakuEventName)
                    Assert.AreEqual(20, .HyojijunNo)
                End With
                With grouping.Result(1)
                    Assert.AreEqual(1001, .HyojijunNo)
                End With
            End Sub

        End Class

        Public Class 違うテスト : Inherits VoGroupingTest

            <Test()> Public Sub Group化の結果は渡した値順に返される()

                Dim grouping As New VoGrouping(Of FakeEventVo)(Function(rule As IVoGroupingLocator, vo As FakeEventVo) _
                                                                   rule.By(vo.ShisakuEventName))
                grouping.Group(NewVo("", "a", 10), NewVo("", "a", 9), NewVo("", "z", 8), NewVo("", "z", 7), _
                               NewVo("", "c", 6), NewVo("", "c", 5), NewVo("", "z", 4))

                Assert.AreEqual(3, grouping.Result.Count)
                Assert.AreEqual("a", grouping.Result(0).ShisakuEventName)
                Assert.AreEqual("z", grouping.Result(1).ShisakuEventName)
                Assert.AreEqual("c", grouping.Result(2).ShisakuEventName)
            End Sub
        End Class

    End Class
End Namespace
