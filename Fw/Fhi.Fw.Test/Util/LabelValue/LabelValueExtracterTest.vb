Imports System
Imports System.Collections.Generic
Imports Fhi.Fw.Util.LabelValue
Imports NUnit.Framework

Namespace Util.LabelValue

    <TestFixture()> Public MustInherit Class LabelValueExtracterTest

        Public Class 従来の方式 : Inherits LabelValueExtracterTest

            <CLSCompliant(False)> Public Shared Sub ExtractionImpl(ByVal aLocator As ILabelValueBaseLocator)
                Dim vo As New FakeSampleVo
                aLocator.IsA(vo).Label(vo.HogeName).Value(vo.HogeId)
            End Sub

            <Test()> Public Sub Extract_通常()
                Dim args As New List(Of FakeSampleVo)
                args.Add(New FakeSampleVo(123, "hogeName", Nothing, CDec("123.456789")))
                args.Add(New FakeSampleVo(12, "nameFuga", Nothing, CDec("12.3456789")))

                Dim testee As New LabelValueExtracter(Of FakeSampleVo)(args)
                Dim actuals As List(Of LabelValueVo) = testee.Extract(AddressOf ExtractionImpl)

                Assert.IsNotNull(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("123", actuals(0).Value)
                Assert.AreEqual("nameFuga", actuals(1).Label)
            End Sub

            <CLSCompliant(False)> Public Shared Sub ExtractionImpl2(ByVal aLocator As ILabelValueBaseLocator)
                Dim vo As New FakeSampleVo
                aLocator.IsA(vo).Label(vo.HogeDecimal).Value(vo.HogeId)
            End Sub

            <Test()> Public Sub Extract_通常2()
                Dim args As New List(Of FakeSampleVo)
                args.Add(New FakeSampleVo(123, "hogeName", Nothing, CDec("123.456789")))
                args.Add(New FakeSampleVo(12, "nameFuga", Nothing, CDec("12.3456789")))

                Dim testee As New LabelValueExtracter(Of FakeSampleVo)(args)
                Dim actuals As List(Of LabelValueVo) = testee.Extract(AddressOf ExtractionImpl2)

                Assert.IsNotNull(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("123.456789", actuals(0).Label)
                Assert.AreEqual("12.3456789", actuals(1).Label)
            End Sub

            Public Class ExtractRuleImpl3 : Implements ILabelValueRule
                <CLSCompliant(False)> _
                Public Sub Extraction(ByVal aLocator As ILabelValueBaseLocator) Implements ILabelValueRule.Extraction
                    Dim vo As New FakeSampleVo
                    aLocator.IsA(vo).Label(vo.HogeName).Value(vo.HogeId)
                End Sub
            End Class

            <Test()> Public Sub Extract_Shared_Interface_通常()
                Dim args As New List(Of FakeSampleVo)
                args.Add(New FakeSampleVo(123, "hogeName", Nothing, CDec("123.456789")))
                args.Add(New FakeSampleVo(12, "nameFuga", Nothing, CDec("12.3456789")))

                Dim actuals As List(Of LabelValueVo) = _
                    LabelValueExtracter(Of FakeSampleVo).Extract(args, New ExtractRuleImpl3)

                Assert.IsNotNull(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("123", actuals(0).Value)
                Assert.AreEqual("nameFuga", actuals(1).Label)
            End Sub

            <Test()> Public Sub Extract_Shared_Interface_重複データは除外する()
                Dim args As New List(Of FakeSampleVo)
                args.Add(New FakeSampleVo(123, "A", Nothing, CDec("123.456789")))
                args.Add(New FakeSampleVo(12, "B", Nothing, CDec("12.3456789")))
                args.Add(New FakeSampleVo(123, "A", Nothing, CDec("3.14159")))
                args.Add(New FakeSampleVo(123, "B", Nothing, CDec("12.3456789")))

                Dim actuals As List(Of LabelValueVo) = _
                    LabelValueExtracter(Of FakeSampleVo).Extract(args, New ExtractRuleImpl3)

                Assert.IsNotNull(actuals)
                Assert.AreEqual(3, actuals.Count)
                Assert.AreEqual("123", actuals(0).Value)
                Assert.AreEqual("12", actuals(1).Value)
                Assert.AreEqual("123", actuals(2).Value)
            End Sub

        End Class

        Public Class ラムダ式方式 : Inherits LabelValueExtracterTest

            <Test()> Public Sub Extract_通常()
                Dim args As New List(Of FakeSampleVo)
                args.Add(New FakeSampleVo(123, "hogeName", Nothing, CDec("123.456789")))
                args.Add(New FakeSampleVo(12, "nameFuga", Nothing, CDec("12.3456789")))

                Dim actuals As List(Of LabelValueVo) = _
                    LabelValueExtracter(Of FakeSampleVo).Extract(args, Function(locator As ILabelValueLocator, vo As FakeSampleVo) _
                                                                           locator.Label(vo.HogeName).Value(vo.HogeId))
                Assert.IsNotNull(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("123", actuals(0).Value)
                Assert.AreEqual("nameFuga", actuals(1).Label)
            End Sub

            <Test()> Public Sub Extract_通常2()
                Dim args As New List(Of FakeSampleVo)
                args.Add(New FakeSampleVo(123, "hogeName", Nothing, CDec("123.456789")))
                args.Add(New FakeSampleVo(12, "nameFuga", Nothing, CDec("12.3456789")))

                Dim actuals As List(Of LabelValueVo) = _
                    LabelValueExtracter(Of FakeSampleVo).Extract(args, Function(locator As ILabelValueLocator, vo As FakeSampleVo) _
                                                                           locator.Label(vo.HogeDecimal).Value(vo.HogeId))
                Assert.IsNotNull(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("123.456789", actuals(0).Label)
                Assert.AreEqual("12.3456789", actuals(1).Label)
            End Sub

            <Test()> Public Sub Extract_Shared_Interface_通常()
                Dim args As New List(Of FakeSampleVo)
                args.Add(New FakeSampleVo(123, "hogeName", Nothing, CDec("123.456789")))
                args.Add(New FakeSampleVo(12, "nameFuga", Nothing, CDec("12.3456789")))

                Dim actuals As List(Of LabelValueVo) = _
                    LabelValueExtracter(Of FakeSampleVo).Extract(args, Function(locator As ILabelValueLocator, vo As FakeSampleVo) _
                                                                           locator.Label(vo.HogeName).Value(vo.HogeId))
                Assert.IsNotNull(actuals)
                Assert.AreEqual(2, actuals.Count)
                Assert.AreEqual("123", actuals(0).Value)
                Assert.AreEqual("nameFuga", actuals(1).Label)
            End Sub

            <Test()> Public Sub Extract_Shared_Interface_重複データは除外する()
                Dim args As New List(Of FakeSampleVo)
                args.Add(New FakeSampleVo(123, "A", Nothing, CDec("123.456789")))
                args.Add(New FakeSampleVo(12, "B", Nothing, CDec("12.3456789")))
                args.Add(New FakeSampleVo(123, "A", Nothing, CDec("3.14159")))
                args.Add(New FakeSampleVo(123, "B", Nothing, CDec("12.3456789")))

                Dim actuals As List(Of LabelValueVo) = _
                    LabelValueExtracter(Of FakeSampleVo).Extract(args, Function(locator As ILabelValueLocator, vo As FakeSampleVo) _
                                                                           locator.Label(vo.HogeName).Value(vo.HogeId))
                Assert.IsNotNull(actuals)
                Assert.AreEqual(3, actuals.Count)
                Assert.AreEqual("123", actuals(0).Value)
                Assert.AreEqual("12", actuals(1).Value)
                Assert.AreEqual("123", actuals(2).Value)
            End Sub

        End Class

    End Class
End Namespace
