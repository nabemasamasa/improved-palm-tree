Imports System.Collections.Generic
Imports Fhi.Fw.Util.LabelValue
Imports NUnit.Framework

Namespace Util.LabelValue

    <TestFixture()> Public Class LabelValueComparerTest

        <Test()> Public Sub 通常はラベルでソート()
            Dim vos As New List(Of LabelValueVo)
            Dim vo1 As New LabelValueVo
            vo1.Label = "Z"
            vo1.Value = "A"
            Dim vo2 As New LabelValueVo
            vo2.Label = "Y"
            vo2.Value = "B"
            vos.Add(vo1)
            vos.Add(vo2)
            vos.Sort(New LabelValueComparer)

            Assert.AreEqual("Y", vos(0).Label)
            Assert.AreEqual("Z", vos(1).Label)
        End Sub

        <Test()> Public Sub 引数false指定で値でソート()
            Dim vos As New List(Of LabelValueVo)
            Dim vo1 As New LabelValueVo
            vo1.Label = "Z"
            vo1.Value = "A"
            Dim vo2 As New LabelValueVo
            vo2.Label = "Y"
            vo2.Value = "B"
            vos.Add(vo1)
            vos.Add(vo2)
            vos.Sort(New LabelValueComparer(False))

            Assert.AreEqual("Z", vos(0).Label)
            Assert.AreEqual("Y", vos(1).Label)
        End Sub

        <Test()> Public Sub nothing値もソート()
            Dim vos As New List(Of LabelValueVo)
            Dim vo1 As New LabelValueVo
            vo1.Label = Nothing
            vo1.Value = "A"
            Dim vo2 As New LabelValueVo
            vo2.Label = "Y"
            vo2.Value = Nothing
            Dim vo3 As New LabelValueVo
            vo3.Label = Nothing
            vo3.Value = "1"
            vos.Add(vo1)
            vos.Add(vo2)
            vos.Add(vo3)
            vos.Sort(New LabelValueComparer)

            Assert.AreEqual("Y", vos(0).Label)

            Assert.IsNull(vos(1).Label)
            Assert.AreEqual("1", vos(1).Value)

            Assert.IsNull(vos(2).Label)
            Assert.AreEqual("A", vos(2).Value)
        End Sub
    End Class
End Namespace
