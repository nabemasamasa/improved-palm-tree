Imports NUnit.Framework

Namespace Db
    Public MustInherit Class CriteriaTest

#Region "Nested Classes"
        Private Class HogeVo
            Public Property Name() As String
            Public Property SubName() As String
        End Class
#End Region

        Private vo As HogeVo
        Private sut As CriteriaBinder

        <SetUp()> Public Sub SetUp()
            vo = New HogeVo()
            sut = New Criteria(Of HogeVo)(vo)
        End Sub

        Public Class GetValueByIdentifyNameTest : Inherits CriteriaTest

            <Test()> Public Sub プロパティ名と設定順のナンバーを繋げた識別名で_設定した値を取得できる()
                sut.Equal(vo.Name, "ねえむ").Not.Equal(vo.SubName, "さぶ")

                Assert.That(sut.GetValueByIdentifyName("Name0"), [Is].EqualTo("ねえむ"))
                Assert.That(sut.GetValueByIdentifyName("SubName1"), [Is].EqualTo("さぶ"))
            End Sub

            <Test()> Public Sub 同じプロパティに対し条件設定しても識別名が違うのでそれぞれ取得できる()
                sut.GreaterThan(vo.Name, "ぜろ").LessThan(vo.Name, "わん").GreaterEqual(vo.SubName, "つー").LessEqual(vo.SubName, "すりー")

                Assert.That(sut.GetValueByIdentifyName("Name0"), [Is].EqualTo("ぜろ"))
                Assert.That(sut.GetValueByIdentifyName("Name1"), [Is].EqualTo("わん"))
                Assert.That(sut.GetValueByIdentifyName("SubName2"), [Is].EqualTo("つー"))
                Assert.That(sut.GetValueByIdentifyName("SubName3"), [Is].EqualTo("すりー"))
            End Sub

            <Test()> Public Sub Anyで設定した配列型の値は_括弧インデックスで指定すれば個別に取得できる()
                sut.Any(vo.Name, {"ぜろぜろ", "ぜろわん", "ぜろつー"})

                Assert.That(sut.GetValueByIdentifyName("Name0(0)"), [Is].EqualTo("ぜろぜろ"))
                Assert.That(sut.GetValueByIdentifyName("Name0(1)"), [Is].EqualTo("ぜろわん"))
                Assert.That(sut.GetValueByIdentifyName("Name0(2)"), [Is].EqualTo("ぜろつー"))
            End Sub

            <Test()> Public Sub Anyで設定したList型の値は_括弧インデックスで指定すれば個別に取得できる()
                sut.Any(vo.Name, New List(Of String)({"いち", "に", "さん"}))

                Assert.That(sut.GetValueByIdentifyName("Name0(0)"), [Is].EqualTo("いち"))
                Assert.That(sut.GetValueByIdentifyName("Name0(1)"), [Is].EqualTo("に"))
                Assert.That(sut.GetValueByIdentifyName("Name0(2)"), [Is].EqualTo("さん"))
            End Sub

            <Test()> Public Sub Anyで設定したIEnumerable型の値は_括弧インデックスで指定すれば個別に取得できる()
                Dim values As IEnumerable(Of String) = {"うの", "どす", "とぅれす"}.Where(Function(value) True)
                sut.Any(vo.Name, values)

                Assert.That(sut.GetValueByIdentifyName("Name0(0)"), [Is].EqualTo("うの"))
                Assert.That(sut.GetValueByIdentifyName("Name0(1)"), [Is].EqualTo("どす"))
                Assert.That(sut.GetValueByIdentifyName("Name0(2)"), [Is].EqualTo("とぅれす"))
            End Sub

        End Class

    End Class
End Namespace