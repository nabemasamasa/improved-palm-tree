Imports NUnit.Framework

''' <summary>
''' 浮動小数点の挙動を学習
''' </summary>
''' <remarks></remarks>
Public MustInherit Class LearningFloatTest

    <SetUp()> Public Overridable Sub SetUp()

    End Sub

    Public Class 浮動小数点で持つ小数点以下は誤差がつきものTest : Inherits LearningFloatTest

        <Test()> Public Sub Double型だと_01plus02equal03にならない()
            ' 0.1 + 0.2 = 0.3 にならない
            Dim answer As Double = 0.3R
            Dim calculatedValue As Double = 0.1R + 0.2R
            Assert.That(answer = calculatedValue, [Is].False, "0.1 + 0.2 = 0.3 にならない")
            Assert.That(answer, [Is].Not.EqualTo(calculatedValue), "0.1 + 0.2 = 0.3 にならない")

            Assert.That(answer.ToString, [Is].EqualTo("0.3"), "ToStringだと同じにみえる")
            Assert.That(calculatedValue.ToString, [Is].EqualTo("0.3"), "ToStringだと同じにみえる")
            Assert.That(answer.ToString("G17"), [Is].EqualTo("0.29999999999999999"))
            Assert.That(calculatedValue.ToString("G17"), [Is].EqualTo("0.30000000000000004"))
        End Sub

        <Test()> Public Sub Single型だと_01plus02equal03になる()
            ' 0.1 + 0.2 = 0.3 にならない
            Dim answer As Single = 0.3F
            Dim calculatedValue As Single = 0.1F + 0.2F
            Assert.That(answer = calculatedValue, [Is].True, "0.1 + 0.2 = 0.3 にならない")
            Assert.That(answer, [Is].EqualTo(calculatedValue), "0.1 + 0.2 = 0.3 になる")
            Assert.That(answer.ToString("G17"), [Is].EqualTo("0.300000012"))
        End Sub

        <Test()> Public Sub Single型だと_01plus02equal03になる_Double値からSingle型()
            ' 0.1 + 0.2 = 0.3 にならない
            Dim answer As Single = 0.3R
            Dim calculatedValue As Single = 0.1R + 0.2R
            Assert.That(answer = calculatedValue, [Is].True, "0.1 + 0.2 = 0.3 になる")
            Assert.That(answer, [Is].EqualTo(calculatedValue), "0.1 + 0.2 = 0.3 になる")
            Assert.That(answer.ToString("G17"), [Is].EqualTo("0.300000012"))
        End Sub

    End Class

End Class