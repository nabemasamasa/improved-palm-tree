Imports NUnit.Framework

Namespace App
    Public MustInherit Class InternetExplorerManipulatorTest

        <SetUp()> Public Overridable Sub SetUp()

        End Sub

        Public Class RemoveTagsTest : Inherits InternetExplorerManipulatorTest

            <Test(), Sequential()> Public Sub 内包されたタグを除去する( _
                    <Values("<A>HOGE", "FUGA</A>", "<A>PIYO</A>", "<IMG src='...'>ABC", "CONT<INPUT type='hidden'>AINS")> ByVal html As String, _
                    <Values("HOGE", "FUGA", "PIYO", "ABC", "CONTAINS")> ByVal expected As String)
                Assert.That(InternetExplorerManipulator.RemoveTags(html), [Is].EqualTo(expected))
            End Sub

            <Test(), Sequential()> Public Sub nullならnullを返すだけ()
                Assert.That(InternetExplorerManipulator.RemoveTags(Nothing), [Is].Null)
            End Sub

        End Class

    End Class
End Namespace
