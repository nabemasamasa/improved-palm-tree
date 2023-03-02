Imports Fhi.Fw.Util.Fixed
Imports NUnit.Framework

Namespace Util.Fixed
    Public Class FixedFieldTest

        <Test()> Public Sub ContainsName_一階層()
            Dim root As New FixedGroup( _
                New FixedField("NO", 2, False), _
                New FixedField("NAME", 10, False))
            Assert.IsTrue(root.ContainsName("NO"))
            Assert.IsTrue(root.ContainsName("NAME"))
        End Sub

        <Test()> Public Sub ContainsName_三階層()
            Dim root As New FixedGroup( _
                New FixedGroup("LV2", 1, _
                    New FixedField("NO", 2, False), _
                    New FixedGroup("LV3", 1, _
                        New FixedField("ID", 2, False), _
                        New FixedField("NUM", 3, False) _
                        ), _
                    New FixedField("NAME", 10, False) _
                    ))
            Assert.IsTrue(root.ContainsName("LV2"))
            With root.GetChlid("LV2")
                Assert.IsTrue(.ContainsName("LV3"))
                With .GetChlid("LV3")
                    Assert.IsTrue(.ContainsName("ID"))
                    Assert.IsTrue(.ContainsName("NUM"))
                End With
            End With
        End Sub

        <Test()> Public Sub InitializeOffset_三階層()
            Dim root As New FixedGroup( _
                New FixedField("ID", 2, False), _
                New FixedField("NAME", 10, False), _
                New FixedGroup("HOGE", 1, _
                    New FixedField("NUM", 3, False), _
                    New FixedGroup("FUGA", 2, _
                        New FixedField("SEQ", 4, False), _
                        New FixedField("FIRST", 10, False) _
                        ), _
                    New FixedField("SECOND", 20, False) _
                    ))
            Assert.AreEqual(0, root.GetChlid("ID").Offset)
            Assert.AreEqual(2, root.GetChlid("NAME").Offset)
            Assert.AreEqual(12, root.GetChlid("HOGE").Offset)

            With root.GetChlid("HOGE")
                Assert.AreEqual(0, .GetChlid("NUM").Offset)
                Assert.AreEqual(3, .GetChlid("FUGA").Offset)
                Assert.AreEqual(31, .GetChlid("SECOND").Offset)

                With .GetChlid("FUGA")
                    Assert.AreEqual(0, .GetChlid("SEQ").Offset)
                    Assert.AreEqual(4, .GetChlid("FIRST").Offset)
                End With
            End With
        End Sub

        <Test()> Public Sub Length_三階層()
            Dim root As New FixedGroup( _
                New FixedField("ID", 2, False), _
                New FixedField("NAME", 10, False), _
                New FixedGroup("HOGE", 2, _
                    New FixedField("NUM", 3, False), _
                    New FixedGroup("FUGA", 3, _
                        New FixedField("SEQ", 4, False), _
                        New FixedField("FIRST", 10, False) _
                        ), _
                    New FixedField("SECOND", 20, False) _
                    ))
            Assert.AreEqual(2, root.GetChlid("ID").Length)
            Assert.AreEqual(10, root.GetChlid("NAME").Length)
            Assert.AreEqual(65, root.GetChlid("HOGE").Length)

            With root.GetChlid("HOGE")
                Assert.AreEqual(3, .GetChlid("NUM").Length)
                Assert.AreEqual(14, .GetChlid("FUGA").Length)
                Assert.AreEqual(20, .GetChlid("SECOND").Length)

                With .GetChlid("FUGA")
                    Assert.AreEqual(4, .GetChlid("SEQ").Length)
                    Assert.AreEqual(10, .GetChlid("FIRST").Length)
                End With
            End With
        End Sub

    End Class
End Namespace