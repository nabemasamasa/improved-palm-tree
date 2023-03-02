Imports Fhi.Fw.Util.Fixed
Imports NUnit.Framework

Namespace Util.Fixed
    Public Class FixedGroupHelperTest

        <Test()> Public Sub GetEntry_存在しない名前なら例外発生()
            Dim helper As New FixedGroupHelper(New FixedGroup(New FixedField("DUMMY", 10, False)))
            Try
                helper.GetEntry("NAME")
                Assert.Fail()
            Catch expected As ArgumentException
                Assert.IsTrue(True)
            End Try
        End Sub

        <Test()> Public Sub GetOffset_通常()
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
            Dim helper As New FixedGroupHelper(root)

            Assert.AreEqual(0, helper.GetOffset("ID"))
            Assert.AreEqual(12, helper.GetOffset("HOGE"))
            Assert.AreEqual(12, helper.GetOffset("HOGE.NUM"))
            Assert.AreEqual(77, helper.GetOffset("HOGE[1].NUM"))
            Assert.AreEqual(15, helper.GetOffset("HOGE[0].FUGA"))
            Assert.AreEqual(33, helper.GetOffset("HOGE[0].FUGA[1].FIRST"))
        End Sub

    End Class
End Namespace