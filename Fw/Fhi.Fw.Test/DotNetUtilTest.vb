Imports NUnit.Framework

Public MustInherit Class DotNetUtilTest

    Public Class HasInstalledDotNet35 : Inherits DotNetUtilTest

        <Test()> Public Sub 当プロジェクトはDotNet35だからtrue()
            Assert.IsTrue(DotNetUtil.HasInstalledDotNet35, "Fhi.Fwプロジェクトは.NET3.5だから、true")
        End Sub

    End Class

    Public Class ConvCompiledDateTest : Inherits DotNetUtilTest

        <Test()> Public Sub _1_1は1日後で2秒後()
            Assert.That(DotNetUtil.ConvCompiledDate(New Version(1, 1, 1, 1)), [Is].EqualTo(New Date(2000, 1, 2, 0, 0, 2)))
        End Sub

        <Test()> Public Sub _10_100は10日後で200秒後()
            Assert.That(DotNetUtil.ConvCompiledDate(New Version(1, 1, 10, 100)), [Is].EqualTo(New Date(2000, 1, 11, 0, 3, 20)))
        End Sub

    End Class

End Class
