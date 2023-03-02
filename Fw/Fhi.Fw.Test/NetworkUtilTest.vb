Imports NUnit.Framework

Public Class NetworkUtilTest

    <Test()> Public Sub IsIPv4_String()
        Assert.AreEqual(True, NetworkUtil.IsIPv4("192.168.0.0"))
        Assert.AreEqual(False, NetworkUtil.IsIPv4("192.168.0.-1"))
        Assert.AreEqual(False, NetworkUtil.IsIPv4("abcd::1234"))
    End Sub

End Class
