Imports NUnit.Framework

Public Class FakeLoggerTest
    Private sut As FakeLogger

    <SetUp()> Sub SetUp()
        sut = New FakeLogger()
    End Sub

    <Test()> Sub DEBUGログのテスト()
        sut.Debug("通常メッセージ")
        sut.Debug("例外メッセージ", New Exception("俺だよ"))
        sut.DebugFormat("整形メッセ1 {0}", "1")
        sut.DebugFormat("整形メッセ2 {0} {1}", "１", "２")
        sut.DebugFormat("整形メッセ3 {0} {1} {2}", "壱", "弐", "参")
        sut.DebugFormat("整形メッセージ {0} {1} {2} {3}", "あ", "イ", "ｳ", "E")

        Assert.That(sut.GetLog()(0), [Is].EqualTo("DEBUG 通常メッセージ"))
        Assert.That(sut.GetLog()(1), [Is].EqualTo("DEBUG 例外メッセージ" & vbCrLf & _
                                                  "System.Exception: 俺だよ"))
        Assert.That(sut.GetLog()(2), [Is].EqualTo("DEBUG 整形メッセ1 1"))
        Assert.That(sut.GetLog()(3), [Is].EqualTo("DEBUG 整形メッセ2 １ ２"))
        Assert.That(sut.GetLog()(4), [Is].EqualTo("DEBUG 整形メッセ3 壱 弐 参"))
        Assert.That(sut.GetLog()(5), [Is].EqualTo("DEBUG 整形メッセージ あ イ ｳ E"))
    End Sub

    <Test()> Sub INFOログのテスト()
        sut.Info("通常メッセージ")
        sut.Info("例外メッセージ", New Exception("わしじゃよ"))
        sut.InfoFormat("整形メッセ1 {0}", "1")
        sut.InfoFormat("整形メッセ2 {0} {1}", "１", "２")
        sut.InfoFormat("整形メッセ3 {0} {1} {2}", "壱", "弐", "参")
        sut.InfoFormat("整形メッセージ {0} {1} {2} {3}", "か", "キ", "ｸ", "KE")

        Assert.That(sut.GetLog()(0), [Is].EqualTo("INFO  通常メッセージ"))
        Assert.That(sut.GetLog()(1), [Is].EqualTo("INFO  例外メッセージ" & vbCrLf & _
                                                  "System.Exception: わしじゃよ"))
        Assert.That(sut.GetLog()(2), [Is].EqualTo("INFO  整形メッセ1 1"))
        Assert.That(sut.GetLog()(3), [Is].EqualTo("INFO  整形メッセ2 １ ２"))
        Assert.That(sut.GetLog()(4), [Is].EqualTo("INFO  整形メッセ3 壱 弐 参"))
        Assert.That(sut.GetLog()(5), [Is].EqualTo("INFO  整形メッセージ か キ ｸ KE"))
    End Sub

    <Test()> Sub WARNログのテスト()
        sut.Warn("通常メッセージ")
        sut.Warn("例外メッセージ", New Exception("わたしです"))
        sut.WarnFormat("整形メッセ1 {0}", "1")
        sut.WarnFormat("整形メッセ2 {0} {1}", "１", "２")
        sut.WarnFormat("整形メッセ3 {0} {1} {2}", "壱", "弐", "参")
        sut.WarnFormat("整形メッセージ {0} {1} {2} {3}", "さ", "シ", "ｽ", "SE")

        Assert.That(sut.GetLog()(0), [Is].EqualTo("WARN  通常メッセージ"))
        Assert.That(sut.GetLog()(1), [Is].EqualTo("WARN  例外メッセージ" & vbCrLf & _
                                                  "System.Exception: わたしです"))
        Assert.That(sut.GetLog()(2), [Is].EqualTo("WARN  整形メッセ1 1"))
        Assert.That(sut.GetLog()(3), [Is].EqualTo("WARN  整形メッセ2 １ ２"))
        Assert.That(sut.GetLog()(4), [Is].EqualTo("WARN  整形メッセ3 壱 弐 参"))
        Assert.That(sut.GetLog()(5), [Is].EqualTo("WARN  整形メッセージ さ シ ｽ SE"))
    End Sub

    <Test()> Sub ERRORログのテスト()
        sut.Error("通常メッセージ")
        sut.Error("例外メッセージ", New Exception("ボクニャウ！"))
        sut.ErrorFormat("整形メッセ1 {0}", "1")
        sut.ErrorFormat("整形メッセ2 {0} {1}", "１", "２")
        sut.ErrorFormat("整形メッセ3 {0} {1} {2}", "壱", "弐", "参")
        sut.ErrorFormat("整形メッセージ {0} {1} {2} {3}", "た", "チ", "ﾂ", "TE")

        Assert.That(sut.GetLog()(0), [Is].EqualTo("ERROR 通常メッセージ"))
        Assert.That(sut.GetLog()(1), [Is].EqualTo("ERROR 例外メッセージ" & vbCrLf & _
                                                  "System.Exception: ボクニャウ！"))
        Assert.That(sut.GetLog()(2), [Is].EqualTo("ERROR 整形メッセ1 1"))
        Assert.That(sut.GetLog()(3), [Is].EqualTo("ERROR 整形メッセ2 １ ２"))
        Assert.That(sut.GetLog()(4), [Is].EqualTo("ERROR 整形メッセ3 壱 弐 参"))
        Assert.That(sut.GetLog()(5), [Is].EqualTo("ERROR 整形メッセージ た チ ﾂ TE"))
    End Sub

    <Test()> Sub FATALログのテスト()
        sut.Fatal("通常メッセージ")
        sut.Fatal("例外メッセージ", New Exception("くもじいじゃ"))
        sut.FatalFormat("整形メッセ1 {0}", "1")
        sut.FatalFormat("整形メッセ2 {0} {1}", "１", "２")
        sut.FatalFormat("整形メッセ3 {0} {1} {2}", "壱", "弐", "参")
        sut.FatalFormat("整形メッセージ {0} {1} {2} {3}", "な", "ニ", "ﾇ", "NE")

        Assert.That(sut.GetLog()(0), [Is].EqualTo("FATAL 通常メッセージ"))
        Assert.That(sut.GetLog()(1), [Is].EqualTo("FATAL 例外メッセージ" & vbCrLf & _
                                                  "System.Exception: くもじいじゃ"))
        Assert.That(sut.GetLog()(2), [Is].EqualTo("FATAL 整形メッセ1 1"))
        Assert.That(sut.GetLog()(3), [Is].EqualTo("FATAL 整形メッセ2 １ ２"))
        Assert.That(sut.GetLog()(4), [Is].EqualTo("FATAL 整形メッセ3 壱 弐 参"))
        Assert.That(sut.GetLog()(5), [Is].EqualTo("FATAL 整形メッセージ な ニ ﾇ NE"))
    End Sub
End Class
