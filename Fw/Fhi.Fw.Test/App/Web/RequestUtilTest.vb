Imports System.Web
Imports NUnit.Framework

Namespace App.Web
    Public MustInherit Class RequestUtilTest

#Region "Fake Classes"
        Private Class FakeRequest : Inherits HttpRequestBase
            Public ResultUserAgent As String
            Public Overrides ReadOnly Property UserAgent As String
                Get
                    Return ResultUserAgent
                End Get
            End Property
        End Class
#End Region

        Private Const USER_AGENT_GOOGLE_CHROME As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36"
        Private Const USER_AGENT_GOOGLE_CHROME_1 As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.92"

        Private Const USER_AGENT_OLD_EDGE As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.18363"
        Private Const USER_AGENT_CHROMIUM_EDGE As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36 Edg/80.0.361.111"

        Private Const USER_AGENT_PURE_IE11 As String = "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko"
        Private Const USER_AGENT_IE11_WITH_COMPATIBLE_IE8 As String = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Zoom 3.6.0)"

        Private request As FakeRequest

        <SetUp>
        Public Sub Setup()
            request = New FakeRequest
        End Sub

        Public Class UsesEdgeTest : Inherits RequestUtilTest

            <Test>
            Public Sub 旧EdgeならTrue()
                request.ResultUserAgent = USER_AGENT_OLD_EDGE
                Assert.IsTrue(RequestUtil.UsesEdge(request))
            End Sub

            <Test>
            Public Sub Chromium版EdgeならTrue()
                request.ResultUserAgent = USER_AGENT_CHROMIUM_EDGE
                Assert.IsTrue(RequestUtil.UsesEdge(request))
            End Sub

            <TestCase(USER_AGENT_GOOGLE_CHROME)>
            <TestCase(USER_AGENT_GOOGLE_CHROME_1)>
            Public Sub GoogleChromeならFalse(userAgent As String)
                request.ResultUserAgent = userAgent
                Assert.IsFalse(RequestUtil.UsesEdge(request))
            End Sub

            <TestCase(USER_AGENT_PURE_IE11)>
            <TestCase(USER_AGENT_IE11_WITH_COMPATIBLE_IE8)>
            Public Sub IEならFalse(userAgent As String)
                request.ResultUserAgent = userAgent
                Assert.IsFalse(RequestUtil.UsesEdge(request))
            End Sub
        End Class

        Public Class UsesOldEdgeTest : Inherits RequestUtilTest

            <Test>
            Public Sub 旧EdgeならTrue()
                request.ResultUserAgent = USER_AGENT_OLD_EDGE
                Assert.IsTrue(RequestUtil.UsesOldEdge(request))
            End Sub

            <TestCase(USER_AGENT_CHROMIUM_EDGE)>
            <TestCase(USER_AGENT_GOOGLE_CHROME)>
            <TestCase(USER_AGENT_GOOGLE_CHROME_1)>
            <TestCase(USER_AGENT_PURE_IE11)>
            <TestCase(USER_AGENT_IE11_WITH_COMPATIBLE_IE8)>
            Public Sub 旧Edge以外のブラウザならFalse(userAgent As String)
                request.ResultUserAgent = userAgent
                Assert.IsFalse(RequestUtil.UsesOldEdge(request))
            End Sub
        End Class

        Public Class UsesChromiumEdgeTest : Inherits RequestUtilTest

            <Test>
            Public Sub Chromium版EdgeならTrue()
                request.ResultUserAgent = USER_AGENT_CHROMIUM_EDGE
                Assert.IsTrue(RequestUtil.UsesChromiumEdge(request))
            End Sub

            <TestCase(USER_AGENT_OLD_EDGE)>
            <TestCase(USER_AGENT_GOOGLE_CHROME)>
            <TestCase(USER_AGENT_GOOGLE_CHROME_1)>
            <TestCase(USER_AGENT_PURE_IE11)>
            <TestCase(USER_AGENT_IE11_WITH_COMPATIBLE_IE8)>
            Public Sub Chromium版Edge以外のブラウザならFalse(userAgent As String)
                request.ResultUserAgent = userAgent
                Assert.IsFalse(RequestUtil.UsesChromiumEdge(request))
            End Sub
        End Class

        Public Class UsesIETest : Inherits RequestUtilTest

            <TestCase(USER_AGENT_PURE_IE11)>
            <TestCase(USER_AGENT_IE11_WITH_COMPATIBLE_IE8)>
            Public Sub IEならTrue(userAgent As String)
                request.ResultUserAgent = userAgent
                Assert.IsTrue(RequestUtil.UsesIE(request))
            End Sub

            <TestCase(USER_AGENT_GOOGLE_CHROME)>
            <TestCase(USER_AGENT_GOOGLE_CHROME_1)>
            Public Sub EdgeならFalse(userAgent As String)
                request.ResultUserAgent = userAgent
                Assert.IsFalse(RequestUtil.UsesIE(request))
            End Sub

            <TestCase(USER_AGENT_GOOGLE_CHROME)>
            <TestCase(USER_AGENT_GOOGLE_CHROME_1)>
            Public Sub GoogleChromeならFalse(userAgent As String)
                request.ResultUserAgent = userAgent
                Assert.IsFalse(RequestUtil.UsesIE(request))
            End Sub
        End Class

    End Class
End Namespace