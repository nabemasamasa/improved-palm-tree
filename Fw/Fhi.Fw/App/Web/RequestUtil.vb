Imports System.Web

Namespace App.Web

    ''' <summary>
    ''' Requestに対する単純な関数をまとめたクラス
    ''' </summary>
    Public Class RequestUtil

        ''' <summary>IE（10以前）のユーザーエージェント</summary>
        Private Const UA_EARLY_IE As String = "MSIE"
        ''' <summary>IE（8以降）のユーザーエージェント</summary>
        Private Const UA_LATE_IE As String = "Trident/"
        ''' <summary>旧Edgeのユーザーエージェント</summary>
        Private Const UA_OLD_EDGE As String = "Edge/"
        ''' <summary>Chromium版Edgeのユーザーエージェント</summary>
        Private Const UA_CHROMIUM_EDGE As String = "Edg/"

        ''' <summary>
        ''' Edgeからのアクセスか
        ''' </summary>
        ''' <param name="request"></param>
        ''' <returns>EdgeからのアクセスならTrue</returns>
        Public Shared Function UsesEdge(request As HttpRequestBase) As Boolean
            Return UsesBy({UA_OLD_EDGE, UA_CHROMIUM_EDGE}, request)
        End Function

        ''' <summary>
        ''' 旧Edgeからのアクセスか
        ''' </summary>
        ''' <param name="request"></param>
        ''' <returns>旧EdgeからのアクセスならTrue</returns>
        Public Shared Function UsesOldEdge(request As HttpRequestBase) As Boolean
            Return UsesBy({UA_OLD_EDGE}, request)
        End Function

        ''' <summary>
        ''' Chromium版Edgeからのアクセスか
        ''' </summary>
        ''' <param name="request"></param>
        ''' <returns>Chromium版EdgeからのアクセスならTrue</returns>
        Public Shared Function UsesChromiumEdge(request As HttpRequestBase) As Boolean
            Return UsesBy({UA_CHROMIUM_EDGE}, request)
        End Function

        ''' <summary>
        ''' IEからのアクセスか
        ''' </summary>
        ''' <param name="request"></param>
        ''' <returns>IEからのアクセスならTrue</returns>
        Public Shared Function UsesIE(request As HttpRequestBase) As Boolean
            Return UsesBy({UA_EARLY_IE, UA_LATE_IE}, request)
        End Function

        Private Shared Function UsesBy(identifiers As String(), request As HttpRequestBase) As Boolean
            Return StringUtil.IsNotEmpty(request.UserAgent) AndAlso identifiers.Any(Function(identifier) request.UserAgent.Contains(identifier))
        End Function

    End Class
End Namespace