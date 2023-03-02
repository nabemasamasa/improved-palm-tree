Namespace Sys
    ''' <summary>
    ''' 当フレームワーク（Fhi.Fw）の初期設定を担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FwSetup
        ''' <summary>
        ''' Fhi.Fwの初期設定を行う
        ''' </summary>
        ''' <param name="systemDateBehavior">システム日付取得の振る舞い</param>
        ''' <remarks></remarks>
        Public Shared Sub Initialize(ByVal systemDateBehavior As SystemDate.IBehavior)
            SystemDate.DefaultBehavior = systemDateBehavior
        End Sub

        ''' <summary>
        ''' Fhi.Fwの初期設定をクリアする
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Clear()
            SystemDate.DefaultBehavior = Nothing
        End Sub
    End Class
End Namespace