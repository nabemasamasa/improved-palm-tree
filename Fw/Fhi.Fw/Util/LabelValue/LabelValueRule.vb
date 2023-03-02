Namespace Util.LabelValue
    ''' <summary>
    ''' Label化/Value化設定を担うクラス
    ''' </summary>
    ''' <typeparam name="T">グループ化設定を行うVo型</typeparam>
    Public Class LabelValueRule(Of T) : Implements ILabelValueRule(Of T)

        Delegate Function RuleConfigure(ByVal locator As ILabelValueLocator, ByVal vo As T) As ILabelValueLocator

        Private ReadOnly rule As RuleConfigure

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">グループ化設定ルール（ラムダ式）</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As RuleConfigure)
            Me.rule = rule
        End Sub

        ''' <summary>
        ''' Label化/Value化定義を行う
        ''' </summary>
        ''' <param name="locator"></param>
        ''' <param name="vo"></param>
        ''' <remarks></remarks>
        Public Sub Configure(ByVal locator As ILabelValueLocator, ByVal vo As T) Implements ILabelValueRule(Of T).Configure
            Me.rule(locator, vo)
        End Sub

    End Class
End Namespace