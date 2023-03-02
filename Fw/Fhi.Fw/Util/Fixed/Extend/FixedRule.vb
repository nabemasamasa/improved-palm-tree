Namespace Util.Fixed.Extend
    ''' <summary>
    ''' 固定長の列定義設定を担うクラス
    ''' </summary>
    ''' <typeparam name="T">固定長列設定を行うVo型</typeparam>
    Public Class FixedRule(Of T) : Implements IFixedRule(Of T)

        Delegate Function RuleConfigure(ByVal defineBy As IFixedRuleLocator, ByVal vo As T) As IFixedRuleLocator

        Private ReadOnly rule As RuleConfigure

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">固定長列設定ルール（ラムダ式）</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As RuleConfigure)
            Me.rule = rule
        End Sub

        ''' <summary>
        ''' 固定長の列定義を行う
        ''' </summary>
        ''' <param name="defineBy"></param>
        ''' <param name="vo"></param>
        ''' <remarks></remarks>
        Public Sub Configure(ByVal defineBy As IFixedRuleLocator, ByVal vo As T) Implements IFixedRule(Of T).Configure
            rule(defineBy, vo)
        End Sub
    End Class
End Namespace