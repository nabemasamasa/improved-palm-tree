Namespace Util.Csv
    ''' <summary>
    ''' 固定長の列定義設定を担うクラス
    ''' </summary>
    ''' <typeparam name="T">固定長列設定を行うVo型</typeparam>
    Public Class CsvRule(Of T) : Implements ICsvRule(Of T)

        Delegate Function RuleConfigure(ByVal defineBy As ICsvRuleLocator, ByVal vo As T) As ICsvRuleLocator

        Private ReadOnly rule As RuleConfigure

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">区切り列順設定ルール（ラムダ式）</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As RuleConfigure)
            Me.rule = rule
        End Sub

        ''' <summary>
        ''' CSVの列順定義を行う
        ''' </summary>
        ''' <param name="defineBy"></param>
        ''' <param name="vo"></param>
        ''' <remarks></remarks>
        Public Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As T) Implements ICsvRule(Of T).Configure
            rule(defineBy, vo)
        End Sub
    End Class
End Namespace