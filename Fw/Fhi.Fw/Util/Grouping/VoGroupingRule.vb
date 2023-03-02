Namespace Util.Grouping
    ''' <summary>
    ''' グループ化設定を担うクラス
    ''' </summary>
    ''' <typeparam name="T">グループ化設定を行うVo型</typeparam>
    Public Class VoGroupingRule(Of T) : Implements IVoGroupingRule(Of T)

        Delegate Function RuleConfigure(ByVal rule As IVoGroupingLocator, ByVal vo As T) As IVoGroupingLocator

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
        ''' グループ化義を行う
        ''' </summary>
        ''' <param name="group"></param>
        ''' <param name="vo"></param>
        ''' <remarks></remarks>
        Public Sub Configure(ByVal group As IVoGroupingLocator, ByVal vo As T) Implements IVoGroupingRule(Of T).Configure
            Me.rule(group, vo)
        End Sub
    End Class
End Namespace