Namespace Util.Fixed.Extend
    ''' <summary>
    ''' 固定長の列定義設定を担うインターフェース
    ''' </summary>
    ''' <typeparam name="T">固定長列設定を行うVo型</typeparam>
    ''' <remarks></remarks>
    Public Interface IFixedRule(Of T)

        ''' <summary>
        ''' 固定長の列定義を行う
        ''' </summary>
        ''' <param name="defineBy"></param>
        ''' <param name="vo"></param>
        ''' <remarks></remarks>
        Sub Configure(ByVal defineBy As IFixedRuleLocator, ByVal vo As T)

    End Interface
End Namespace