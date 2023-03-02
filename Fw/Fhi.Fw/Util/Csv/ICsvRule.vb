Namespace Util.Csv
    Public Interface ICsvRule(Of T)

        ''' <summary>
        ''' CSVの列順定義を行う
        ''' </summary>
        ''' <param name="defineBy"></param>
        ''' <param name="vo"></param>
        ''' <remarks></remarks>
        Sub Configure(ByVal defineBy As ICsvRuleLocator, ByVal vo As T)
    End Interface
End Namespace