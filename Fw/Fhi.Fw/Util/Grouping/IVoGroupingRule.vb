Namespace Util.Grouping
    Public Interface IVoGroupingRule(Of T)
        ''' <summary>
        ''' グループ化義を行う
        ''' </summary>
        ''' <param name="group"></param>
        ''' <param name="vo"></param>
        ''' <remarks></remarks>
        Sub Configure(ByVal group As IVoGroupingLocator, ByVal vo As T)
    End Interface
End Namespace