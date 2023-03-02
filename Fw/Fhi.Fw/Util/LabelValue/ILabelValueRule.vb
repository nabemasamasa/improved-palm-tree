Namespace Util.LabelValue
    Public Interface ILabelValueRule
        ''' <summary>
        ''' Label化/Value化定義を行う
        ''' </summary>
        ''' <param name="aLocator">VOの Label と Value を管理するインターフェース</param>
        ''' <remarks></remarks>
        Sub Extraction(ByVal aLocator As ILabelValueBaseLocator)
    End Interface

    Public Interface ILabelValueRule(Of T)
        ''' <summary>
        ''' Label化/Value化定義を行う
        ''' </summary>
        ''' <param name="group"></param>
        ''' <param name="vo"></param>
        ''' <remarks></remarks>
        Sub Configure(ByVal group As ILabelValueLocator, ByVal vo As T)
    End Interface
End Namespace