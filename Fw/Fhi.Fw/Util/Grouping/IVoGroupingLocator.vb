Namespace Util.Grouping
    ''' <summary>
    ''' 項目のLocatorインターフェース
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IVoGroupingLocator

        ' Minが欲しい →必要なときに作ろう

        ''' <summary>
        ''' グループ化する項目を指定する
        ''' </summary>
        ''' <param name="groupingFields">グループ化する項目[]</param>
        ''' <returns>項目のLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function By(ByVal ParamArray groupingFields As Object()) As IVoGroupingLocator

        ''' <summary>
        ''' 最大値を取得する項目を指定する
        ''' </summary>
        ''' <param name="maxValueField">最大値を取得する項目</param>
        ''' <returns>項目のLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function Max(ByVal maxValueField As Object) As IVoGroupingLocator

        ''' <summary>
        ''' 先頭値を取得する項目を指定する
        ''' </summary>
        ''' <param name="maxValueField">先頭値を取得する項目</param>
        ''' <returns>項目のLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function Top(ByVal maxValueField As Object) As IVoGroupingLocator
    End Interface
End Namespace