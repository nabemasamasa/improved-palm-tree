Namespace Util.Fixed.Extend
    ''' <summary>
    ''' 固定長ルールのLocatorインターフェース
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IFixedRuleLocator

        ''' <summary>
        ''' 半角項目を設定する
        ''' </summary>
        ''' <param name="field">半角項目</param>
        ''' <returns>固定長ルールのLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function Hankaku(ByVal field As Object, Optional ByVal decorateToVo As Func(Of String, Object) = Nothing, Optional ByVal decorateToString As Func(Of Object, String) = Nothing) As IFixedRuleLocator

        ''' <summary>
        ''' 半角項目を設定する
        ''' </summary>
        ''' <param name="field">半角項目</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <returns>固定長ルールのLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function Hankaku(ByVal field As Object, ByVal length As Integer, Optional ByVal decorateToVo As Func(Of String, Object) = Nothing, Optional ByVal decorateToString As Func(Of Object, String) = Nothing) As IFixedRuleLocator

        ''' <summary>
        ''' 半角項目を設定する
        ''' </summary>
        ''' <typeparam name="E">繰り返しの型</typeparam>
        ''' <param name="field">半角項目</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <returns>固定長ルールのLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function HankakuRepeat(Of E)(ByVal field As ICollection(Of E), ByVal length As Integer, ByVal repeat As Integer) As IFixedRuleLocator

        ''' <summary>
        ''' 全角項目を設定する
        ''' </summary>
        ''' <param name="field">全角項目</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <returns>固定長ルールのLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function Zenkaku(ByVal field As Object, ByVal length As Integer) As IFixedRuleLocator

        ''' <summary>
        ''' 全角項目を設定する
        ''' </summary>
        ''' <typeparam name="E">繰り返しの型</typeparam>
        ''' <param name="field">全角項目</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <returns>固定長ルールのLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function ZenkakuRepeat(Of E)(ByVal field As ICollection(Of E), ByVal length As Integer, ByVal repeat As Integer) As IFixedRuleLocator

        ''' <summary>
        ''' 数値項目を設定する
        ''' </summary>
        ''' <param name="field">数値項目</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="scale">桁数（文字数）のうち小数桁数</param>
        ''' <returns>固定長ルールのLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function Number(ByVal field As Object, ByVal length As Integer, ByVal scale As Integer) As IFixedRuleLocator

        ''' <summary>
        ''' 数値項目を設定する
        ''' </summary>
        ''' <typeparam name="E">繰り返しの型</typeparam>
        ''' <param name="field">数値項目</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="scale">桁数（文字数）のうち小数桁数</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <returns>固定長ルールのLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function NumberRepeat(Of E)(ByVal field As ICollection(Of E), ByVal length As Integer, ByVal scale As Integer, ByVal repeat As Integer) As IFixedRuleLocator

        ''' <summary>
        ''' （繰り返し単位などの）グループ化を設定する
        ''' </summary>
        ''' <typeparam name="T">グループ化設定する型</typeparam>
        ''' <param name="groupField">グループ化項目</param>
        ''' <param name="groupRule">グループ化項目の固定長ルール</param>
        ''' <returns>固定長ルールのLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function Group(Of T)(ByVal groupField As T, ByVal groupRule As IFixedRule(Of T)) As IFixedRuleLocator

        ''' <summary>
        ''' （繰り返し単位などの）グループ化を設定する
        ''' </summary>
        ''' <typeparam name="T">グループ化設定する型</typeparam>
        ''' <param name="groupCollectionField">グループ化項目</param>
        ''' <param name="groupRule">グループ化項目の固定長ルール</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <returns>固定長ルールのLocatorインターフェース</returns>
        ''' <remarks></remarks>
        Function GroupRepeat(Of T)(ByVal groupCollectionField As ICollection(Of T), ByVal groupRule As IFixedRule(Of T), ByVal repeat As Integer) As IFixedRuleLocator

    End Interface
End Namespace