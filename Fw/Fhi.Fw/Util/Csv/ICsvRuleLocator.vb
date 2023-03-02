Namespace Util.Csv
    ''' <summary>
    ''' CSVの列定義を担うインターフェース
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface ICsvRuleLocator
        ''' <summary>
        ''' CSV項目の並び順を設定する
        ''' </summary>
        ''' <param name="fields">並び順に沿った項目[]</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Field(ByVal ParamArray fields As Object()) As ICsvRuleLocator

        ''' <summary>
        ''' CSV項目の繰り返しを設定する
        ''' </summary>
        ''' <param name="collectionField">繰り返し設定する項目</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <returns>自身</returns>
        ''' <remarks></remarks>
        Function FieldRepeat(Of E)(ByVal collectionField As ICollection(Of E), ByVal repeat As Integer) As ICsvRuleLocator

        ''' <summary>
        ''' 一まとめのCSV項目たちを設定する
        ''' </summary>
        ''' <typeparam name="G">「一まとめのCSV項目」のVo型</typeparam>
        ''' <param name="groupField">「一まとめのCSV項目」を設定したVo項目</param>
        ''' <param name="groupRule">「一まとめのCSV項目」の列定義</param>
        ''' <returns>自身</returns>
        ''' <remarks></remarks>
        Function Group(Of G)(ByVal groupField As G, ByVal groupRule As ICsvRule(Of G)) As ICsvRuleLocator

        ''' <summary>
        ''' 一まとめのCSV項目たちを設定する
        ''' </summary>
        ''' <typeparam name="G">「一まとめのCSV項目」のVo型</typeparam>
        ''' <param name="groupField">「一まとめのCSV項目」を設定したVo項目</param>
        ''' <param name="groupRule">「一まとめのCSV項目」の列定義</param>
        ''' <returns>自身</returns>
        ''' <remarks></remarks>
        Function Group(Of G)(ByVal groupField As G, ByVal groupRule As CsvRule(Of G).RuleConfigure) As ICsvRuleLocator

        ''' <summary>
        ''' 一まとめのCSV項目たちを繰り返し設定する
        ''' </summary>
        ''' <typeparam name="G">「一まとめのCSV項目」のVo型</typeparam>
        ''' <param name="groupCollectionField">「一まとめのCSV項目」を設定したVo項目のコレクション</param>
        ''' <param name="groupRule">「一まとめのCSV項目」の列定義</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GroupRepeat(Of G)(ByVal groupCollectionField As ICollection(Of G), ByVal groupRule As ICsvRule(Of G), ByVal repeat As Integer) As ICsvRuleLocator

        ''' <summary>
        ''' 一まとめのCSV項目たちを繰り返し設定する
        ''' </summary>
        ''' <typeparam name="G">「一まとめのCSV項目」のVo型</typeparam>
        ''' <param name="groupCollectionField">「一まとめのCSV項目」を設定したVo項目のコレクション</param>
        ''' <param name="groupRule">「一まとめのCSV項目」の列定義</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GroupRepeat(Of G)(ByVal groupCollectionField As ICollection(Of G), ByVal groupRule As CsvRule(Of G).RuleConfigure, ByVal repeat As Integer) As ICsvRuleLocator

        ''' <summary>
        ''' 空列を追加する
        ''' </summary>
        ''' <param name="title">Title出力用の項目名</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function TitleOnly(title As String) As ICsvRuleLocator

        ''' <summary>
        ''' CSV項目用に装飾を施す
        ''' </summary>
        ''' <param name="field">項目</param>
        ''' <param name="toCsvDecorator">CSV出力時の装飾処理</param>
        ''' <param name="toVoDecorator">Vo変換時の装飾処理</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function FieldWithDecorator(field As Object, Optional toCsvDecorator As Func(Of Object, String) = Nothing, Optional toVoDecorator As Func(Of String, Object) = Nothing) As ICsvRuleLocator

    End Interface
End Namespace