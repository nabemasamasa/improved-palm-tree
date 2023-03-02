Namespace App.Xls
    ''' <summary>
    ''' データのプロパティを選択する
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface XlVoSelector
        ''' <summary>
        ''' どのプロパティを使用するか決定する
        ''' </summary>
        ''' <param name="prop">プロパティ</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Use(prop As Object) As XlVoSelector

        ''' <summary>
        ''' どのプロパティを使用するか決定する
        ''' </summary>
        ''' <param name="props">プロパティ</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Use(ParamArray props As Object()) As XlVoSelector

        ''' <summary>
        ''' 変換時の装飾及びプロパティを設定する<br/>※装飾は必ずどちらか設定する事
        ''' </summary>
        ''' <param name="prop">プロパティ</param>
        ''' <param name="toVoDecorator">(任意) Voへ設定する時の装飾処理</param>
        ''' <param name="toXlsDecorator">(任意) Excelへ出力する時の装飾処理</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function UseWithFunc(ByVal prop As Object, Optional ByVal toVoDecorator As Func(Of Object, Object) = Nothing, Optional ByVal toXlsDecorator As Func(Of Object, Object) = Nothing) As XlVoSelector
    End Interface
End Namespace