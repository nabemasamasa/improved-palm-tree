Namespace Util.VoCopy
    Public Interface IVoCopyPropertyDefine

        ''' <summary>
        ''' コピーするプロパティを設定する
        ''' </summary>
        ''' <param name="x">関連づけるプロパティ値</param>
        ''' <param name="y">関連づけるプロパティ値</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Bind(ByVal x As Object, ByVal y As Object) As IVoCopyPropertyDefine

    End Interface
End Namespace