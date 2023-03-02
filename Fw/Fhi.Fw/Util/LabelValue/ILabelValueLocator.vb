Namespace Util.LabelValue
    ''' <summary>
    ''' Label と Value を設定するインターフェース
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface ILabelValueLocator
        ''' <summary>
        ''' Labelプロパティである事を設定する
        ''' </summary>
        ''' <param name="labelProperty">Labelプロパティ</param>
        ''' <returns>Label と Value を設定するMessengerインターフェース</returns>
        ''' <remarks></remarks>
        Function Label(ByVal labelProperty As Object) As ILabelValueLocator

        ''' <summary>
        ''' Valueプロパティである事を設定する
        ''' </summary>
        ''' <param name="valueProperty">Valueプロパティ</param>
        ''' <returns>Label と Value を設定するMessengerインターフェース</returns>
        ''' <remarks></remarks>
        Function Value(ByVal valueProperty As Object) As ILabelValueLocator
    End Interface
End Namespace
