Namespace TestUtil.DebugString
    ''' <summary>
    ''' 検証文字列作成を担うワイルドカード
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface DebugStringMakerWildcard

        ''' <summary>
        ''' 検証用値情報のタイトル一覧を作成する
        ''' </summary>
        ''' <param name="parentTitle">親タイトル</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function MakeTitles(Optional parentTitle As String = Nothing) As String()

        ''' <summary>
        ''' 検証用値情報を作成して格納する
        ''' </summary>
        ''' <param name="record"></param>
        ''' <remarks></remarks>
        Sub StoreAfterMaking(record As Object)

        ''' <summary>
        ''' 格納情報から値情報を構築して格納情報をクリアするCallbackを返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetCallbackThatBuildValuesAndClearStored() As Func(Of String()())

        ''' <summary>
        ''' Empty値の値情報作成Callbackを返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetCallbackThatMakeEmptyValues() As Func(Of String())
    End Interface
End NameSpace