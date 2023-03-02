Namespace Util.Fixed

    Public Interface IFixedEntry

        ''' <summary>Group内の開始位置offset</summary>
        ReadOnly Property Offset() As Integer

        ''' <summary>固定長の長さ（Groupの場合、内包する長さ）</summary>
        ReadOnly Property Length() As Integer

        ''' <summary>名前</summary>
        ReadOnly Property Name() As String

        ''' <summary>繰り返し数</summary>
        ReadOnly Property Repeat() As Integer

        ''' <summary>
        ''' 子属性名があるか？を返す
        ''' </summary>
        ''' <param name="childName">子属性名</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function ContainsName(ByVal childName As String) As Boolean

        ''' <summary>
        ''' 子属性を返す
        ''' </summary>
        ''' <param name="childName">属性名</param>
        ''' <returns>属性（見つからない場合、nothing）</returns>
        ''' <remarks></remarks>
        Function GetChlid(ByVal childName As String) As IFixedEntry

        ''' <summary>
        ''' 開始位置offsetを初期化する
        ''' </summary>
        ''' <param name="offset">開始位置offset</param>
        ''' <remarks></remarks>
        Sub InitializeOffset(ByVal offset As Integer)

        ''' <summary>
        ''' 値を固定長文字列にする
        ''' </summary>
        ''' <param name="value">値</param>
        ''' <returns>固定長文字列</returns>
        ''' <remarks></remarks>
        Function Format(ByVal value As Object) As String

        ''' <summary>
        ''' 固定長文字列を値にする
        ''' </summary>
        ''' <param name="fixedString">固定長文字列</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Function Parse(ByVal fixedString As String) As Object

    End Interface
End Namespace