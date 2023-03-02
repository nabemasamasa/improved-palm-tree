Imports System.Reflection

Namespace Util.Csv
    ''' <summary>
    ''' Voアクセス用のCSV列設定情報をもつクラス
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class CsvRuleProperty
        ''' <summary>属性名</summary>
        Public Name As String

        ''' <summary>属性アクセス用のPropertyInfo</summary>
        Public Info As PropertyInfo

        ''' <summary>繰返し数</summary>
        Public Repeat As Integer

        ''' <summary>CSV列順設定の管理を担うOwner</summary>
        Public Builder As CsvRuleBuilder

        ''' <summary>CSV列の見出し名</summary>
        Public Title As String

        ''' <summary>CSV出力用の装飾処理</summary>
        Public ToCsvDecorator As Func(Of Object, String)

        ''' <summary>Vo変換用の装飾処理</summary>
        Public ToVoDecorator As Func(Of String, Object)

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="title">CSV列の見出し名</param>
        ''' <remarks></remarks>
        Public Sub New(title As String)
            Me.New(Nothing, Nothing)
            Me.Title = title
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal info As PropertyInfo)
            Me.New(name, info, 1)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal info As PropertyInfo, ByVal repeat As Integer)
            Me.New(name, info, repeat, Nothing)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <param name="builder">CSV列順設定の管理を担うOwner</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal info As PropertyInfo, ByVal repeat As Integer, ByVal builder As CsvRuleBuilder)
            Me.Name = name
            Me.Info = info
            Me.Repeat = repeat
            Me.Builder = builder
        End Sub

        ''' <summary>
        ''' グループ設定かを返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function IsGroup() As Boolean
            Return Builder IsNot Nothing
        End Function

        ''' <summary>
        ''' 見出しのみの項目か
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsTitleOnly() As Boolean
            Return Info Is Nothing
        End Function
    End Class
End Namespace