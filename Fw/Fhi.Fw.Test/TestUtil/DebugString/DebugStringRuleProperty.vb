Imports System.Reflection
Imports System.Linq.Expressions

Namespace TestUtil.DebugString
    ''' <summary>
    ''' 検証用文字列作成のための項目情報をもつクラス
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class DebugStringRuleProperty
        ''' <summary>属性名</summary>
        Public Name As String

        ''' <summary>属性取得Callback</summary>
        Public GetValueCallback As Func(Of Object, Object)

        ''' <summary>出力タイトル</summary>
        Private ReadOnly title As String

        ''' <summary>出力値のLambda</summary>
        Public Lambda As LambdaExpression

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="title">出力タイトル</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal info As PropertyInfo, ByVal title As String)
            Me.New(name, Function(vo) info.GetValue(vo, Nothing), Nothing, title)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="getValueCallback">属性取得Callback</param>
        ''' <param name="title">出力タイトル</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal getValueCallback As Func(Of Object, Object), ByVal title As String)
            Me.New(name, getValueCallback, Nothing, title)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="lambda">出力値のLambda</param>
        ''' <param name="title">出力タイトル</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal lambda As LambdaExpression, ByVal title As String)
            Me.New(name, Nothing, lambda, title)
        End Sub

        Private Sub New(ByVal name As String, ByVal getValueCallback As Func(Of Object, Object), ByVal lambda As LambdaExpression, ByVal title As String)
            Me.Name = name
            Me.GetValueCallback = getValueCallback
            Me.Lambda = lambda
            Me.title = title
        End Sub

        ''' <summary>
        ''' 出力タイトルを（未指定なら）作成する
        ''' </summary>
        ''' <param name="parentTitle">親タイトル</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MakeTitleIfNecessary(parentTitle As String) As String
            Const SEPARATOR As String = "."
            If StringUtil.IsNotEmpty(title) Then
                Return title
            ElseIf StringUtil.IsEmpty(parentTitle) Then
                Return Name
            End If
            Dim names As List(Of String) = Split(parentTitle, SEPARATOR).Union({Name}).Where(Function(v) StringUtil.IsNotEmpty(v)).ToList
            Return Join(Enumerable.Range(0, names.Count).Select(Function(i) If(i = names.Count - 1, names(i), names(i).Substring(0, 1))).ToArray, SEPARATOR)
        End Function

    End Class
End Namespace