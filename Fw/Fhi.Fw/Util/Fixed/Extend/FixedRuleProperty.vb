Imports System.Reflection

Namespace Util.Fixed.Extend
    ''' <summary>
    ''' Voアクセス用の固定長列設定情報をもつクラス
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class FixedRuleProperty
        ''' <summary>属性名</summary>
        Public Name As String

        ''' <summary>属性アクセス用のPropertyInfo</summary>
        Public Info As PropertyInfo

        ''' <summary>繰返し数</summary>
        Public Repeat As Integer

        ''' <summary>固定長列設定の管理を担うOwner</summary>
        Public Builder As FixedRuleBuilder

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
        ''' <param name="builder">固定長列設定の管理を担うOwner</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal info As PropertyInfo, ByVal repeat As Integer, ByVal builder As FixedRuleBuilder)
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
    End Class
End Namespace