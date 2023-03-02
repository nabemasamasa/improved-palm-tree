Imports System.Reflection

Namespace Util.Csv
    Public Class CsvRuleBuilder

        Private ReadOnly OrdinalProperties As New List(Of CsvRuleProperty)

        ''' <summary>Voアクセス用の区切り列順設定情報[]</summary>
        Friend ReadOnly Property ResultProperties() As CsvRuleProperty()
            Get
                Return OrdinalProperties.ToArray
            End Get
        End Property

        Protected Sub Add(ByVal name As String, ByVal info As PropertyInfo)
            OrdinalProperties.Add(New CsvRuleProperty(name, info))
        End Sub

        Protected Sub AddRepeat(ByVal name As String, ByVal info As PropertyInfo, ByVal repeat As Integer)
            OrdinalProperties.Add(New CsvRuleProperty(name, info, repeat))
        End Sub

        ''' <summary>
        ''' 固定長列設定の属性情報を追加する
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <param name="owner">固定長列設定の管理を担うOwner</param>
        ''' <remarks></remarks>
        Protected Sub AddGroup(ByVal name As String, ByVal info As PropertyInfo, ByVal repeat As Integer, ByVal owner As CsvRuleBuilder)
            If Not TypeUtil.IsTypeCollection(info.PropertyType) AndAlso 1 < repeat Then
                Throw New InvalidOperationException(String.Format("グループ項目プロパティ {0} は、配列型ではないが、繰り返し数に {1} が指定されている. 配列型以外のグループ項目は繰り返し数を 1 にすべき.", info.Name, repeat))
            End If
            If repeat < 1 Then
                Throw New InvalidOperationException(String.Format("グループ項目プロパティ {0} の繰り返し数に {1} が指定されている. 繰り返し数は 1 以上にすべき.", info.Name, repeat))
            End If
            OrdinalProperties.Add(New CsvRuleProperty(name, info, repeat, owner))
        End Sub

        Protected Sub AddTitleOnly(title As String)
            OrdinalProperties.Add(New CsvRuleProperty(title))
        End Sub

        Protected Sub AddWithDecorator(name As String, ByVal info As PropertyInfo, toCsvDecorator As Func(Of Object, String), toVoDecorator As Func(Of String, Object))
            If toCsvDecorator Is Nothing AndAlso toVoDecorator Is Nothing Then
                Throw New ArgumentException("WithDecoratorなのに装飾処理が何も指定されてない")
            End If
            Dim rule As New CsvRuleProperty(name, info)
            rule.ToCsvDecorator = toCsvDecorator
            rule.ToVoDecorator = toVoDecorator
            OrdinalProperties.Add(rule)
        End Sub

    End Class

    Public Class CsvRuleBuilder(Of T) : Inherits CsvRuleBuilder

        Private ReadOnly rule As ICsvRule(Of T)

        Private ReadOnly voMarker As New VoPropertyMarker

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">区切り列順設定のルール</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As CsvRule(Of T).RuleConfigure)
            Me.New(New CsvRule(Of T)(rule))
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">区切り列順設定のルール</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As ICsvRule(Of T))
            Me.rule = rule

            Dim aType As Type = GetType(T)
            Dim vo As T = CType(Activator.CreateInstance(aType), T)
            voMarker.MarkVo(vo)

            rule.Configure(New LocatorImpl(Me), vo)
        End Sub

        ''' <summary>
        ''' 半角項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">半角項目の値</param>
        ''' <remarks></remarks>
        Private Sub Message(ByVal value As Object)

            Add(voMarker.GetPropertyInfo(value).Name, voMarker.GetPropertyInfo(value))
        End Sub

        ''' <summary>
        ''' Group化項目列設定のメッセージ
        ''' </summary>
        ''' <typeparam name="G">グループ化項目の型</typeparam>
        ''' <param name="values">グループ化項目値</param>
        ''' <param name="rule">グループ化項目の固定長列設定</param>
        ''' <remarks></remarks>
        Friend Sub MessageGroup(Of G)(ByVal values As G, ByVal rule As ICsvRule(Of G))
            Dim owner As New CsvRuleBuilder(Of G)(rule)
            AddGroup(voMarker.GetPropertyInfo(values).Name, voMarker.GetPropertyInfo(values), 1, owner)
        End Sub

        ''' <summary>
        ''' Group化項目列設定のメッセージ
        ''' </summary>
        ''' <typeparam name="G">グループ化項目の型</typeparam>
        ''' <param name="values">グループ化項目値</param>
        ''' <param name="rule">グループ化項目の固定長列設定</param>
        ''' <remarks></remarks>
        Friend Sub MessageGroup(Of G)(ByVal values As G, ByVal rule As CsvRule(Of G).RuleConfigure)
            Dim owner As New CsvRuleBuilder(Of G)(rule)
            AddGroup(voMarker.GetPropertyInfo(values).Name, voMarker.GetPropertyInfo(values), 1, owner)
        End Sub

        ''' <summary>
        ''' Group化項目列設定のメッセージ（コレクション）
        ''' </summary>
        ''' <typeparam name="G">グループ化項目の型</typeparam>
        ''' <param name="values">グループ化項目値</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <param name="rule">グループ化項目の固定長列設定</param>
        ''' <remarks></remarks>
        Friend Sub MessageGroup(Of G)(ByVal values As ICollection(Of G), ByVal repeat As Integer, ByVal rule As ICsvRule(Of G))
            Dim owner As New CsvRuleBuilder(Of G)(rule)
            AddGroup(voMarker.GetPropertyInfo(values).Name, voMarker.GetPropertyInfo(values), repeat, owner)
        End Sub

        ''' <summary>
        ''' Group化項目列設定のメッセージ（コレクション）
        ''' </summary>
        ''' <typeparam name="G">グループ化項目の型</typeparam>
        ''' <param name="values">グループ化項目値</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <param name="rule">グループ化項目の固定長列設定</param>
        ''' <remarks></remarks>
        Friend Sub MessageGroup(Of G)(ByVal values As ICollection(Of G), ByVal repeat As Integer, ByVal rule As CsvRule(Of G).RuleConfigure)
            Dim owner As New CsvRuleBuilder(Of G)(rule)
            AddGroup(voMarker.GetPropertyInfo(values).Name, voMarker.GetPropertyInfo(values), repeat, owner)
        End Sub

        ''' <summary>
        ''' 繰り返し項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">項目値</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <remarks></remarks>
        Friend Sub MessageRepeat(ByVal value As Object, ByVal repeat As Integer)
            AddRepeat(voMarker.GetPropertyInfo(value).Name, voMarker.GetPropertyInfo(value), repeat)
        End Sub

        ''' <summary>
        ''' 見出し項目のみ列設定のメッセージ
        ''' </summary>
        ''' <param name="title"></param>
        ''' <remarks></remarks>
        Friend Sub MessageTitleOnly(title As String)
            AddTitleOnly(title)
        End Sub

        ''' <summary>
        ''' 装飾処理が必要な項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">項目値</param>
        ''' <param name="toCsvDecorator">CSV出力時の装飾処理</param>
        ''' <param name="toVoDecorator">Vo変換時の装飾処理</param>
        ''' <remarks></remarks>
        Friend Sub MessageWithDecorator(value As Object, toCsvDecorator As Func(Of Object, String), toVoDecorator As Func(Of String, Object))
            AddWithDecorator(voMarker.GetPropertyInfo(value).Name, voMarker.GetPropertyInfo(value), toCsvDecorator, toVoDecorator)
        End Sub

        Private Class LocatorImpl : Implements ICsvRuleLocator

            Private ReadOnly builder As CsvRuleBuilder(Of T)

            Public Sub New(ByVal builder As CsvRuleBuilder(Of T))
                Me.builder = builder
            End Sub

            Public Function Field(ByVal ParamArray fields As Object()) As ICsvRuleLocator Implements ICsvRuleLocator.Field
                For Each aField As Object In fields
                    builder.Message(aField)
                Next
                Return Me
            End Function

            Public Function FieldRepeat(Of E)(ByVal collectionField As ICollection(Of e), ByVal repeat As Integer) As ICsvRuleLocator Implements ICsvRuleLocator.FieldRepeat
                builder.MessageRepeat(collectionField, repeat)
                Return Me
            End Function

            Public Function Group(Of G)(ByVal groupField As G, ByVal groupRule As ICsvRule(Of G)) As ICsvRuleLocator Implements ICsvRuleLocator.Group
                builder.MessageGroup(groupField, groupRule)
                Return Me
            End Function

            Public Function Group(Of G)(ByVal groupField As G, ByVal groupRule As CsvRule(Of G).RuleConfigure) As ICsvRuleLocator Implements ICsvRuleLocator.Group
                builder.MessageGroup(groupField, groupRule)
                Return Me
            End Function

            Public Function GroupRepeat(Of G)(ByVal groupCollectionField As ICollection(Of G), ByVal groupRule As ICsvRule _
                                            (Of G), ByVal repeat As Integer) As ICsvRuleLocator Implements ICsvRuleLocator.GroupRepeat
                builder.MessageGroup(groupCollectionField, repeat, groupRule)
                Return Me
            End Function

            Public Function GroupRepeat(Of G)(ByVal groupCollectionField As ICollection(Of G), ByVal groupRule As CsvRule(Of G).RuleConfigure, _
                                              ByVal repeat As Integer) As ICsvRuleLocator Implements ICsvRuleLocator.GroupRepeat
                builder.MessageGroup(groupCollectionField, repeat, groupRule)
                Return Me
            End Function

            Public Function TitleOnly(ByVal title As String) As ICsvRuleLocator Implements ICsvRuleLocator.TitleOnly
                builder.MessageTitleOnly(title)
                Return Me
            End Function

            Public Function FieldWithDecorator(ByVal field As Object, Optional ByVal toCsvDecorator As Func(Of Object, String) = Nothing, Optional ByVal toVoDecorator As Func(Of String, Object) = Nothing) As ICsvRuleLocator Implements ICsvRuleLocator.FieldWithDecorator
                builder.MessageWithDecorator(field, toCsvDecorator, toVoDecorator)
                Return Me
            End Function
        End Class

    End Class
End Namespace