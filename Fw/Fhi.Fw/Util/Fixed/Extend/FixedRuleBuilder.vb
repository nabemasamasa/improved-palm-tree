Imports System.Reflection

Namespace Util.Fixed.Extend
    ''' <summary>
    ''' 固定長列設定情報の構築を担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class FixedRuleBuilder

        Private ReadOnly fixedEntries As New List(Of IFixedEntry)

        Private ReadOnly OrdinalProperties As New List(Of FixedRuleProperty)

        ''' <summary>固定長列設定情報[]</summary>
        Public ReadOnly Property ResultEntries() As IFixedEntry()
            Get
                Return fixedEntries.ToArray
            End Get
        End Property

        ''' <summary>Voアクセス用の固定長列設定情報[]</summary>
        Friend ReadOnly Property ResultProperties() As FixedRuleProperty()
            Get
                Return OrdinalProperties.ToArray
            End Get
        End Property

        ''' <summary>
        ''' グループ化列設定を作成する
        ''' </summary>
        ''' <param name="name">グループ名</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <returns>グループ化列設定</returns>
        ''' <remarks></remarks>
        Public Function MakeGroup(ByVal name As String, ByVal repeat As Integer) As FixedGroup
            Return New FixedGroup(name, repeat, fixedEntries.ToArray)
        End Function

        ''' <summary>
        ''' （Rootとなる）グループ化列設定を作成する
        ''' </summary>
        ''' <returns>（Rootとなる）グループ化列設定</returns>
        ''' <remarks></remarks>
        Public Function MakeRootGroup() As FixedGroup
            Return MakeGroup(Nothing, 1)
        End Function

        ''' <summary>
        ''' 固定長列設定の属性情報を追加する
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <remarks></remarks>
        Protected Sub AddBag(ByVal name As String, ByVal info As PropertyInfo)
            OrdinalProperties.Add(New FixedRuleProperty(name, info))
        End Sub

        ''' <summary>
        ''' 半角項目の固定長列設定情報を追加する
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="length">桁数(文字数)</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <remarks></remarks>
        Protected Sub AddHankaku(ByVal name As String, ByVal info As PropertyInfo, ByVal length As Integer, ByVal repeat As Integer)
            AddField(name, info, length, False, repeat)
        End Sub

        ''' <summary>
        ''' 半角項目の固定長列設定情報を追加する
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="length">桁数(文字数)</param>
        ''' <param name="decorateToVo">固定長文字列からVoへ変換時の装飾処理</param>
        ''' <param name="decorateToString">Voから固定長文字列へ変換時の装飾処理</param>
        ''' <remarks></remarks>
        Public Sub AddHankaku(ByVal name As String, ByVal info As PropertyInfo, ByVal length As Integer, ByVal decorateToVo As Func(Of String, Object), ByVal decorateToString As Func(Of Object, String))
            fixedEntries.Add(New FixedDecorator(name, length, False, 1, decorateToVo, decorateToString))
            OrdinalProperties.Add(New FixedRuleProperty(name, info, 1))
        End Sub

        ''' <summary>
        ''' 半角項目の固定長列設定情報を追加する
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="decorateToVo">固定長文字列からVoへ変換時の装飾処理</param>
        ''' <param name="decorateToString">Voから固定長文字列へ変換時の装飾処理</param>
        ''' <remarks></remarks>
        Public Sub AddHankaku(ByVal name As String, ByVal info As PropertyInfo, ByVal decorateToVo As Func(Of String, Object), ByVal decorateToString As Func(Of Object, String))
            fixedEntries.Add(New FixedDecorator(name, False, 1, decorateToVo, decorateToString))
            OrdinalProperties.Add(New FixedRuleProperty(name, info, 1))
        End Sub

        ''' <summary>
        ''' 全角項目の固定長列設定情報を追加する
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="length">桁数(文字数)</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <remarks></remarks>
        Protected Sub AddZenkaku(ByVal name As String, ByVal info As PropertyInfo, ByVal length As Integer, ByVal repeat As Integer)
            AddField(name, info, length, True, repeat)
        End Sub

        ''' <summary>
        ''' 文字項目の固定長列設定情報を追加する
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="length">桁数(文字数)</param>
        ''' <param name="isZenkaku">全角の場合、true</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <remarks></remarks>
        Private Sub AddField(ByVal name As String, ByVal info As PropertyInfo, ByVal length As Integer, ByVal isZenkaku As Boolean, ByVal repeat As Integer)
            fixedEntries.Add(New FixedField(name, length, isZenkaku, repeat))
            OrdinalProperties.Add(New FixedRuleProperty(name, info, repeat))
        End Sub

        ''' <summary>
        ''' 数値項目の固定長列設定情報を追加する
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="length">桁数(文字数)</param>
        ''' <param name="scale">文字数のうち小数桁数</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <remarks></remarks>
        Protected Sub AddNumber(ByVal name As String, ByVal info As PropertyInfo, ByVal length As Integer, ByVal scale As Integer, ByVal repeat As Integer)
            fixedEntries.Add(New FixedNumber(name, info.PropertyType, length, scale, repeat))
            OrdinalProperties.Add(New FixedRuleProperty(name, info, repeat))
        End Sub

        ''' <summary>
        ''' 固定長列設定の属性情報を追加する
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="info">属性アクセス用のPropertyInfo</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <param name="owner">固定長列設定の管理を担うOwner</param>
        ''' <remarks></remarks>
        Protected Sub AddGroup(ByVal name As String, ByVal info As PropertyInfo, ByVal repeat As Integer, ByVal owner As FixedRuleBuilder)
            If Not TypeUtil.IsTypeCollection(info.PropertyType) AndAlso 1 < repeat Then
                Throw New InvalidOperationException(String.Format("グループ項目プロパティ {0} は、配列型ではないが、繰り返し数に {1} が指定されている. 配列型以外のグループ項目は繰り返し数を 1 にすべき.", info.Name, repeat))
            End If
            If repeat < 1 Then
                Throw New InvalidOperationException(String.Format("グループ項目プロパティ {0} の繰り返し数に {1} が指定されている. 繰り返し数は 1 以上にすべき.", info.Name, repeat))
            End If
            fixedEntries.Add(owner.MakeGroup(name, repeat))
            OrdinalProperties.Add(New FixedRuleProperty(name, info, repeat, owner))
        End Sub
    End Class

    ''' <summary>
    ''' 固定長列設定情報の構築を担うクラス
    ''' </summary>
    ''' <typeparam name="T">固定長列設定を行う型</typeparam>
    ''' <remarks></remarks>
    Public Class FixedRuleBuilder(Of T) : Inherits FixedRuleBuilder

        Private ReadOnly rule As IFixedRule(Of T)

        Private ReadOnly voMarker As New VoPropertyMarker

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="rule">固定長列設定のルール</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal rule As IFixedRule(Of T))
            Me.rule = rule

            Dim aType As Type = GetType(T)
            Dim vo As T = CType(Activator.CreateInstance(aType), T)
            voMarker.MarkVo(vo)

            rule.Configure(New LocatorImpl(Me), vo)
        End Sub

        ''' <summary>
        ''' Group化項目列設定のメッセージ
        ''' </summary>
        ''' <typeparam name="G">グループ化項目の型</typeparam>
        ''' <param name="value">グループ化項目値</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <param name="rule">グループ化項目の固定長列設定</param>
        ''' <remarks></remarks>
        Friend Sub MessageGroup(Of G)(ByVal value As G, ByVal repeat As Integer, ByVal rule As IFixedRule(Of G))
            Dim owner As New FixedRuleBuilder(Of G)(rule)
            AddGroup(voMarker.GetPropertyInfo(value).Name, voMarker.GetPropertyInfo(value), repeat, owner)
        End Sub

        ''' <summary>
        ''' Group化項目列設定のメッセージ
        ''' </summary>
        ''' <typeparam name="G">グループ化項目の型</typeparam>
        ''' <param name="value">グループ化項目値</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <param name="rule">グループ化項目の固定長列設定</param>
        ''' <remarks></remarks>
        Friend Sub MessageGroup(Of G)(ByVal value As ICollection(Of G), ByVal repeat As Integer, ByVal rule As IFixedRule(Of G))
            Dim owner As New FixedRuleBuilder(Of G)(rule)
            AddGroup(voMarker.GetPropertyInfo(value).Name, voMarker.GetPropertyInfo(value), repeat, owner)
        End Sub

        ''' <summary>
        ''' 半角項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">半角項目の値</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <remarks></remarks>
        Friend Sub MessageHankaku(ByVal value As Object, ByVal length As Integer)
            MessageHankakuRepeat(value, length, 1)
        End Sub

        ''' <summary>
        ''' 半角項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">半角項目の値</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <remarks></remarks>
        Friend Sub MessageHankakuRepeat(ByVal value As Object, ByVal length As Integer, ByVal repeat As Integer)
            AddHankaku(voMarker.GetPropertyInfo(value).Name, voMarker.GetPropertyInfo(value), length, repeat)
        End Sub


        ''' <summary>
        ''' 半角項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">半角項目の値</param>
        ''' <param name="decorateToVo">固定長文字列からVoへ変換時の装飾処理</param>
        ''' <param name="decorateToString">Voから固定長文字列へ変換時の装飾処理</param>
        ''' <remarks></remarks>
        Friend Sub MessageHankaku(ByVal value As Object, ByVal decorateToVo As Func(Of String, Object), ByVal decorateToString As Func(Of Object, String))
            AddHankaku(voMarker.GetPropertyInfo(value).Name, voMarker.GetPropertyInfo(value), decorateToVo, decorateToString)
        End Sub

        ''' <summary>
        ''' 半角項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">半角項目の値</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="decorateToVo">固定長文字列からVoへ変換時の装飾処理</param>
        ''' <param name="decorateToString">Voから固定長文字列へ変換時の装飾処理</param>
        ''' <remarks></remarks>
        Friend Sub MessageHankaku(ByVal value As Object, ByVal length As Integer, ByVal decorateToVo As Func(Of String, Object), ByVal decorateToString As Func(Of Object, String))
            AddHankaku(voMarker.GetPropertyInfo(value).Name, voMarker.GetPropertyInfo(value), length, decorateToVo, decorateToString)
        End Sub

        ''' <summary>
        ''' 全角項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">全角項目の値</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <remarks></remarks>
        Friend Sub MessageZenkaku(ByVal value As Object, ByVal length As Integer)
            MessageZenkakuRepeat(value, length, 1)
        End Sub

        ''' <summary>
        ''' 全角項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">全角項目の値</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="repeat">繰り返し数</param>
        ''' <remarks></remarks>
        Friend Sub MessageZenkakuRepeat(ByVal value As Object, ByVal length As Integer, ByVal repeat As Integer)
            AddZenkaku(voMarker.GetPropertyInfo(value).Name, voMarker.GetPropertyInfo(value), length, repeat)
        End Sub

        ''' <summary>
        ''' 数値項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">数値項目の値</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="scale">桁数（文字数）のうち小数桁</param>
        ''' <remarks></remarks>
        Friend Sub MessageNumber(ByVal value As Object, ByVal length As Integer, ByVal scale As Integer)
            MessageNumberRepeat(value, length, scale, 1)
        End Sub

        ''' <summary>
        ''' 数値項目列設定のメッセージ
        ''' </summary>
        ''' <param name="value">数値項目の値</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="scale">桁数（文字数）のうち小数桁</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <remarks></remarks>
        Friend Sub MessageNumberRepeat(ByVal value As Object, ByVal length As Integer, ByVal scale As Integer, ByVal repeat As Integer)
            AddNumber(voMarker.GetPropertyInfo(value).Name, voMarker.GetPropertyInfo(value), length, scale, repeat)
        End Sub

        Private Class LocatorImpl : Implements IFixedRuleLocator

            Private enable As Boolean = True
            Private ReadOnly builder As FixedRuleBuilder(Of T)

            Public Sub New(ByVal builder As FixedRuleBuilder(Of T))
                Me.builder = builder
            End Sub

            Private Sub AssertEnable()
                If Not enable Then
                    Throw New InvalidOperationException("固定長文字列の最後まで処理されています")
                End If
            End Sub

            Private Sub ValidateLength(ByVal length As Integer)
                If length = AbstractFixedDefine.LENGTH_AS_OTHERS Then
                    enable = False
                End If
            End Sub

            ''' <summary>
            ''' 半角項目を設定する
            ''' </summary>
            ''' <param name="field">半角項目</param>
            ''' <param name="decorateToVo">固定長文字列からVoへ変換時の装飾処理</param>
            ''' <param name="decorateToString">Voから固定長文字列へ変換時の装飾処理</param>
            ''' <returns>固定長ルールのLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function Hankaku(ByVal field As Object, Optional ByVal decorateToVo As Func(Of String, Object) = Nothing, Optional ByVal decorateToString As Func(Of Object, String) = Nothing) As IFixedRuleLocator Implements IFixedRuleLocator.Hankaku
                AssertEnable()
                enable = False
                builder.MessageHankaku(field, decorateToVo, decorateToString)
                Return Me
            End Function

            ''' <summary>
            ''' 半角項目を設定する
            ''' </summary>
            ''' <param name="field">半角項目</param>
            ''' <param name="length">桁数（文字数）</param>
            ''' <param name="decorateToVo">固定長文字列からVoへ変換時の装飾処理</param>
            ''' <param name="decorateToString">Voから固定長文字列へ変換時の装飾処理</param>
            ''' <returns>固定長ルールのLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function Hankaku(ByVal field As Object, ByVal length As Integer, Optional ByVal decorateToVo As Func(Of String, Object) = Nothing, Optional ByVal decorateToString As Func(Of Object, String) = Nothing) As IFixedRuleLocator Implements IFixedRuleLocator.Hankaku
                AssertEnable()
                ValidateLength(length)
                If decorateToString Is Nothing AndAlso decorateToVo Is Nothing Then
                    builder.MessageHankaku(field, length)
                Else
                    builder.MessageHankaku(field, length, decorateToVo, decorateToString)
                End If
                Return Me
            End Function

            ''' <summary>
            ''' 半角項目を設定する
            ''' </summary>
            ''' <typeparam name="E">繰り返しの型</typeparam>
            ''' <param name="field">半角項目</param>
            ''' <param name="length">桁数（文字数）</param>
            ''' <param name="repeat">繰り返し数</param>
            ''' <returns>固定長ルールのLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function HankakuRepeat(Of E)(ByVal field As ICollection(Of E), ByVal length As Integer, ByVal repeat As Integer) As IFixedRuleLocator Implements IFixedRuleLocator.HankakuRepeat
                AssertEnable()
                ValidateLength(length)
                builder.MessageHankakuRepeat(field, length, repeat)
                Return Me
            End Function

            ''' <summary>
            ''' 全角項目を設定する
            ''' </summary>
            ''' <param name="field">全角項目</param>
            ''' <param name="length">桁数（文字数）</param>
            ''' <returns>固定長ルールのLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function Zenkaku(ByVal field As Object, ByVal length As Integer) As IFixedRuleLocator Implements IFixedRuleLocator.Zenkaku
                AssertEnable()
                ValidateLength(length)
                builder.MessageZenkaku(field, length)
                Return Me
            End Function

            ''' <summary>
            ''' 全角項目を設定する
            ''' </summary>
            ''' <typeparam name="E">繰り返しの型</typeparam>
            ''' <param name="field">全角項目</param>
            ''' <param name="length">桁数（文字数）</param>
            ''' <param name="repeat">繰り返し数</param>
            ''' <returns>固定長ルールのLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function ZenkakuRepeat(Of E)(ByVal field As ICollection(Of E), ByVal length As Integer, ByVal repeat As Integer) As IFixedRuleLocator Implements IFixedRuleLocator.ZenkakuRepeat
                AssertEnable()
                ValidateLength(length)
                builder.MessageZenkakuRepeat(field, length, repeat)
                Return Me
            End Function

            ''' <summary>
            ''' 数値項目を設定する
            ''' </summary>
            ''' <param name="field">数値項目</param>
            ''' <param name="length">桁数（文字数）</param>
            ''' <param name="scale">桁数（文字数）のうち小数桁数</param>
            ''' <returns>固定長ルールのLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function Number(ByVal field As Object, ByVal length As Integer, ByVal scale As Integer) As IFixedRuleLocator Implements IFixedRuleLocator.Number
                AssertEnable()
                ValidateLength(length)
                builder.MessageNumber(field, length, scale)
                Return Me
            End Function

            ''' <summary>
            ''' 数値項目を設定する
            ''' </summary>
            ''' <typeparam name="E">繰り返しの型</typeparam>
            ''' <param name="field">数値項目</param>
            ''' <param name="length">桁数（文字数）</param>
            ''' <param name="scale">桁数（文字数）のうち小数桁数</param>
            ''' <param name="repeat">繰り返し数</param>
            ''' <returns>固定長ルールのLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function NumberRepeat(Of E)(ByVal field As ICollection(Of E), ByVal length As Integer, ByVal scale As Integer, ByVal repeat As Integer) As IFixedRuleLocator Implements IFixedRuleLocator.NumberRepeat
                AssertEnable()
                ValidateLength(length)
                builder.MessageNumberRepeat(field, length, scale, repeat)
                Return Me
            End Function

            ''' <summary>
            ''' （繰り返し単位などの）グループ化を設定する
            ''' </summary>
            ''' <typeparam name="T1">グループ化設定する型</typeparam>
            ''' <param name="groupField">グループ化項目</param>
            ''' <param name="groupRule">グループ化項目の固定長ルール</param>
            ''' <returns>固定長ルールのLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function Group(Of T1)(ByVal groupField As T1, ByVal groupRule As IFixedRule(Of T1)) As IFixedRuleLocator Implements IFixedRuleLocator.Group
                AssertEnable()
                builder.MessageGroup(groupField, 1, groupRule)
                Return Me
            End Function

            ''' <summary>
            ''' （繰り返し単位などの）グループ化を設定する
            ''' </summary>
            ''' <typeparam name="T1">グループ化設定する型</typeparam>
            ''' <param name="groupCollectionField">グループ化項目</param>
            ''' <param name="groupRule">グループ化項目の固定長ルール</param>
            ''' <returns>固定長ルールのLocatorインターフェース</returns>
            ''' <remarks></remarks>
            Public Function GroupRepeat(Of T1)(ByVal groupCollectionField As ICollection(Of T1), ByVal groupRule As IFixedRule _
                                                  (Of T1), ByVal repeat As Integer) As IFixedRuleLocator Implements IFixedRuleLocator.GroupRepeat
                AssertEnable()
                builder.MessageGroup(groupCollectionField, repeat, groupRule)
                Return Me
            End Function

        End Class

    End Class
End Namespace