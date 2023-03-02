Imports Fhi.Fw.Util

Namespace Domain
    ''' <summary>
    ''' 不変のコレクションオブジェクトで変化点観察を担うクラス
    ''' </summary>
    ''' <typeparam name="T">要素</typeparam>
    ''' <remarks>
    ''' 要素`T`が
    '''   int値、String値、ValueObject値等の不変オブジェクトは、そのままコピーして変更前後を判断する
    '''   それ以外はPublicプロパティをシャローコピーして、変更前後を判断する
    ''' </remarks>
    Public Class CollectionJudgeerObject(Of T) : Inherits CollectionObject(Of T)
        ''' <summary>変化点判定</summary>
        Protected ReadOnly Judgeer As ChangeJudgeer(Of T)

        Public Sub New()
            MyBase.New()
            Me.Judgeer = BuildJudgeer()
        End Sub

        Public Sub New(ByVal initialList As IEnumerable(Of T))
            MyBase.New(initialList)
            Me.Judgeer = BuildJudgeer()
        End Sub

        Public Sub New(ByVal src As CollectionJudgeerObject(Of T))
            MyBase.New(src)
            Dim judgeer As ChangeJudgeer(Of T) = If(src Is Nothing, Nothing, src.Judgeer)
            Me.Judgeer = BuildJudgeer(judgeer)
        End Sub

        Private Function BuildJudgeer(Optional ByVal judgeer As ChangeJudgeer(Of T) = Nothing) As ChangeJudgeer(Of T)
            Return If(judgeer, New ChangeJudgeer(Of T)(Me.InternalList, New PluginImpl(Me)))
        End Function

        ''' <summary>
        ''' 複製する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overridable Overloads Function Clone() As CollectionJudgeerObject(Of T)
            Return Clone(Of CollectionJudgeerObject(Of T))()
        End Function
        ''' <summary>
        ''' 複製する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overridable Overloads Function Clone(Of TResult As CollectionJudgeerObject(Of T))() As TResult
            Return MyBase.Clone(Of TResult)()
        End Function

        Private Class PluginImpl : Implements ChangeJudgeer(Of T).Plugin
            Private ReadOnly this As CollectionJudgeerObject(Of T)
            Public Sub New(this As CollectionJudgeerObject(Of T))
                Me.this = this
            End Sub
            Public Function MakeKey(ByVal obj As T) As Object Implements ChangeJudgeer(Of T).Plugin.MakeKey
                Return this.MakeKeyForChangeJudgeer(obj)
            End Function
        End Class

        ''' <summary>
        ''' データを識別するキーを作成する
        ''' </summary>
        ''' <param name="obj">データ</param>
        ''' <returns>データを識別するキー値</returns>
        ''' <remarks></remarks>
        Protected Overridable Function MakeKeyForChangeJudgeer(obj As T) As Object
            Return obj
        End Function

        ''' <summary>
        ''' 変化点判定から除外するプリミティブ値プロパティに付与した属性を返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overridable Function GetAttributeTypesToIgnorePropertyFromChange() As Type()
            Return Nothing
        End Function

        ''' <summary>
        ''' 変化点判定に含めるインスタンス値プロパティに付与した属性を返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overridable Function GetAttributeTypesToIncludePropertyFromChange() As Type()
            Return Nothing
        End Function

        ''' <summary>
        ''' 変更前データを設定する
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub SupersedeBeforeUpdatedItems()
            Judgeer.IncludePropertyAttributes = GetAttributeTypesToIncludePropertyFromChange()
            Judgeer.SupersedeBeforeUpdatedVos(Me.InternalList)
        End Sub

        ''' <summary>
        ''' 変更後データを設定する
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub SetUpdatedItems()
            Judgeer.IgnorePropertyAttributes = GetAttributeTypesToIgnorePropertyFromChange()
            Judgeer.SetUpdatedVos(Me.InternalList)
        End Sub

        ''' <summary>
        ''' 変更後データが、変更後データでなくなった（保存処理完了など）ことを通知する
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub NotifySaveEnd()
            Judgeer.NotifySaveEnd()
        End Sub

        ''' <summary>
        ''' 変更前データと変更後データとに差異があるかを返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Overridable Function HasChanged() As Boolean
            Return Judgeer.HasChanged()
        End Function

        ''' <summary>
        ''' 削除データがあるかを返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function WasDeleted() As Boolean
            Return Judgeer.WasDeleted()
        End Function

        ''' <summary>
        ''' 削除データを抽出する
        ''' </summary>
        ''' <returns>削除されたデータ</returns>
        ''' <remarks></remarks>
        Public Function ExtractDeletedItems() As T()
            Return Judgeer.ExtractDeletedVos()
        End Function

        ''' <summary>
        ''' 追加データがあるかを返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function WasInserted() As Boolean
            Return Judgeer.WasInserted()
        End Function

        ''' <summary>
        ''' 追加データを抽出する
        ''' </summary>
        ''' <returns>追加したデータ</returns>
        ''' <remarks></remarks>
        Public Function ExtractInsertedItems() As T()
            Return Judgeer.ExtractInsertedVos()
        End Function

        ''' <summary>
        ''' 変更データがあるかを返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function WasUpdated() As Boolean
            Return Judgeer.WasUpdated()
        End Function

        ''' <summary>
        ''' 変更データを抽出する
        ''' </summary>
        ''' <returns>変更したデータ</returns>
        ''' <remarks></remarks>
        Public Function ExtractUpdatedItems() As ChangeJudgeer(Of T).UpdateInfo()
            Return Judgeer.ExtractUpdatedVos()
        End Function

        ''' <summary>
        ''' #ExtractXxxxで抽出した変更点Voから、Voの元々のインスタンスを返す（第三者によってインスタンス内の値は変更されている可能性がある）
        ''' </summary>
        ''' <param name="extractsVo">変更点Vo</param>
        ''' <returns>Voの元々のインスタンス</returns>
        ''' <remarks></remarks>
        Public Function DetectOriginInstance(ByVal extractsVo As T) As T
            Return Judgeer.DetectOriginInstance(extractsVo)
        End Function

    End Class
End Namespace