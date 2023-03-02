Imports System.Reflection
Imports System.Threading

Namespace Util
    ''' <summary>
    ''' 変更前のデータと、変更後のデータとを比較し、変更点をリストアップするクラス
    ''' </summary>
    ''' <typeparam name="T">データの型</typeparam>
    ''' <remarks>
    ''' 1. 変更前のデータと同じインスタンスを、変更後データとして指定すること
    ''' 2. 読み書き可能なプリミティブ値プロパティで比較している（読取専用プロパティは対象外）
    '''     - インスタンス値プロパティも比較する手段に IncludePropertyAttributes がある
    ''' </remarks>
    Public Class ChangeJudgeer(Of T)

        ''' <summary>
        ''' 更新(修正)情報クラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class UpdateInfo
            ''' <summary>更新前情報</summary>
            Public BeforeVo As T
            ''' <summary>更新後情報</summary>
            Public AfterVo As T
            ''' <summary>コンストラクタ</summary>
            Public Sub New(ByVal beforeVo As T, ByVal afterVo As T)
                Me.BeforeVo = beforeVo
                Me.AfterVo = afterVo
            End Sub
        End Class

        Public Interface Plugin
            ''' <summary>
            ''' データを識別するキーを作成する
            ''' </summary>
            ''' <param name="obj">データ</param>
            ''' <returns>データを識別するキー値</returns>
            ''' <remarks></remarks>
            Function MakeKey(ByVal obj As T) As Object
        End Interface

        ''' <summary>
        ''' データを識別するキーを作成する
        ''' </summary>
        ''' <param name="obj">データ</param>
        ''' <returns>データを識別するキー値</returns>
        ''' <remarks>同レコード判定に使うキーを作成するデリゲート</remarks>
        Public Delegate Function MakeKeyCallback(ByVal obj As T) As Object

        Private Class DefaultPlugin : Implements Plugin
            Public Function MakeKey(ByVal obj As T) As Object Implements Plugin.MakeKey
                Return obj
            End Function
        End Class

        Private Class DefaultDlgtPlugin : Implements Plugin
            Private dlgtPlugin As MakeKeyCallback
            Public Sub New(ByVal dlgtPlugin As MakeKeyCallback)
                Me.dlgtPlugin = dlgtPlugin
            End Sub
            Public Function MakeKey(ByVal obj As T) As Object Implements Plugin.MakeKey
                Return dlgtPlugin.Invoke(obj)
            End Function
        End Class

        Private Class WrappedDefaultPluginImpl : Implements Plugin
            Private aPlugin As Plugin
            Public Sub New(ByVal aPlugin As Plugin)
                Me.aPlugin = aPlugin
            End Sub
            Public Function MakeKey(ByVal obj As T) As Object Implements Plugin.MakeKey
                Dim key As Object = aPlugin.MakeKey(obj)
                If key Is Nothing Then
                    Throw New InvalidOperationException(String.Format("Plugin実装クラス {0}" & vbCrLf & "#MakeKey() の結果が null は不正", aPlugin.GetType.FullName))
                End If
                Return key
            End Function
        End Class

        Private ReadOnly aPlugin As Plugin
        Private _updatedKeys As List(Of Object)
        Private _updatedVos As List(Of T)
        Private _updatedVoByKey As Dictionary(Of Object, T)
        Private _updatedInstanceByKey As Dictionary(Of Object, T)

        Private _beforeKeys As List(Of Object)
        Private _beforeVos As List(Of T)
        Private _beforeVoByKey As Dictionary(Of Object, T)
        Private _beforeInstanceByKey As Dictionary(Of Object, T)

        Private cacheRwLock As New ReaderWriterLockSlim

#Region "Public properties..."
        ''' <summary>不変オブジェクトならTrue ※Int32等は自動でTrueになる。自作不変オブジェクトは手動でTrueにすること</summary>
        Public IsTypeImmutable As Boolean

        ''' <summary>変更点の判定から除外したいプリミティブ値プロパティに付与した属性[]</summary>
        Public Property IgnorePropertyAttributes As Type()

        Private _IncludePropertyAttributes As Type()
        ''' <summary>変更点の判定に含めたいインスタンス値プロパティに付与した属性[]</summary>
        Public Property IncludePropertyAttributes As Type()
            Get
                Return _IncludePropertyAttributes
            End Get
            Set(value As Type())
                _IncludePropertyAttributes = value
                If CollectionUtil.IsEmpty(BeforeVos) Then
                    Return
                End If
                SupersedeBeforeUpdatedVos(BeforeVos)
            End Set
        End Property

        ''' <summary>
        ''' 変更前データ
        ''' </summary>
        Public ReadOnly Property BeforeVos() As List(Of T)
            Get
                Return _beforeVos
            End Get
        End Property
#End Region

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me.New(DirectCast(Nothing, IEnumerable(Of T)))
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="beforeVos">変更前データ</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal ParamArray beforeVos As T())
            Me.New(DirectCast(beforeVos, IEnumerable(Of T)))
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="beforeVos">変更前データ</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal beforeVos As IEnumerable(Of T))
            Me.New(beforeVos, New DefaultPlugin)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="aPlugin">変更前データ</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal aPlugin As Plugin)
            Me.New(Nothing, aPlugin)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="beforeVos">変更前データ</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal beforeVos As IEnumerable(Of T), ByVal aPlugin As Plugin)
            EzUtil.AssertParameterIsNotNull(aPlugin, "aPlugin")
            Me.aPlugin = New WrappedDefaultPluginImpl(aPlugin)
            Me.IsTypeImmutable = TypeUtil.IsTypeImmutable(GetType(T)) _
                    OrElse (GetType(T).HasElementType _
                                AndAlso (TypeUtil.IsTypeImmutable(GetType(T).GetElementType) _
                                         OrElse (GetType(Object) Is GetType(T).GetElementType)))
            SupersedeBeforeUpdatedVos(If(beforeVos Is Nothing, NewArrayInstance(0), beforeVos))
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="makeKey">同レコード判定に使うキーを作成するデリゲート</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal makeKey As MakeKeyCallback)
            Me.New(Nothing, makeKey)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="beforeVos">変更前データ</param>
        ''' <param name="makeKey">同レコード判定に使うキーを作成するデリゲート</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal beforeVos As IEnumerable(Of T), ByVal makeKey As MakeKeyCallback)
            Me.New(beforeVos, New DefaultDlgtPlugin(makeKey))
        End Sub

        ''' <summary>
        ''' 配列インスタンスを生成する
        ''' </summary>
        ''' <param name="length">配列の長さ</param>
        ''' <returns>新しい配列インスタンス</returns>
        ''' <remarks></remarks>
        Private Shared Function NewArrayInstance(ByVal length As Integer) As T()
            Dim result(length - 1) As T
            Return result
        End Function

        ''' <summary>
        ''' 変更前データを設定する
        ''' </summary>
        ''' <param name="beforeVos">変更前データ</param>
        ''' <remarks></remarks>
        Public Sub SupersedeBeforeUpdatedVos(ByVal ParamArray beforeVos As T())
            SupersedeBeforeUpdatedVos(DirectCast(beforeVos, IEnumerable(Of T)))
        End Sub

        ''' <summary>
        ''' 変更前データを設定する
        ''' </summary>
        ''' <param name="beforeVos">変更前データ</param>
        ''' <remarks></remarks>
        Public Sub SupersedeBeforeUpdatedVos(ByVal beforeVos As IEnumerable(Of T))
            cacheRwLock.EnterWriteLock()
            Try
                Me._beforeKeys = New List(Of Object)
                Me._beforeVos = New List(Of T)
                Me._beforeVoByKey = New Dictionary(Of Object, T)
                Me._beforeInstanceByKey = New Dictionary(Of Object, T)

                If 0 < beforeVos.Count AndAlso Not IsTypeImmutable Then
                    AssertCanNewInstance()
                End If
                InternalAddBeforeUpdatedVos(beforeVos)
            Finally
                cacheRwLock.ExitWriteLock()
            End Try
        End Sub

        Private Sub AssertCanNewInstance()
            Try
                NewInstanceFrom(Nothing)
            Catch ex As Exception
                Throw New InvalidOperationException("自作した不変オブジェクトを使用する場合、以下の順序が必要" & vbCrLf & "(1)コンストラクタにbeforeVosを指定しない (2)インスタンス生成直後に #IsTypeImmutable を True にする (3)SupersedeBeforeUpdateVos()で beforeVos を指定", ex)
            End Try
        End Sub

        ''' <summary>
        ''' 変更前データを追加する
        ''' </summary>
        ''' <param name="beforeVos">変更前データ</param>
        ''' <remarks></remarks>
        Public Sub AddBeforeUpdatedVos(ByVal ParamArray beforeVos As T())
            AddBeforeUpdatedVos(DirectCast(beforeVos, IEnumerable(Of T)))
        End Sub

        ''' <summary>
        ''' 変更前データを追加する
        ''' </summary>
        ''' <param name="beforeVos">変更前データ</param>
        ''' <remarks></remarks>
        Public Sub AddBeforeUpdatedVos(ByVal beforeVos As IEnumerable(Of T))
            cacheRwLock.EnterWriteLock()
            Try
                InternalAddBeforeUpdatedVos(beforeVos)
            Finally
                cacheRwLock.ExitWriteLock()
            End Try
        End Sub

        ''' <summary>
        ''' 「変更前データ」を追加する（同一キーが存在したら上書きする）
        ''' </summary>
        ''' <param name="beforeVos"></param>
        ''' <remarks></remarks>
        Private Sub InternalAddBeforeUpdatedVos(ByVal beforeVos As IEnumerable(Of T))
            If 0 < _beforeKeys.Count Then
                For Each obj As T In beforeVos
                    Dim key As Object = aPlugin.MakeKey(obj)
                    If Not _beforeKeys.Contains(key) Then
                        Continue For
                    End If
                    Me._beforeKeys.Remove(key)
                    Me._beforeVos.Remove(Me._beforeVoByKey(key))
                    Me._beforeVoByKey.Remove(key)
                    Me._beforeInstanceByKey.Remove(key)
                    If 0 = _beforeKeys.Count Then
                        Exit For
                    End If
                Next
            End If
            For Each obj As T In beforeVos
                Dim newInstance As T = NewInstanceFrom(obj)
                Dim key As Object = aPlugin.MakeKey(obj)
                Me._beforeKeys.Add(key)
                Me._beforeVos.Add(newInstance)
                Me._beforeVoByKey.Add(key, newInstance)
                Me._beforeInstanceByKey.Add(key, obj)
            Next
        End Sub

        ''' <summary>
        ''' 変更後データを設定する
        ''' </summary>
        ''' <param name="updatedVos">変更後データ</param>
        ''' <remarks>変更前データと同じインスタンスを、変更後データとして指定すること</remarks>
        Public Sub SetUpdatedVos(ByVal ParamArray updatedVos As T())
            SetUpdatedVos(DirectCast(updatedVos, IEnumerable(Of T)))
        End Sub

        ''' <summary>
        ''' 変更後データを設定する
        ''' </summary>
        ''' <param name="updatedVos">変更後データ</param>
        ''' <remarks>変更前データと同じインスタンスを、変更後データとして指定すること</remarks>
        Public Sub SetUpdatedVos(ByVal updatedVos As IEnumerable(Of T))
            cacheRwLock.EnterWriteLock()
            Try
                Me._updatedKeys = New List(Of Object)
                Me._updatedVos = New List(Of T)
                Me._updatedVoByKey = New Dictionary(Of Object, T)
                Me._updatedInstanceByKey = New Dictionary(Of Object, T)
                For Each updVo As T In updatedVos
                    Dim newInstance As T = NewInstanceFrom(updVo)
                    Dim key As Object = aPlugin.MakeKey(updVo)
                    Me._updatedVos.Add(newInstance)
                    Me._updatedKeys.Add(key)
#If DEBUG Then
                    If Me._updatedVoByKey.ContainsKey(key) Then
                        Throw New ArgumentException(String.Format("ChangeJudgeer(of {0})のキー '{1}' が重複している", GetType(T).Name, key))
                    End If
#End If
                    Me._updatedVoByKey.Add(key, newInstance)
                    Me._updatedInstanceByKey.Add(key, updVo)
                Next

            Finally
                cacheRwLock.ExitWriteLock()
            End Try
        End Sub

        Private Function NewInstanceFrom(ByVal updVo As T) As T
            If GetType(T).IsArray Then
                Dim updArray As Array = DirectCast(DirectCast(updVo, Object), Array)
                Dim elementType As Type = GetType(T).GetElementType
                Dim results As Array = Array.CreateInstance(elementType, If(updArray Is Nothing, 0, updArray.Length))
                If updArray IsNot Nothing Then
                    For i As Integer = 0 To updArray.Length - 1
                        results.SetValue(If(IsTypeImmutable, updArray.GetValue(i), PerformNewInstance(elementType, updArray.GetValue(i))), i)
                    Next
                End If
                Return DirectCast(DirectCast(results, Object), T)
            ElseIf IsTypeImmutable Then
                Return updVo
            End If
            Return PerformNewInstance(Of T)(updVo)
        End Function

        ''' <summary>
        ''' （指定されていればオブジェクト値の差分判定できる）インスタンスを生成する
        ''' </summary>
        ''' <typeparam name="T1">値の型</typeparam>
        ''' <param name="value">元になる値</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function PerformNewInstance(Of T1)(value As T1) As T1
            Return DirectCast(PerformNewInstance(GetType(T1), value), T1)
        End Function
        ''' <summary>
        ''' （指定されていればオブジェクト値の差分判定できる）インスタンスを生成する
        ''' </summary>
        ''' <param name="aType">値の型</param>
        ''' <param name="value">元になる値</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function PerformNewInstance(aType As Type, value As Object) As Object
            Dim newInstance As Object = VoUtil.NewInstance(aType, value)
            If value Is Nothing OrElse CollectionUtil.IsEmpty(IncludePropertyAttributes) Then
                Return newInstance
            End If
            For Each info As PropertyInfo In aType.GetProperties.Where(Function(p)
                                                                           Dim customAttributes As Object() = p.GetCustomAttributes(inherit:=False)
                                                                           Return p.GetGetMethod IsNot Nothing AndAlso Not TypeUtil.IsTypeImmutable(p.PropertyType) _
                                                                                  AndAlso customAttributes IsNot Nothing _
                                                                                  AndAlso customAttributes.Any(Function(attr) IncludePropertyAttributes.Contains(attr.GetType)) 
                                                                       End Function)
                Dim srcEntity As Object = info.GetGetMethod.Invoke(value, Nothing)
                Dim destEntity As Object = info.GetGetMethod.Invoke(newInstance, Nothing)
                If srcEntity Is destEntity Then
                    If info.GetSetMethod Is Nothing Then
                        Continue For
                    End If
                    destEntity = PerformNewInstance(info.PropertyType, srcEntity)
                    info.GetSetMethod.Invoke(newInstance, New Object() {destEntity})
                End If
                VoUtil.CopyProperties(srcEntity, destEntity)
            Next
            Return newInstance
        End Function

        ''' <summary>
        ''' 変更後データが、変更後データでなくなった（保存処理完了など）ことを通知する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub NotifySaveEnd()
            AssertUpdatedVoIsNotNull()
            SupersedeBeforeUpdatedVos(From updatedKey As Object In _updatedKeys Select _updatedInstanceByKey(updatedKey))

            cacheRwLock.EnterWriteLock()
            Try
                Dim beforeKeyByInstance As Dictionary(Of T, Object) = _beforeInstanceByKey.ToDictionary(Function(pair) pair.Value, Function(pair) pair.Key)
                For Each pair As KeyValuePair(Of Object, T) In _updatedInstanceByKey
                    Dim updatedKey As Object = pair.Key
                    Dim originInstance As T = pair.Value
                    Dim updatedVo As T = _updatedVoByKey(updatedKey)
                    Dim beforeKey As Object = beforeKeyByInstance(originInstance)
                    Dim beforeVo As T = _beforeVoByKey(beforeKey)
                    VoUtil.CopyProperties(updatedVo, beforeVo)
                Next

                _updatedKeys = Nothing
                _updatedVos = Nothing
                _updatedVoByKey = Nothing
                _updatedInstanceByKey = Nothing

            Finally
                cacheRwLock.ExitWriteLock()
            End Try
        End Sub

        ''' <summary>
        ''' #ExtractXxxxで抽出した変更点Voから、Voの元々のインスタンスを返す（第三者によってインスタンス内の値は変更されている可能性がある）
        ''' </summary>
        ''' <param name="extractsVo">変更点Vo</param>
        ''' <returns>Voの元々のインスタンス</returns>
        ''' <remarks></remarks>
        Public Function DetectOriginInstance(ByVal extractsVo As T) As T
            cacheRwLock.EnterReadLock()
            Try
                Dim result As T
                result = PerformDetectOriginInstance(extractsVo, _updatedVoByKey, _updatedInstanceByKey)
                If result IsNot Nothing Then
                    Return result
                End If
                Return PerformDetectOriginInstance(extractsVo, _beforeVoByKey, _beforeInstanceByKey)
            Finally
                cacheRwLock.ExitReadLock()
            End Try
        End Function

        Private Function PerformDetectOriginInstance(ByVal extractsVo As T, ByVal updatedVoByKey As Dictionary _
                                                  (Of Object, T), ByVal updatedInstanceByKey As Dictionary(Of Object, T)) As T

            If updatedVoByKey Is Nothing Then
                Return Nothing
            End If

            Dim extractsKey As Object = aPlugin.MakeKey(extractsVo)
            If updatedInstanceByKey.ContainsKey(extractsKey) Then
                Return updatedInstanceByKey(extractsKey)
            End If

            If updatedVoByKey.ContainsValue(extractsVo) Then
                Dim extractsObj As Object = extractsVo
                For Each pair As KeyValuePair(Of Object, T) In updatedVoByKey
                    Dim obj As Object = pair.Value
                    If obj Is extractsObj Then
                        Return updatedInstanceByKey(pair.Key)
                    End If
                Next
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' 変更前データと変更後データとに差異があるかを返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function HasChanged() As Boolean
            AssertUpdatedVoIsNotNull()

            cacheRwLock.EnterReadLock()
            Try
                If _updatedKeys.Count <> _beforeKeys.Count Then
                    Return True
                End If

                For Each pair As KeyValuePair(Of Object, T) In _updatedVoByKey
                    If Not _beforeVoByKey.ContainsKey(pair.Key) Then
                        Return True
                    End If
                    If Not VoUtil.IsEquals(pair.Value, _beforeVoByKey(pair.Key), IgnorePropertyAttributes) Then
                        Return True
                    End If
                Next

                Return False

            Finally
                cacheRwLock.ExitReadLock()
            End Try
        End Function

        ''' <summary>
        ''' 削除データがあるかを返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function WasDeleted() As Boolean
            AssertUpdatedVoIsNotNull()
            cacheRwLock.EnterReadLock()
            Try
                For Each key As Object In _beforeVoByKey.Keys
                    If Not _updatedVoByKey.ContainsKey(key) Then
                        Return True
                    End If
                Next
                Return False

            Finally
                cacheRwLock.ExitReadLock()
            End Try
        End Function

        ''' <summary>
        ''' 削除データを抽出する
        ''' </summary>
        ''' <returns>削除されたデータ</returns>
        ''' <remarks></remarks>
        Public Function ExtractDeletedVos() As T()
            AssertUpdatedVoIsNotNull()
            cacheRwLock.EnterReadLock()
            Try
                Dim result As New List(Of T)
                For Each key As Object In _beforeKeys
                    If Not _updatedVoByKey.ContainsKey(key) Then
                        result.Add(_beforeVoByKey(key))
                    End If
                Next
                Return result.ToArray

            Finally
                cacheRwLock.ExitReadLock()
            End Try
        End Function

        ''' <summary>
        ''' 追加データがあるかを返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function WasInserted() As Boolean
            AssertUpdatedVoIsNotNull()
            cacheRwLock.EnterReadLock()
            Try
                For Each key As Object In _updatedVoByKey.Keys
                    If Not _beforeVoByKey.ContainsKey(key) Then
                        Return True
                    End If
                Next
                Return False

            Finally
                cacheRwLock.ExitReadLock()
            End Try
        End Function

        ''' <summary>
        ''' 追加データを抽出する
        ''' </summary>
        ''' <returns>追加したデータ</returns>
        ''' <remarks></remarks>
        Public Function ExtractInsertedVos() As T()
            AssertUpdatedVoIsNotNull()
            cacheRwLock.EnterReadLock()
            Try
                Dim result As New List(Of T)
                For Each key As Object In _updatedKeys
                    If Not _beforeVoByKey.ContainsKey(key) Then
                        result.Add(_updatedVoByKey(key))
                    End If
                Next
                Return result.ToArray

            Finally
                cacheRwLock.ExitReadLock()
            End Try
        End Function

        ''' <summary>
        ''' 変更データがあるかを返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function WasUpdated() As Boolean
            AssertUpdatedVoIsNotNull()
            cacheRwLock.EnterReadLock()
            Try
                For Each pair As KeyValuePair(Of Object, T) In _updatedVoByKey
                    If Not _beforeVoByKey.ContainsKey(pair.Key) Then
                        Continue For
                    End If
                    If Not VoUtil.IsEquals(pair.Value, _beforeVoByKey(pair.Key), IgnorePropertyAttributes) Then
                        Return True
                    End If
                Next
                Return False

            Finally
                cacheRwLock.ExitReadLock()
            End Try

        End Function

        ''' <summary>
        ''' #SetUpdatedVos()にて更新後データが指定されている事を保証する
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub AssertUpdatedVoIsNotNull()
            If _updatedKeys Is Nothing Then
                Throw New InvalidOperationException("#SetUpdatedVosを呼び出されていない. 変更データが設定されていない.")
            End If
        End Sub

        ''' <summary>
        ''' 変更データを抽出する
        ''' </summary>
        ''' <returns>変更したデータ</returns>
        ''' <remarks></remarks>
        Public Function ExtractUpdatedVos() As UpdateInfo()
            AssertUpdatedVoIsNotNull()
            cacheRwLock.EnterReadLock()
            Try
                Dim result As New List(Of UpdateInfo)
                For Each key As Object In _updatedKeys
                    If Not _beforeVoByKey.ContainsKey(key) Then
                        Continue For
                    End If
                    If Not VoUtil.IsEquals(_updatedVoByKey(key), _beforeVoByKey(key), IgnorePropertyAttributes) Then
                        result.Add(New UpdateInfo(_beforeVoByKey(key), _updatedVoByKey(key)))
                    End If
                Next
                Return result.ToArray

            Finally
                cacheRwLock.ExitReadLock()
            End Try
        End Function

    End Class
End Namespace