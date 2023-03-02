Imports System.Threading

Namespace Util
    ''' <summary>
    ''' LRU(Least Recently Used) CacheのFhi.Fw実装
    ''' </summary>
    ''' <remarks>
    ''' LruCacheはDictionaryに適さない理由
    ''' 
    ''' Dictionaryでいいなら以下のように書きたい!!
    ''' ContainsKey()で存在チェックして、無ければ取得する方法で
    ''' ```vb:Dictionary版
    ''' Private hoge As New LruCacheDictionary(Of String, String)()
    ''' Public Function GetFuga(key As String) As String
    '''     If Not hoge.ContainsKey(key) Then  ' (1)
    '''         hoge.Add(key, poo.FindByPk(key)) ' `poo`は外部参照する「なにか」
    '''     End If
    '''     Return hoge(key)                   ' (2)
    ''' End Function
    ''' ```
    ''' 
    ''' しかし、(1)でContainsKey()がtrueを返したとしても、(2)の時点で値が返される保証はない
    ''' なぜなら、(1)の処理が終わった瞬間に、hogeからkeyが削除される場合があるため
    ''' 「ロックする」手法も無くはないが、それなら素直にLruCacheを使ったほうがシンプル(以下が例)
    ''' ```vb:LruCache版
    ''' Private hoge As New LruCache(Of String, String)(New TimeSpan(0, 0, seconds:=60))
    ''' Public Function GetFuga(key As String) As String
    '''     Dim result As String = hoge.Get(key)
    '''     If result Is Nothing Then
    '''         result = poo.FindByPk(key)
    '''         hoge.Add(key, result)
    '''     End If
    '''     Return result
    ''' End Function
    ''' ```
    ''' </remarks>
    Public Class LruCache(Of TKey, TValue)

#Region "CacheNode ..."
        Protected Class CacheNode
            Private _Key As TKey
            Private _Value As TValue
            Private _AccessCount As Integer
            Private _LinkedListNode As LinkedListNode(Of CacheNode)
            Private ReadOnly timeoutDate As DateTime
            Private ReadOnly behavior As IBehavior

            ''' <summary>key値</summary>
            Public Property Key() As TKey
                Get
                    Return _Key
                End Get
                Set(ByVal value As TKey)
                    _Key = value
                End Set
            End Property

            ''' <summary>value値</summary>
            Public Property Value() As TValue
                Get
                    Return _Value
                End Get
                Set(ByVal value As TValue)
                    _Value = value
                End Set
            End Property

            ''' <summary>参照回数</summary>
            Public Property AccessCount() As Integer
                Get
                    Return _AccessCount
                End Get
                Set(ByVal value As Integer)
                    _AccessCount = value
                End Set
            End Property

            ''' <summary>LinkedListノード</summary>
            Public Property LinkedListNode() As LinkedListNode(Of CacheNode)
                Get
                    Return _LinkedListNode
                End Get
                Set(ByVal value As LinkedListNode(Of CacheNode))
                    _LinkedListNode = value
                End Set
            End Property

            Public Sub New(ByVal key As TKey, ByVal value As TValue, ByVal timeoutDate As DateTime, behavior As IBehavior)
                Me.Key = key
                Me.Value = value
                Me.timeoutDate = timeoutDate
                Me.behavior = behavior
            End Sub

            ''' <summary>
            ''' 期限切れかどうか
            ''' </summary>
            ''' <returns>期限切れならTrue、それ以外False</returns>
            ''' <remarks></remarks>
            Public Function IsExpired() As Boolean
                Return timeoutDate <= behavior.GetNow
            End Function
        End Class
#End Region
#Region "Nested class"
        Public Interface IBehavior
            ''' <summary>
            ''' 時間を取得する
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function GetNow() As DateTime
        End Interface
        Private Class DefaultBehavior : Implements IBehavior
            Public Function GetNow() As DateTime Implements IBehavior.GetNow
                Return DateUtil.GetNowDateTime()
            End Function
        End Class
#End Region

        Protected ReadOnly cached As New Dictionary(Of TKey, CacheNode)
        Protected ReadOnly lruLinkedList As New LinkedList(Of CacheNode)
        Private ReadOnly maxSize As Integer
        Private ReadOnly timeout As TimeSpan
        Protected ReadOnly lock As New ReaderWriterLockSlim
        Private ReadOnly behavior As IBehavior

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="expiryTimeoutMillis">生存期間(ms)</param>
        ''' <param name="maxCacheSize">LruCaheに保持できるサイズ</param>
        ''' <param name="memoryRefreshInterval">(内部の)お掃除スレッドの実行間隔</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal expiryTimeoutMillis As Integer, Optional ByVal maxCacheSize As Integer = Integer.MaxValue, Optional ByVal memoryRefreshInterval As Integer = 1000, Optional behavior As IBehavior = Nothing)
            Me.New(New TimeSpan(0, 0, 0, 0, milliseconds:=expiryTimeoutMillis), maxCacheSize, memoryRefreshInterval, behavior)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="expiryTimeout">生存期間</param>
        ''' <param name="maxCacheSize">LruCaheに保持できるサイズ</param>
        ''' <param name="memoryRefreshInterval">(内部の)お掃除スレッドの実行間隔</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal expiryTimeout As TimeSpan, Optional ByVal maxCacheSize As Integer = Integer.MaxValue, Optional ByVal memoryRefreshInterval As Integer = 1000, Optional behavior As IBehavior = Nothing)
            Me.timeout = expiryTimeout
            Me.maxSize = maxCacheSize
            Me.behavior = If(behavior, New DefaultBehavior)
        End Sub

        Private Sub LogDebug(message As String, ParamArray args As Object())
            ' テスト時に使いたいのでコメントアウト
            'EzUtil.logDebug(message, args)
        End Sub

        ''' <summary>
        ''' key値に紐付くvalue値をLRUキャッシュに追加する
        ''' </summary>
        ''' <param name="key">key値</param>
        ''' <param name="value">value値</param>
        ''' <remarks></remarks>
        Public Sub Add(ByVal key As TKey, ByVal value As TValue)
            AssertExistsKey(key)

            lock.EnterWriteLock()
            Try
                LogDebug("key: {0} をLRUキャッシュに追加します", key)
                Dim node As CacheNode = Nothing
                If cached.TryGetValue(key, node) Then
                    Delete(node)
                End If

                ShrinkToSize(maxSize - 1)
                CreateNodeAndAddToList(key, value)
            Finally
                lock.ExitWriteLock()
            End Try
        End Sub

        Private Sub AssertExistsKey(ByVal key As TKey)
            If key Is Nothing Then
                Throw New ArgumentException("key値がありません")
            End If
        End Sub

        Private Sub ShrinkToSize(ByVal desiredSize As Integer)
            While desiredSize < cached.Count
                RemoveLeastRecentlyUsedNode()
            End While
        End Sub

        Private Sub RemoveLeastRecentlyUsedNode()
            If lruLinkedList.Last.Value IsNot Nothing Then
                Dim node As CacheNode = lruLinkedList.Last.Value
                Delete(node)
            End If
        End Sub

        Private Sub Delete(ByVal node As CacheNode)
            LogDebug("key: {0} をLRUキャッシュから削除します", node.Key.ToString())
            lruLinkedList.Remove(node.LinkedListNode)
            cached.Remove(node.Key)
        End Sub

        Private Sub CreateNodeAndAddToList(ByVal key As TKey, ByVal value As TValue)
            Dim timeoutDate As DateTime = If(DateTime.MaxValue.Subtract(behavior.GetNow) < timeout, DateTime.MaxValue, behavior.GetNow.Add(timeout))
            Dim node As New CacheNode(key, value, timeoutDate, behavior)
            node.LinkedListNode = lruLinkedList.AddFirst(node)
            cached(key) = node
        End Sub

        ''' <summary>
        ''' key値に紐付くvalue値を取得する
        ''' </summary>
        ''' <param name="key">key値</param>
        ''' <returns>value値</returns>
        ''' <remarks></remarks>
        Public Function [Get](ByVal key As TKey) As TValue
            AssertExistsKey(key)
            Return GetValueSafely(key)
        End Function

        Private Function GetValueSafely(ByVal key As TKey) As TValue
            Dim value As TValue
            Return If(PerformTryGetValue(key, value), value, Nothing)
        End Function

        Protected Function PerformTryGetValue(ByVal key As TKey, ByRef value As TValue) As Boolean
            Dim node As CacheNode = Nothing
            lock.EnterWriteLock()
            Try
                If cached.TryGetValue(key, node) Then
                    If node IsNot Nothing AndAlso Not node.IsExpired Then
                        LogDebug("key: {0} のLRUキャッシュが見つかりました", key)
                        node.AccessCount += 1
                        value = node.Value

                        lruLinkedList.Remove(node.LinkedListNode)
                        node.LinkedListNode = InsertByAccessCount(node)
                        Return value IsNot Nothing
                    ElseIf node IsNot Nothing AndAlso node.IsExpired Then
                        Delete(node)
                    End If
                End If
                LogDebug("key: {0} のLRUキャッシュはありませんでした", key)
                Return False
            Finally
                lock.ExitWriteLock()
            End Try
        End Function

        Private Function InsertByAccessCount(ByVal node As CacheNode) As LinkedListNode(Of CacheNode)
            For i As Integer = lruLinkedList.Count - 1 To 0 Step -1
                Dim cacheNode As CacheNode = lruLinkedList(i)

                If cacheNode.AccessCount = node.AccessCount Then
                    Return lruLinkedList.AddAfter(cacheNode.LinkedListNode, node)
                Else
                    Continue For
                End If
            Next

            Return lruLinkedList.AddFirst(node)
        End Function

    End Class
End Namespace