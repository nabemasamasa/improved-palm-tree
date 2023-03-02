Imports Fhi.Fw.Lang.Threading

Namespace Util
    ''' <summary>
    ''' タイムアウトで内容を破棄（クリア）するDictionary ※スレッドセーフではない
    ''' </summary>
    ''' <typeparam name="TKey"></typeparam>
    ''' <typeparam name="TValue"></typeparam>
    ''' <remarks></remarks>
    Public Class TimeoutDictionary(Of TKey, TValue) : Implements IDictionary(Of TKey, TValue)

        Private ReadOnly origin As Dictionary(Of TKey, TValue)

        Public Sub New()
            origin = New Dictionary(Of TKey, TValue)
        End Sub

        Public Sub New(ByVal capacity As Integer)
            origin = New Dictionary(Of TKey, TValue)(capacity)
        End Sub

        Public Sub New(ByVal comparer As IEqualityComparer(Of TKey))
            origin = New Dictionary(Of TKey, TValue)(comparer)
        End Sub

        Public Sub New(ByVal capacity As Integer, ByVal comparer As IEqualityComparer(Of TKey))
            origin = New Dictionary(Of TKey, TValue)(capacity, comparer)
        End Sub

        Public Sub New(ByVal dictionary As IDictionary(Of TKey, TValue))
            origin = New Dictionary(Of TKey, TValue)(dictionary)
        End Sub

        Public Sub New(ByVal dictionary As IDictionary(Of TKey, TValue), ByVal comparer As IEqualityComparer(Of TKey))
            origin = New Dictionary(Of TKey, TValue)(dictionary, comparer)
        End Sub

        Public Function ContainsKey(ByVal key As TKey) As Boolean Implements IDictionary(Of TKey, TValue).ContainsKey
            Return origin.ContainsKey(key)
        End Function

        Public Function Remove(ByVal key As TKey) As Boolean Implements IDictionary(Of TKey, TValue).Remove
            Return origin.Remove(key)
        End Function

        Public Function TryGetValue(ByVal key As TKey, ByRef value As TValue) As Boolean Implements IDictionary(Of TKey, TValue).TryGetValue
            Return origin.TryGetValue(key, value)
        End Function

        Default Public Property Item(ByVal key As TKey) As TValue Implements IDictionary(Of TKey, TValue).Item
            Get
                Return origin.Item(key)
            End Get
            Set(ByVal value As TValue)
                origin.Item(key) = value
            End Set
        End Property

        Public ReadOnly Property Keys() As ICollection(Of TKey) Implements IDictionary(Of TKey, TValue).Keys
            Get
                Return origin.Keys
            End Get
        End Property

        Public ReadOnly Property Values() As ICollection(Of TValue) Implements IDictionary(Of TKey, TValue).Values
            Get
                Return origin.Values
            End Get
        End Property

        Public Sub Add(ByVal item As KeyValuePair(Of TKey, TValue)) Implements ICollection(Of KeyValuePair(Of TKey, TValue)).Add
            origin.Add(item.Key, item.Value)
        End Sub

        Public Function Contains(ByVal item As KeyValuePair(Of TKey, TValue)) As Boolean Implements ICollection(Of KeyValuePair(Of TKey, TValue)).Contains
            Return origin.Contains(item)
        End Function

        Public Sub CopyTo(ByVal array As KeyValuePair(Of TKey, TValue)(), ByVal arrayIndex As Integer) Implements ICollection(Of KeyValuePair(Of TKey, TValue)).CopyTo
            DirectCast(origin, ICollection(Of KeyValuePair(Of TKey, TValue))).CopyTo(array, arrayIndex)
        End Sub

        Public Function Remove(ByVal item As KeyValuePair(Of TKey, TValue)) As Boolean Implements ICollection(Of KeyValuePair(Of TKey, TValue)).Remove
            Return origin.Remove(item.Key)
        End Function

        Public ReadOnly Property Count() As Integer Implements ICollection(Of KeyValuePair(Of TKey, TValue)).Count
            Get
                Return origin.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly() As Boolean Implements ICollection(Of KeyValuePair(Of TKey, TValue)).IsReadOnly
            Get
                Return DirectCast(origin, ICollection(Of KeyValuePair(Of TKey, TValue))).IsReadOnly
            End Get
        End Property

        Public Function GetEnumerator() As IEnumerator(Of KeyValuePair(Of TKey, TValue)) Implements IEnumerable(Of KeyValuePair(Of TKey, TValue)).GetEnumerator
            Return origin.GetEnumerator
        End Function

        Public Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return origin.GetEnumerator
        End Function

        Public Sub Add(ByVal key As TKey, ByVal value As TValue) Implements IDictionary(Of TKey, TValue).Add
            origin.Add(key, value)
            If Not isTimeoutProcessing AndAlso 0 < timeoutMillis Then
                SetTimeout(timeoutMillis)
            End If
        End Sub

        Public Sub Clear() Implements ICollection(Of KeyValuePair(Of TKey, TValue)).Clear
            PerformClearThread()
            origin.Clear()
        End Sub

        Private isTimeoutProcessing As Boolean
        Private timeoutMillis As Long
        Private aClearThread As ClearThread

        ''' <summary>
        ''' クリアスレッド
        ''' </summary>
        ''' <remarks></remarks>
        Private Class ClearThread : Inherits AThread
            Private ReadOnly parent As TimeoutDictionary(Of TKey, TValue)
            Public Sub New(ByVal parent As TimeoutDictionary(Of TKey, TValue))
                Me.parent = parent
            End Sub
            Private _requestsShutdown As Boolean
            Public Sub Shutdown()
                _requestsShutdown = True
            End Sub
            Public Overrides Sub Run()
                Dim time As New Stopwatch
                time.Start()
                Dim sleepMillis As Integer = Math.Min(100, CInt(Math.Ceiling(parent.timeoutMillis / 30)))
                While time.ElapsedMilliseconds < parent.timeoutMillis
                    If _requestsShutdown Then
                        Return
                    End If
                    Threading.Thread.Sleep(sleepMillis)
                End While
                parent.Clear()
            End Sub
        End Class

        ''' <summary>
        ''' クリア時間をクリアする
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ClearTimeout()
            SetTimeout(0)
        End Sub

        ''' <summary>
        ''' クリア時間を取得する
        ''' </summary>
        ''' <returns>クリアするまでの時間(ms)</returns>
        ''' <remarks></remarks>
        Public Function GetTimeout() As Long
            Return timeoutMillis
        End Function

        ''' <summary>
        ''' クリアするまでの時間を設定する
        ''' </summary>
        ''' <param name="timeoutMillis">クリアするまでの時間(ms)</param>
        ''' <remarks></remarks>
        Public Sub SetTimeout(ByVal timeoutMillis As Long)
            Me.timeoutMillis = timeoutMillis
            PerformClearThread()
            If timeoutMillis <= 0 OrElse Not origin.Any Then
                Return
            End If
            aClearThread = New ClearThread(Me)
            aClearThread.Start()
            isTimeoutProcessing = True
        End Sub

        Private Sub PerformClearThread()
            If aClearThread Is Nothing OrElse Not aClearThread.IsAlive Then
                Return
            End If
            aClearThread.Shutdown()
            aClearThread = Nothing
            isTimeoutProcessing = False
        End Sub

    End Class
End Namespace