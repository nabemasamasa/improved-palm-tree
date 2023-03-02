Imports System.Runtime.InteropServices

Namespace Util
    ''' <summary>
    ''' 追加順を維持するDictionary
    ''' </summary>
    ''' <typeparam name="TKey"></typeparam>
    ''' <typeparam name="TValue"></typeparam>
    ''' <remarks></remarks>
    Public Class SequencedDictionary(Of TKey, TValue) : Implements IDictionary(Of TKey, TValue)

        Private ReadOnly dict As Dictionary(Of TKey, TValue)
        Private ReadOnly sequencedKeys As New List(Of TKey)

        Public Sub New()
            dict = New Dictionary(Of TKey, TValue)
        End Sub

        Public Sub New(capacity As Integer)
            dict = New Dictionary(Of TKey, TValue)(capacity)
        End Sub

        Public Sub New(dictionary As IDictionary(Of TKey, TValue))
            dict = New Dictionary(Of TKey, TValue)(dictionary)
        End Sub

        Public Sub New(comparer As IEqualityComparer(Of TKey))
            dict = New Dictionary(Of TKey, TValue)(Comparer)
        End Sub

        Public Sub New(capacity As Integer, comparer As IEqualityComparer(Of TKey))
            dict = New Dictionary(Of TKey, TValue)(capacity, comparer)
        End Sub

        Public Sub New(dictionary As IDictionary(Of TKey, TValue), comparer As IEqualityComparer(Of TKey))
            dict = New Dictionary(Of TKey, TValue)(dictionary, comparer)
        End Sub

        Public Function IEnumerable_GetEnumerator() As IEnumerator(Of KeyValuePair(Of TKey, TValue)) Implements IEnumerable(Of KeyValuePair(Of TKey, TValue)).GetEnumerator
            Return DirectCast(MakeSequencedList(), IEnumerable(Of KeyValuePair(Of TKey, TValue))).GetEnumerator
        End Function

        Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return DirectCast(MakeSequencedList(), IEnumerable).GetEnumerator
        End Function

        Private Function MakeSequencedList() As List(Of KeyValuePair(Of TKey, TValue))
            Dim pairByKey As Dictionary(Of TKey, KeyValuePair(Of TKey, TValue)) = dict.ToDictionary(Function(pair) pair.Key)
            Return sequencedKeys.Select(Function(key) pairByKey(key)).ToList
        End Function

        Public Sub Add(item As KeyValuePair(Of TKey, TValue)) Implements ICollection(Of KeyValuePair(Of TKey, TValue)).Add
            DirectCast(dict, ICollection(Of KeyValuePair(Of TKey, TValue))).Add(item)
            If sequencedKeys.Contains(item.Key) Then
                Throw New InvalidOperationException("二重key = " & item.Key.ToString)
            End If
            sequencedKeys.Add(item.Key)
        End Sub

        Public Sub Clear() Implements ICollection(Of KeyValuePair(Of TKey, TValue)).Clear
            dict.Clear()
            sequencedKeys.Clear()
        End Sub

        Public Function Contains(item As KeyValuePair(Of TKey, TValue)) As Boolean Implements ICollection(Of KeyValuePair(Of TKey, TValue)).Contains
            Return DirectCast(dict, ICollection(Of KeyValuePair(Of TKey, TValue))).Contains(item)
        End Function

        Public Sub CopyTo(array As KeyValuePair(Of TKey, TValue)(), arrayIndex As Integer) Implements ICollection(Of KeyValuePair(Of TKey, TValue)).CopyTo
            MakeSequencedList.CopyTo(array, arrayIndex)
        End Sub

        Public Function Remove(item As KeyValuePair(Of TKey, TValue)) As Boolean Implements ICollection(Of KeyValuePair(Of TKey, TValue)).Remove
            Return DirectCast(dict, ICollection(Of KeyValuePair(Of TKey, TValue))).Remove(item)
        End Function

        Public ReadOnly Property Count As Integer Implements ICollection(Of KeyValuePair(Of TKey, TValue)).Count
            Get
                Return DirectCast(dict, ICollection(Of KeyValuePair(Of TKey, TValue))).Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of KeyValuePair(Of TKey, TValue)).IsReadOnly
            Get
                Return DirectCast(dict, ICollection(Of KeyValuePair(Of TKey, TValue))).IsReadOnly
            End Get
        End Property

        Public Function ContainsKey(key As TKey) As Boolean Implements IDictionary(Of TKey, TValue).ContainsKey
            Return dict.ContainsKey(key)
        End Function

        Public Sub Add(key As TKey, value As TValue) Implements IDictionary(Of TKey, TValue).Add
            dict.Add(key, value)
            sequencedKeys.Add(key)
        End Sub

        Public Function Remove(key As TKey) As Boolean Implements IDictionary(Of TKey, TValue).Remove
            sequencedKeys.Remove(key)
            Return dict.Remove(key)
        End Function

        Public Function TryGetValue(key As TKey, <Out()> ByRef value As TValue) As Boolean Implements IDictionary(Of TKey, TValue).TryGetValue
            Return dict.TryGetValue(key, value)
        End Function

        Default Public Property Item(key As TKey) As TValue Implements IDictionary(Of TKey, TValue).Item
            Get
                Return dict.Item(key)
            End Get
            Set(value As TValue)
                dict.Item(key) = value
            End Set
        End Property

        Public ReadOnly Property Keys As ICollection(Of TKey) Implements IDictionary(Of TKey, TValue).Keys
            Get
                Return sequencedKeys
            End Get
        End Property

        Public ReadOnly Property Values As ICollection(Of TValue) Implements IDictionary(Of TKey, TValue).Values
            Get
                Return sequencedKeys.Select(Function(key) dict(key)).ToList
            End Get
        End Property

    End Class
End Namespace