Imports System.Threading

Namespace Util
    ''' <summary>
    ''' 同期化を施した IndexedList
    ''' （.NET Framework 3.5 を使用できる環境なら ReaderWriterLock の代わりに ReaderWriterLockSlim
    ''' を使用した Synchronizedクラスを用意すべき. ReaderWriterLockSlim のほうが早い）
    ''' </summary>
    ''' <typeparam name="T">要素の型</typeparam>
    ''' <remarks></remarks>
    Public Class SynchronizedIndexedList(Of T) : Implements IIndexedList(Of T)

        Private rwLock As New ReaderWriterLock
        Private ReadOnly aList As IndexedList(Of T)

        Public Sub New()
            aList = New IndexedList(Of T)
        End Sub

        Public Sub New(ByVal values As ICollection(Of T))
            aList = New IndexedList(Of T)(values)
        End Sub

        Public Sub New(ByVal IsCreateGeneric As Boolean)
            aList = New IndexedList(Of T)(IsCreateGeneric)
        End Sub

        Public Sub New(ByVal IsCreateGeneric As Boolean, ByVal values As ICollection(Of T))
            aList = New IndexedList(Of T)(IsCreateGeneric, values)
        End Sub

        Public Sub New(ByVal IsCreateGeneric As Boolean, ByVal instanceType As Type)
            aList = New IndexedList(Of T)(IsCreateGeneric, instanceType)
        End Sub

        Public Sub New(ByVal IsCreateGeneric As Boolean, ByVal values As ICollection(Of T), ByVal instanceType As Type)
            aList = New IndexedList(Of T)(IsCreateGeneric, values, instanceType)
        End Sub

        Public Function HasValue(ByVal index As Integer) As Boolean Implements IIndexedList(Of T).HasValue
            rwLock.AcquireReaderLock(Timeout.Infinite)
            Try
                Return aList.HasValue(index)
            Finally
                rwLock.ReleaseReaderLock()
            End Try
        End Function

        Default Public Property Item(ByVal index As Integer) As T Implements IIndexedList(Of T).Item
            Get
                rwLock.AcquireReaderLock(Timeout.Infinite)
                Try
                    Return aList.Item(index)
                Finally
                    rwLock.ReleaseReaderLock()
                End Try
            End Get
            Set(ByVal value As T)
                rwLock.AcquireWriterLock(Timeout.Infinite)
                Try
                    aList.Item(index) = value
                Finally
                    rwLock.ReleaseWriterLock()
                End Try
            End Set
        End Property

        Public ReadOnly Property Items() As ICollection(Of T) Implements IIndexedList(Of T).Items
            Get
                rwLock.AcquireReaderLock(Timeout.Infinite)
                Try
                    Return aList.Items
                Finally
                    rwLock.ReleaseReaderLock()
                End Try
            End Get
        End Property

        Public ReadOnly Property Indexes() As ICollection(Of Integer) Implements IIndexedList(Of T).Indexes
            Get
                rwLock.AcquireReaderLock(Timeout.Infinite)
                Try
                    Return aList.Indexes
                Finally
                    rwLock.ReleaseReaderLock()
                End Try
            End Get
        End Property

        Public Function GetMaxIndex() As Integer? Implements IIndexedList(Of T).GetMaxIndex
            rwLock.AcquireReaderLock(Timeout.Infinite)
            Try
                Return aList.GetMaxIndex
            Finally
                rwLock.ReleaseReaderLock()
            End Try
        End Function

        Public Sub Add(ByVal value As T) Implements IIndexedList(Of T).Add
            rwLock.AcquireWriterLock(Timeout.Infinite)
            Try
                aList.Add(value)
            Finally
                rwLock.ReleaseWriterLock()
            End Try
        End Sub

        Public Sub AddRange(ByVal values As ICollection(Of T)) Implements IIndexedList(Of T).AddRange
            rwLock.AcquireWriterLock(Timeout.Infinite)
            Try
                aList.AddRange(values)
            Finally
                rwLock.ReleaseWriterLock()
            End Try
        End Sub

        Public Sub Clear() Implements IIndexedList(Of T).Clear
            rwLock.AcquireWriterLock(Timeout.Infinite)
            Try
                aList.Clear()
            Finally
                rwLock.ReleaseWriterLock()
            End Try
        End Sub

        Public Sub ClearAt(ByVal index As Integer) Implements IIndexedList(Of T).ClearAt
            rwLock.AcquireWriterLock(Timeout.Infinite)
            Try
                aList.ClearAt(index)
            Finally
                rwLock.ReleaseWriterLock()
            End Try
        End Sub

        Public Sub InsertAt(ByVal index As Integer, Optional ByVal count As Integer = 1) Implements IIndexedList(Of T).InsertAt
            rwLock.AcquireWriterLock(Timeout.Infinite)
            Try
                aList.InsertAt(index, count)
            Finally
                rwLock.ReleaseWriterLock()
            End Try
        End Sub

        Public Sub RemoveAt(ByVal index As Integer, Optional ByVal count As Integer = 1) Implements IIndexedList(Of T).RemoveAt
            rwLock.AcquireWriterLock(Timeout.Infinite)
            Try
                aList.RemoveAt(index, count)
            Finally
                rwLock.ReleaseWriterLock()
            End Try
        End Sub

        Public Function Extract(ByVal rowIndex As Integer, ByVal count As Integer) As T() Implements IIndexedList(Of T).Extract
            rwLock.AcquireReaderLock(Timeout.Infinite)
            Try
                Return aList.Extract(rowIndex, count)
            Finally
                rwLock.ReleaseReaderLock()
            End Try
        End Function

        Public Sub Supersede(ByVal rowIndex As Integer, ByVal recordBags As T()) Implements IIndexedList(Of T).Supersede
            rwLock.AcquireWriterLock(Timeout.Infinite)
            Try
                aList.Supersede(rowIndex, recordBags)
            Finally
                rwLock.ReleaseWriterLock()
            End Try
        End Sub

        Public Function IndexOf(ByVal value As T) As Integer Implements IIndexedList(Of T).IndexOf
            rwLock.AcquireReaderLock(Timeout.Infinite)
            Try
                Return aList.IndexOf(value)
            Finally
                rwLock.ReleaseReaderLock()
            End Try
        End Function

        Public Sub SuppressGap() Implements IIndexedList(Of T).SuppressGap
            rwLock.AcquireWriterLock(Timeout.Infinite)
            Try
                aList.SuppressGap()
            Finally
                rwLock.ReleaseWriterLock()
            End Try
        End Sub

        Public Sub TrimEnd() Implements IIndexedList(Of T).TrimEnd
            rwLock.AcquireWriterLock(Timeout.Infinite)
            Try
                aList.TrimEnd()
            Finally
                rwLock.ReleaseWriterLock()
            End Try
        End Sub

        Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
            rwLock.AcquireReaderLock(Timeout.Infinite)
            Try
                Return aList.GetEnumerator
            Finally
                rwLock.ReleaseReaderLock()
            End Try
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            rwLock.AcquireReaderLock(Timeout.Infinite)
            Try
                Return aList.GetEnumerator
            Finally
                rwLock.ReleaseReaderLock()
            End Try
        End Function
    End Class
End Namespace