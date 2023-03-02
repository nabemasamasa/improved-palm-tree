Imports System.Threading

Namespace Util.Concurrent
    Public Class ArrayBlockingQueue(Of T) : Implements IBlockingQueue(Of T)

        Private ReadOnly buffers As T()

        Private tail As Integer
        Private head As Integer
        Private bufferCount As Integer

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="capacity">キューの容量</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal capacity As Integer)
            If capacity < 1 Then
                Throw New ArgumentOutOfRangeException("capacity", capacity, "容量に1未満は指定できない")
            End If
            Me.buffers = VoUtil.NewArrayInstance(Of T)(capacity)
            Me.tail = 0
            Me.head = 0
            Me.bufferCount = 0
        End Sub

        ''' <summary>要素の数</summary>
        Public ReadOnly Property Count() As Integer
            Get
                Return bufferCount
            End Get
        End Property

        ''' <summary>
        ''' 指定された要素をこのキューの末尾に挿入する
        ''' </summary>
        ''' <param name="value">要素</param>
        ''' <remarks></remarks>
        Public Sub Put(ByVal value As T) Implements IBlockingQueue(Of T).Put
            SyncLock Me
                While buffers.Length <= bufferCount
                    Monitor.Wait(Me)
                End While
                buffers(tail) = value
                tail = (tail + 1) Mod buffers.Length
                bufferCount += 1
                Monitor.PulseAll(Me)
            End SyncLock
        End Sub

        ''' <summary>
        ''' このキューの先頭を取得して削除する
        ''' </summary>
        ''' <returns>キューの先頭</returns>
        ''' <remarks></remarks>
        Public Function Take() As T Implements IBlockingQueue(Of T).Take
            SyncLock Me
                While bufferCount <= 0
                    Monitor.Wait(Me)
                End While

                Dim result As T = buffers(head)
                head = (head + 1) Mod buffers.Length
                bufferCount -= 1
                Monitor.PulseAll(Me)
                Return result
            End SyncLock
        End Function
    End Class
End Namespace