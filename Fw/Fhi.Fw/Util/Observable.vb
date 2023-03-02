Imports System.Threading

Namespace Util
    ''' <summary>
    ''' 監視オブジェクト/被監視オブジェクトにおける被監視オブジェクト
    ''' </summary>
    ''' <remarks>JavaのObservableの模倣</remarks>
    Public Class Observable
        Private changed As Boolean = False
        Private ReadOnly observers As New List(Of Observer)
        Private ReadOnly observersLock As New ReaderWriterLockSlim
        Private ReadOnly changedLock As New ReaderWriterLockSlim

        ''' <summary>
        ''' オブジェクトのオブザーバセットにオブザーバを追加します。
        ''' ただし、すでにオブザーバセットにあるオブザーバと同じでないものに限ります。
        ''' 複数のオブザーバに通知が配信される順序は指定されません。
        ''' クラスのコメントを参照してください。
        ''' </summary>
        ''' <param name="observer">追加するオブザーバ</param>
        ''' <remarks></remarks>
        Public Overridable Sub AddObserver(ByVal observer As Observer)
            If observer Is Nothing Then
                Throw New NullReferenceException
            End If

            observersLock.TryEnterWriteLock(Timeout.Infinite)
            Try
                If Not observers.Contains(observer) Then
                    Call observers.Add(observer)
                End If
            Finally
                observersLock.ExitWriteLock()
            End Try
        End Sub

        ''' <summary>
        ''' オブジェクトのオブザーバセットからオブザーバを削除します。
        ''' このメソッドに null を渡しても、何の効果もありません。
        ''' </summary>
        ''' <param name="observer">削除するオブザーバ</param>
        ''' <remarks></remarks>
        Public Overridable Sub DeleteObserver(ByVal observer As Observer)
            observersLock.TryEnterWriteLock(Timeout.Infinite)
            Try
                observers.Remove(observer)
            Finally
                observersLock.ExitWriteLock()
            End Try
        End Sub

        ''' <summary>
        ''' オブジェクトが、hasChanged メソッドに示されるように変更されていた場合、
        ''' そのすべてのオブザーバにそのことを通知し、次に clearChanged メソッドを
        ''' 呼び出して、このオブジェクトがもはや変更された状態でないことを示します。 
        ''' 
        ''' 各オブザーバの update メソッドが 2 つの引数 (この被監視オブジェクトと null) 
        ''' で呼び出されます。つまり、このメソッドは次と同じになります。
        ''' 
        ''' notifyObservers(null)
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub NotifyObservers()
            Call NotifyObservers(Nothing)
        End Sub

        ''' <summary>
        ''' オブジェクトが、hasChanged メソッドに示されるように変更されていた場合、
        ''' そのすべてのオブザーバにそのことを通知し、次に clearChanged メソッドを
        ''' 呼び出して、このオブジェクトがもはや変更された状態でないことを示します。
        ''' 
        ''' 各オブザーバの update メソッドが 2 つの引数 (この被監視オブジェクトと arg) で呼び出されます。 
        ''' </summary>
        ''' <param name="arg">任意のオブジェクト</param>
        ''' <remarks></remarks>
        Public Overridable Sub NotifyObservers(ByVal arg As Object)

            changedLock.TryEnterWriteLock(Timeout.Infinite)
            Try
                If Not changed Then
                    Return
                End If
                changed = False
            Finally
                changedLock.ExitWriteLock()
            End Try

            Dim localObservers As Observer()
            observersLock.TryEnterReadLock(Timeout.Infinite)
            Try
                localObservers = observers.ToArray()
            Finally
                observersLock.ExitReadLock()
            End Try
            For Each observer As Observer In localObservers
                observer.Update(Me, arg)
            Next
        End Sub

        ''' <summary>
        ''' オブザーバリストを消去します。この結果、オブジェクトのオブザーバは
        ''' 存在しなくなります。
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub DeleteObservers()
            observersLock.TryEnterWriteLock(Timeout.Infinite)
            Try
                observers.Clear()
            Finally
                observersLock.ExitWriteLock()
            End Try
        End Sub

        ''' <summary>
        ''' Observable オブジェクトを変更されたものとしてマーキングします。
        ''' hasChanged メソッドは true を返すようになります。 
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overridable Sub SetChanged()
            changedLock.TryEnterWriteLock(Timeout.Infinite)
            Try
                changed = True
            Finally
                changedLock.ExitWriteLock()
            End Try
        End Sub

        ''' <summary>
        ''' オブジェクトがもはや変更された状態ではないこと、すなわち、
        ''' 最新の変更がすべてオブザーバに通知されたことを示します。
        ''' これにより、hasChanged メソッドは false を返します。
        ''' このメソッドは、notifyObservers メソッドによって自動的に呼び出されます。 
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overridable Sub ClearChanged()
            changedLock.TryEnterWriteLock(Timeout.Infinite)
            Try
                changed = False
            Finally
                changedLock.ExitWriteLock()
            End Try
        End Sub

        ''' <summary>
        ''' オブジェクトが変更されたかどうかを判定します。
        ''' </summary>
        ''' <returns>オブジェクトに対し、clearChanged メソッドよりあとに setChanged メソッドが呼び出された場合にだけ true、そうでない場合は false</returns>
        ''' <remarks></remarks>
        Public Overridable Function HasChanged() As Boolean
            changedLock.TryEnterReadLock(Timeout.Infinite)
            Try
                Return changed
            Finally
                changedLock.ExitReadLock()
            End Try
        End Function

        ''' <summary>
        ''' Observable オブジェクトのオブザーバの数を返します。 
        ''' </summary>
        ''' <returns>オブジェクトのオブザーバの数</returns>
        ''' <remarks></remarks>
        Public Overridable Function CountObservers() As Integer
            observersLock.TryEnterReadLock(Timeout.Infinite)
            Try
                Return observers.Count
            Finally
                observersLock.ExitReadLock()
            End Try
        End Function
    End Class
End Namespace