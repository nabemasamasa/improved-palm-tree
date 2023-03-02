Imports System.Threading

Namespace Util.Concurrent
    ''' <summary>
    ''' JavascriptのPromiseを模倣したクラス
    ''' </summary>
    ''' <see>https://github.com/robotlolita/robotlolita.github.io/tree/master/examples/promises</see>
    ''' <remarks>
    ''' 制限事項
    ''' 1. rejectできる値はExceptionに限定
    ''' 2. onRejectedでReturnできる型は、chain元のPromise値の型に限定
    ''' 3. onFulfilledでPromise(Of T)をReturnしたらT型の値を引き継げるが、Promise非総称型をReturnしたらT型(Object型)を引き継げない(どうやったら良いの？)
    ''' </remarks>
    Public Class Promise : Inherits Promise(Of Object)

#Region "Nested classes..."
        Private Class ThreadMessage
            Public HasResolved As Boolean = False
        End Class
#End Region
#If DEBUG Then
        ''' <summary>非同期処理を抑止する場合、true</summary>
        Public Shared DenyAsync As Boolean
#Else
        ' Releaseビルド時は抑止禁止。ビルドエラーにする。
        Public Shared ReadOnly DenyAsync As Boolean = False
#End If
        ''' <summary>非同期処理実行時の状態</summary>
        Public Shared ApartmentState As ApartmentState = Threading.ApartmentState.MTA

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="resolveOnlyCallback">resolveだけを引数に持つcallback</param>
        ''' <param name="denyAsyncConstructor">コンストラクタの非同期処理を抑止する場合、true（JSのPromiseと同等にするならtrue）</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal resolveOnlyCallback As Action(Of Action(Of Object)), Optional denyAsyncConstructor As Boolean = False)
            MyBase.New(resolveOnlyCallback, denyAsyncConstructor)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="resolveWithRejectCallback">resolve,rejectを引数に持つcallback</param>
        ''' <param name="denyAsyncConstructor">コンストラクタの非同期処理を抑止する場合、true（JSのPromiseと同等にするならtrue）</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal resolveWithRejectCallback As Action(Of Action(Of Object), Action(Of Exception)), Optional denyAsyncConstructor As Boolean = False)
            MyBase.New(resolveWithRejectCallback, denyAsyncConstructor)
        End Sub

        ''' <summary>
        ''' 結果正常のPromiseを作成する
        ''' </summary>
        ''' <param name="value">Promise値</param>
        ''' <returns>結果正常のPromise</returns>
        ''' <remarks></remarks>
        Public Shared Function Resolve(Optional value As Object = Nothing) As Promise
            Return New Promise(Sub(aResolve, aReject) aResolve(value), denyAsyncConstructor:=True)
        End Function

        ''' <summary>
        ''' 結果正常のPromiseを作成する
        ''' </summary>
        ''' <typeparam name="T">Promise値の型</typeparam>
        ''' <param name="value">Promise値</param>
        ''' <returns>結果正常のPromise</returns>
        ''' <remarks></remarks>
        Public Shared Function Resolve(Of T)(Optional value As T = Nothing) As Promise(Of T)
            Return New Promise(Of T)(Sub(aResolve, aReject) aResolve(value), denyAsyncConstructor:=True)
        End Function

        ''' <summary>
        ''' 結果異常のPromiseを作成する
        ''' </summary>
        ''' <param name="ex">異常値</param>
        ''' <returns>結果異常のPromise</returns>
        ''' <remarks></remarks>
        Public Shared Function Reject(ex As Exception) As Promise
            Return New Promise(Sub(aResolve, aReject) aReject(ex), denyAsyncConstructor:=True)
        End Function

        ''' <summary>
        ''' 結果異常のPromiseを作成する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="ex">異常値</param>
        ''' <returns>結果異常のPromise</returns>
        ''' <remarks></remarks>
        Public Shared Function Reject(Of T)(ex As Exception) As Promise(Of T)
            Return New Promise(Of T)(Sub(aResolve, aReject) aReject(ex), denyAsyncConstructor:=True)
        End Function

        ''' <summary>
        ''' 指定Promise全てがresolveされた時のPromiseを返す
        ''' </summary>
        ''' <typeparam name="T">Promise値の型</typeparam>
        ''' <param name="promises">Promise[]</param>
        ''' <returns>指定Promise[]の順序通りの結果[]</returns>
        ''' <remarks></remarks>
        Public Shared Function All(Of T)(ParamArray promises As Promise(Of T)()) As Promise(Of T())
            Return All(DirectCast(promises, IEnumerable(Of Promise(Of T))))
        End Function

        ''' <summary>
        ''' 指定Promise全てがresolveされた時のPromiseを返す
        ''' </summary>
        ''' <typeparam name="T">Promise値の型</typeparam>
        ''' <param name="promises">Promise[]</param>
        ''' <returns>指定Promise[]の順序通りの結果[]</returns>
        ''' <remarks></remarks>
        Public Shared Function All(Of T)(promises As IEnumerable(Of Promise(Of T))) As Promise(Of T())
            Dim pendingCount As Integer = promises.Count
            Dim values As T() = DirectCast(Array.CreateInstance(GetType(T), promises.Count), T())
            Dim message As New ThreadMessage
            Return New Promise(Of T())( _
                Sub(resolve, reject)
                    For i As Integer = 0 To promises.Count - 1
                        Dim index As Integer = i
                        promises(i).Then(Sub(val)
                                             SyncLock (message)
                                                 If message.HasResolved Then
                                                     Return
                                                 End If
                                                 values(index) = val
                                                 pendingCount -= 1
                                                 If pendingCount = 0 Then
                                                     message.HasResolved = True
                                                     resolve(values)
                                                 End If
                                             End SyncLock
                                         End Sub, _
                                         Sub(ex)
                                             SyncLock (message)
                                                 If message.HasResolved Then
                                                     Return
                                                 End If
                                                 message.HasResolved = True
                                                 reject(ex)
                                             End SyncLock
                                         End Sub)
                    Next
                End Sub, denyAsyncConstructor:=True)
        End Function

        ''' <summary>
        ''' 一番最初にresolveまたはrejectされた値をもつPromiseを返す
        ''' </summary>
        ''' <typeparam name="T">Promise値の型</typeparam>
        ''' <param name="promises">Promise[]</param>
        ''' <returns>一番最初にresolve/rejectした値のPromise</returns>
        ''' <remarks></remarks>
        Public Shared Function Race(Of T)(ParamArray promises As Promise(Of T)()) As Promise(Of T)
            Return Race(DirectCast(promises, IEnumerable(Of Promise(Of T))))
        End Function

        ''' <summary>
        ''' 一番最初にresolveまたはrejectされた値をもつPromiseを返す
        ''' </summary>
        ''' <typeparam name="T">Promise値の型</typeparam>
        ''' <param name="promises">Promise[]</param>
        ''' <returns>一番最初にresolve/rejectした値のPromise</returns>
        ''' <remarks></remarks>
        Public Shared Function Race(Of T)(promises As IEnumerable(Of Promise(Of T))) As Promise(Of T)
            Dim message As New ThreadMessage
            Return New Promise(Of T)( _
                Sub(resolve, reject)
                    For Each p As Promise(Of T) In promises
                        p.Then(Sub(val)
                                   SyncLock (message)
                                       If message.HasResolved Then
                                           Return
                                       End If
                                       message.HasResolved = True
                                       resolve(val)
                                   End SyncLock
                               End Sub, _
                               Sub(ex)
                                   SyncLock (message)
                                       If message.HasResolved Then
                                           Return
                                       End If
                                       message.HasResolved = True
                                       reject(ex)
                                   End SyncLock
                               End Sub)
                    Next
                End Sub, denyAsyncConstructor:=True)
        End Function

    End Class
    ''' <summary>
    ''' JavascriptのPromiseを模倣したクラス
    ''' </summary>
    ''' <typeparam name="T">Promiseが保持する値の型</typeparam>
    ''' <see>https://github.com/robotlolita/robotlolita.github.io/tree/master/examples/promises</see>
    ''' <remarks></remarks>
    Public Class Promise(Of T)

#Region "Nested classes..."
        Private Enum Status
            Pending
            Fulfilled
            Rejected
        End Enum
        Private Class Inner

            Private Shared Function IsObject(value As Object) As Boolean
                If value Is Nothing Then
                    Return False
                End If
                Return Not TypeUtil.IsTypeImmutable(value.GetType)
            End Function

            Public Shared Function IsCallable(callback As [Delegate]) As Boolean
                If callback Is Nothing Then
                    Return False
                End If
                Return True
            End Function

            Public Shared Sub EnqueueJob(act As Action)
                If Promise.DenyAsync Then
                    act.Invoke()
                Else
                    Const JOB_NAME As String = "PromiseJob"
                    Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId
                    Dim isRootThread As Boolean = Not StringUtil.Nvl(Thread.CurrentThread.Name).StartsWith(JOB_NAME)
                    Dim t As New Thread( _
                        DirectCast(Sub()
                                       Thread.CurrentThread.Name = String.Format("{0}:{1}{2}>{3}", JOB_NAME, If(isRootThread, "*", ""), threadId, Thread.CurrentThread.ManagedThreadId)
                                       act.Invoke()
                                   End Sub, ThreadStart))
                    t.SetApartmentState(Promise.ApartmentState)
                    t.Start()
                End If
            End Sub

            Private Shared Sub TriggerPromiseReactions(Of TValue)(value As TValue, reactions As IEnumerable(Of Action(Of TValue)))
                For Each reaction As Action(Of TValue) In reactions
                    Dim act As Action(Of TValue) = reaction
                    EnqueueJob(Sub() act(value))
                Next
            End Sub

            Public Shared Function Attempt(Of TValue)(reaction As Action(Of Object), fail As Action(Of Exception), transform As Func(Of TValue, Object)) As Action(Of TValue)
                Return Sub(value As TValue)
                           Try
                               Dim newValue As Object = transform.Invoke(value)
                               reaction.Invoke(newValue)
                           Catch ex As Exception
                               fail(ex)
                           End Try
                       End Sub
            End Function

            Public Shared Sub Reject(p As Promise(Of T), reason As Exception)
                Dim reactions As Action(Of Exception)()
                p.lockState.EnterWriteLock()
                Try
                    If p.state <> Status.Pending Then
                        Return
                    End If
                    p.state = Status.Rejected
                    p.value = reason
                    p.lockReaction.EnterWriteLock()
                    Try
                        reactions = p.rejectReactions.ToArray
                        p.rejectReactions.Clear()
                        p.fulfilReactions.Clear()
                    Finally
                        p.lockReaction.ExitWriteLock()
                    End Try
                Finally
                    p.lockState.ExitWriteLock()
                End Try
                TriggerPromiseReactions(reason, reactions)
            End Sub

            Private Shared Sub DoFulfil(p As Promise(Of T), resolution As T)
                Dim reactions As Action(Of T)()
                p.lockState.EnterWriteLock()
                Try
                    If p.state <> Status.Pending Then
                        Return
                    End If
                    p.state = Status.Fulfilled
                    p.value = resolution
                    p.lockReaction.EnterWriteLock()
                    Try
                        reactions = p.fulfilReactions.ToArray
                        p.rejectReactions.Clear()
                        p.fulfilReactions.Clear()
                    Finally
                        p.lockReaction.ExitWriteLock()
                    End Try
                Finally
                    p.lockState.ExitWriteLock()
                End Try
                TriggerPromiseReactions(resolution, reactions)
            End Sub

            Public Shared Sub Fulfil(p As Promise(Of T), resolution As T)
                p.lockState.EnterUpgradeableReadLock()
                Try
                    If p.state <> Status.Pending Then
                        Return
                    End If

                    Dim objResolution As Object = resolution
                    If objResolution Is p Then
                        Reject(p, New ArgumentException())
                    Else
                        If Not IsObject(resolution) Then
                            DoFulfil(p, resolution)
                        Else
                            Try
                                If TypeOf resolution Is Promise Then
                                    ' Rejectは良いが、Resolveの戻り値がPromise非総称型だと、TResultがPromise非総称型になるので（T型にできないので）戻り値はPromise非総称型で処理する
                                    EnqueueJob(Sub() DirectCast(objResolution, Promise).Then(Sub(value) DoFulfil(p, resolution), Sub(ex) Reject(p, ex)))
                                Else
                                    DoFulfil(p, resolution)
                                End If
                            Catch ex As Exception
                                Reject(p, ex)
                            End Try
                        End If
                    End If
                Finally
                    p.lockState.ExitUpgradeableReadLock()
                End Try
            End Sub
        End Class
#End Region

        Private state As Status
        Private value As Object
        Private ReadOnly rejectReactions As New List(Of Action(Of Exception))
        Private ReadOnly fulfilReactions As New List(Of Action(Of T))
        Private ReadOnly lockReaction As New ReaderWriterLockSlim
        Private ReadOnly lockState As New ReaderWriterLockSlim

        Private Shared Function ConvertFromResolveOnlyToWithReject(resolveOnlyCallback As Action(Of Action(Of T))) As Action(Of Action(Of T), Action(Of Exception))
            If resolveOnlyCallback Is Nothing Then
                Return Nothing
            End If
            Return Sub(resolve, reject)
                       resolveOnlyCallback.Invoke(resolve)
                   End Sub
        End Function

        Private Shared Sub AssertIsNotNullCallback(ByVal callback As Object)
            If callback Is Nothing Then
                Throw New ArgumentException("コールバックは必須", "callback")
            End If
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="resolveOnlyCallback">resolveだけを引数に持つcallback</param>
        ''' <param name="denyAsyncConstructor">コンストラクタの非同期処理を抑止する場合、true（JSのPromiseと同等にするならtrue）</param>
        ''' <remarks></remarks>
        Public Sub New(resolveOnlyCallback As Action(Of Action(Of T)), Optional denyAsyncConstructor As Boolean = False)
            Me.New(ConvertFromResolveOnlyToWithReject(resolveOnlyCallback), denyAsyncConstructor)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="resolveWithRejectCallback">resolve,rejectを引数に持つcallback</param>
        ''' <param name="denyAsyncConstructor">コンストラクタの非同期処理を抑止する場合、true（JSのPromiseと同等にするならtrue）</param>
        ''' <remarks></remarks>
        Public Sub New(resolveWithRejectCallback As Action(Of Action(Of T), Action(Of Exception)), Optional denyAsyncConstructor As Boolean = False)
            AssertIsNotNullCallback(resolveWithRejectCallback)
            Me.state = Status.Pending
            If denyAsyncConstructor Then
                InvokeResolveWithRejectCallback(resolveWithRejectCallback)
            Else
                Inner.EnqueueJob(Sub()
                                     InvokeResolveWithRejectCallback(resolveWithRejectCallback)
                                 End Sub)
            End If
        End Sub

        Private Sub InvokeResolveWithRejectCallback(ByVal resolveWithRejectCallback As Action(Of Action(Of T), Action(Of Exception)))
            Try
                resolveWithRejectCallback(Sub(value As T)
                                              Inner.Fulfil(Me, value)
                                          End Sub, _
                                          Sub(ex As Exception)
                                              Inner.Reject(Me, ex)
                                          End Sub)
            Catch ex As Exception
                Inner.Reject(Me, ex)
            End Try
        End Sub

        ''' <summary>
        ''' resolveした値を処理する
        ''' </summary>
        ''' <param name="onFulfilled">処理するAction</param>
        ''' <returns>Promise</returns>
        ''' <remarks></remarks>
        Public Function [Then](onFulfilled As Action(Of T)) As Promise(Of T)
            Return Me.[Then](onFulfilled, Nothing)
        End Function

        ''' <summary>
        ''' resolve、rejectした値を処理する
        ''' </summary>
        ''' <param name="onFulfilled">resolve値を処理するAction</param>
        ''' <param name="onRejected">reject値を処理するAction</param>
        ''' <returns>Promise</returns>
        ''' <remarks></remarks>
        Public Function [Then](onFulfilled As Action(Of T), onRejected As Action(Of Exception)) As Promise(Of T)
            Dim resolvedFulfilled As Func(Of T, T) = Nothing
            If onFulfilled IsNot Nothing Then
                resolvedFulfilled = Function(value)
                                        onFulfilled.Invoke(value)
                                        Return Nothing
                                    End Function
            End If
            Dim resolvedRejected As Func(Of Exception, T) = Nothing
            If onRejected IsNot Nothing Then
                resolvedRejected = Function(ex)
                                       onRejected(ex)
                                       Return Nothing
                                   End Function
            End If
            Return Me.[Then](resolvedFulfilled, resolvedRejected)
        End Function

        ''' <summary>
        ''' resolveした値を処理する
        ''' </summary>
        ''' <typeparam name="TResult">処理でreturnする型</typeparam>
        ''' <param name="onFulfilled">処理するAction</param>
        ''' <returns>returnした値を持つPromise</returns>
        ''' <remarks></remarks>
        Public Function [Then](Of TResult)(onFulfilled As Func(Of TResult)) As Promise(Of TResult)
            Dim resolvedFulfilled As Func(Of T, TResult) = Nothing
            If onFulfilled IsNot Nothing Then
                resolvedFulfilled = Function(value As T)
                                        Return onFulfilled.Invoke
                                    End Function
            End If
            Return Me.[Then](Of TResult)(resolvedFulfilled)
        End Function

        ''' <summary>
        ''' resolveした値を処理する
        ''' </summary>
        ''' <typeparam name="TResult">onFulfilledでreturnする型</typeparam>
        ''' <param name="onFulfilled">処理するAction</param>
        ''' <returns>returnした値を持つPromise</returns>
        ''' <remarks></remarks>
        Public Function [Then](Of TResult)(onFulfilled As Func(Of T, TResult)) As Promise(Of TResult)
            Return Me.[Then](Of TResult)(onFulfilled, DirectCast(Nothing, Action(Of Exception)))
        End Function

        ''' <summary>
        ''' resolve、rejectした値を処理する
        ''' </summary>
        ''' <typeparam name="TResult">onFulfilledでreturnする型</typeparam>
        ''' <param name="onFulfilled">resolve値を処理するAction</param>
        ''' <param name="onRejected">reject値を処理するAction</param>
        ''' <returns>returnした値を持つPromise</returns>
        ''' <remarks></remarks>
        Public Function [Then](Of TResult)(onFulfilled As Func(Of T, TResult), onRejected As Action(Of Exception)) As Promise(Of TResult)
            Dim resolvedRejected As Func(Of Exception, TResult) = Nothing
            If onRejected IsNot Nothing Then
                resolvedRejected = Function(ex)
                                       onRejected(ex)
                                       Return Nothing
                                   End Function
            End If
            Return Me.[Then](onFulfilled, resolvedRejected)
        End Function

        ''' <summary>
        ''' resolveした値を処理する
        ''' </summary>
        ''' <typeparam name="TResult">処理でreturnする型</typeparam>
        ''' <param name="onFulfilled">処理するAction</param>
        ''' <returns>returnした値を持つPromise</returns>
        ''' <remarks></remarks>
        Public Function [Then](Of TResult)(onFulfilled As Func(Of Promise(Of TResult))) As Promise(Of TResult)
            Dim resolvedFulfilled As Func(Of T, Promise(Of TResult)) = Nothing
            If onFulfilled IsNot Nothing Then
                resolvedFulfilled = Function(value As T)
                                        Return onFulfilled.Invoke
                                    End Function
            End If
            Return Me.[Then](Of TResult)(resolvedFulfilled)
        End Function

        ''' <summary>
        ''' resolveした値を処理する
        ''' </summary>
        ''' <typeparam name="TResult">onFulfilledでreturnする型</typeparam>
        ''' <param name="onFulfilled">処理するAction</param>
        ''' <returns>returnした値を持つPromise</returns>
        ''' <remarks></remarks>
        Public Function [Then](Of TResult)(onFulfilled As Func(Of T, Promise(Of TResult))) As Promise(Of TResult)
            Return Me.[Then](Of TResult)(onFulfilled, DirectCast(Nothing, Action(Of Exception)))
        End Function

        ''' <summary>
        ''' resolve、rejectした値を処理する
        ''' </summary>
        ''' <typeparam name="TResult">onFulfilledでreturnする型</typeparam>
        ''' <param name="onFulfilled">resolve値を処理するAction</param>
        ''' <param name="onRejected">reject値を処理するAction</param>
        ''' <returns>returnした値を持つPromise</returns>
        ''' <remarks></remarks>
        Public Function [Then](Of TResult)(onFulfilled As Func(Of T, Promise(Of TResult)), onRejected As Action(Of Exception)) As Promise(Of TResult)
            Dim resolvedRejected As Func(Of Exception, TResult) = Nothing
            If onRejected IsNot Nothing Then
                resolvedRejected = Function(ex)
                                       onRejected(ex)
                                       Return Nothing
                                   End Function
            End If
            Return Me.[Then](onFulfilled, resolvedRejected)
        End Function

        ''' <summary>
        ''' resolve、rejectした値を処理する
        ''' </summary>
        ''' <typeparam name="TResult">onFulfilledでreturnする型</typeparam>
        ''' <param name="onFulfilled">resolve値を処理するAction</param>
        ''' <param name="onRejected">reject値を処理するAction</param>
        ''' <returns>returnした値を持つPromise</returns>
        ''' <remarks></remarks>
        Public Function [Then](Of TResult)(onFulfilled As Func(Of T, Promise(Of TResult)), onRejected As Func(Of Exception, TResult)) As Promise(Of TResult)
            Dim callback As Action(Of Action(Of TResult), Action(Of Exception)) = CreateCallbackForThen(Of TResult, Promise(Of TResult))(onFulfilled, onRejected, isFulfilResultPromise:=True)
            Return New Promise(Of TResult)(callback, denyAsyncConstructor:=True)
        End Function

        ''' <summary>
        ''' resolve、rejectした値を処理する
        ''' </summary>
        ''' <typeparam name="TResult">onFulfilledでreturnする型</typeparam>
        ''' <param name="onFulfilled">resolve値を処理するAction</param>
        ''' <param name="onRejected">reject値を処理するAction</param>
        ''' <returns>returnした値を持つPromise</returns>
        ''' <remarks></remarks>
        Public Function [Then](Of TResult)(onFulfilled As Func(Of T, TResult), onRejected As Func(Of Exception, TResult)) As Promise(Of TResult)
            Dim callback As Action(Of Action(Of TResult), Action(Of Exception)) = CreateCallbackForThen(Of TResult, TResult)(onFulfilled, onRejected, isFulfilResultPromise:=False)
            Return New Promise(Of TResult)(callback, denyAsyncConstructor:=True)
        End Function

        Private Function CreateCallbackForThen(Of TResult, TFuncResult)(ByVal onFulfilled As Func(Of T, TFuncResult), ByVal onRejected As Func(Of Exception, TResult), isFulfilResultPromise As Boolean) As Action(Of Action(Of TResult), Action(Of Exception))
            Dim this As Promise(Of T) = Me
            Return Sub(resolve As Action(Of TResult), reject As Action(Of Exception))
                       If Not Inner.IsCallable(onFulfilled) Then
                           onFulfilled = Function(value As T)
                                             resolve(DirectCast(this.value, TResult))
                                             Return Nothing
                                         End Function
                       End If
                       If Not Inner.IsCallable(onRejected) Then
                           onRejected = Function(value As Exception)
                                            reject(DirectCast(this.value, Exception))
                                            Return Nothing
                                        End Function
                       End If
                       Dim reaction As Action(Of Object) = Sub(val As Object)
                                                               If isFulfilResultPromise Then
                                                                   If GetType(Promise(Of TResult)).IsAssignableFrom(val.GetType) Then
                                                                       DirectCast(val, Promise(Of TResult)).Then(Sub(v) resolve(v), Sub(ex) reject(ex))
                                                                   Else
                                                                       reject(New InvalidProgramException("あり得ないはず"))
                                                                   End If
                                                               Else
                                                                   resolve(DirectCast(val, TResult))
                                                               End If
                                                           End Sub
                       Dim onSuccess As Func(Of T, Object) = Function(val As T) onFulfilled.Invoke(val)
                       Dim onFailure As Func(Of Exception, Object) = Function(val As Exception) onRejected.Invoke(val)
                       lockState.EnterReadLock()
                       Try
                           Select Case this.state
                               Case Status.Pending
                                   lockReaction.EnterWriteLock()
                                   Try
                                       this.fulfilReactions.Add(Inner.Attempt(reaction, reject, onSuccess))
                                       this.rejectReactions.Add(Inner.Attempt(Of Exception)(reaction, reject, onFailure))
                                   Finally
                                       lockReaction.ExitWriteLock()
                                   End Try
                               Case Status.Fulfilled
                                   Inner.EnqueueJob(Sub() Inner.Attempt(reaction, reject, onSuccess)(DirectCast(this.value, T)))
                               Case Status.Rejected
                                   Inner.EnqueueJob(Sub() Inner.Attempt(reaction, reject, onFailure)(DirectCast(this.value, Exception)))
                           End Select

                       Finally
                           lockState.ExitReadLock()
                       End Try

                   End Sub
        End Function

        ''' <summary>
        ''' 発生した例外、またはrejectした値を処理する
        ''' </summary>
        ''' <param name="onRejected">処理するAction</param>
        ''' <returns>Promise</returns>
        ''' <remarks></remarks>
        Public Function [Catch](onRejected As Action(Of Exception)) As Promise(Of T)
            Return Me.[Catch](Function(ex)
                                  onRejected(ex)
                                  Return Nothing
                              End Function)
        End Function

        ''' <summary>
        ''' 発生した例外、またはrejectした値を処理する
        ''' </summary>
        ''' <param name="onRejected">処理するAction</param>
        ''' <returns>Promise</returns>
        ''' <remarks></remarks>
        Public Function [Catch](onRejected As Func(Of Exception, T)) As Promise(Of T)
            Return Me.[Then](Of T)(DirectCast(Nothing, Func(Of T, T)), onRejected)
        End Function

    End Class
End Namespace
