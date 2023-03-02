Imports NUnit.Framework
Imports System.Threading
Imports Fhi.Fw.TestUtil

Namespace Util.Concurrent
    Public MustInherit Class PromiseTest

#Region "Nested classes..."
        Private Enum Status
            Pending
            Fulfilled
            Rejected
        End Enum
        Private Class TestingException : Inherits Exception
            Public Sub New(ByVal message As String)
                MyBase.New(message)
            End Sub
        End Class
        Private Class TestingHelper
            Private _state As [Enum] = Status.Pending
            Public Property State As Status
                Get
                    Return DirectCast(Volatile.Read(_state), Status)
                End Get
                Set(value As Status)
                    Volatile.Write(_state, value)
                End Set
            End Property
            Public Arg As Object
            Public Sub Wait(Of T)(p As Promise(Of T))
                State = Status.Pending
                Dim p2 As Promise(Of T) = p.Then(Sub(val)
                                                     Arg = val
                                                     State = Status.Fulfilled
                                                 End Sub, _
                                                 Sub(err)
                                                     Arg = err
                                                     State = Status.Rejected
                                                 End Sub)
                For i As Integer = 0 To 10
                    Thread.Sleep(20)
                    If State <> Status.Pending Then
                        Return
                    End If
                Next
                Assert.Fail("timeout on Helper!!")
            End Sub
        End Class
#End Region
        Private Shared Sub WaitAndResolve(resolve As Action(Of Integer), delay As Integer)
            Call (New Thread(CType(Sub()
                                       Thread.Sleep(delay)
                                       resolve(delay)
                                   End Sub, ThreadStart))).Start()
        End Sub

        Private Shared Sub WaitAndReject(reject As Action(Of Exception), delay As Integer)
            Call (New Thread(CType(Sub()
                                       Thread.Sleep(delay)
                                       reject(New TestingException(delay.ToString))
                                   End Sub, ThreadStart))).Start()
        End Sub

        Public Shared Function InspectValue(Of T)(p As Promise(Of T)) As Object
            Return NonPublicUtil.GetField(p, "value")
        End Function

        Public Shared Function InspectState(Of T)(p As Promise(Of T)) As Object
            Return NonPublicUtil.GetField(p, "state")
        End Function

        Private actual As TestingHelper

        <SetUp()> Public Overridable Sub SetUp()
            Promise.DenyAsync = False
            actual = New TestingHelper
        End Sub

        Public Class コンストラクタTest : Inherits PromiseTest

            <Test()> Public Sub 正常終了時はresolve呼んで_値を渡す_ThenのonFulfilledが動く_resolve引数だけのラムダでも動く()
                Dim actual As New TestingHelper
                Dim p As New Promise(Sub(resolve)
                                         resolve("Hoge")
                                     End Sub)
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(actual.Arg, [Is].EqualTo("Hoge"))
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub 正常終了時はresolve呼んで_値を渡す_ThenのonFulfilledが動く_reject有りラムダでも動く()
                Dim actual As New TestingHelper
                Dim p As New Promise(Sub(resolve, rejected)
                                         resolve("Success")
                                     End Sub)
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(actual.Arg, [Is].EqualTo("Success"))
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub 異常終了時はreject呼んで例外を渡す_ThenのonRejectedが動く()
                Dim p As New Promise(Sub(resolve, rejected)
                                         rejected(New InvalidCastException("castEx"))
                                     End Sub)
                Dim state As Status
                Dim actual As Exception = Nothing
                p.Then(onFulfilled:=Sub(value)
                                        state = Status.Fulfilled
                                    End Sub, _
                       onRejected:=Sub(ex)
                                       state = Status.Rejected
                                       actual = ex
                                   End Sub)

                For i As Integer = 0 To 10
                    Thread.Sleep(10)
                    If state <> Status.Pending Then
                        Assert.That(state, [Is].EqualTo(Status.Rejected))
                        Assert.That(actual.Message, [Is].EqualTo("castEx"))
                        Assert.Pass("OK")
                    End If
                Next
                Assert.Fail("ThenもCatchも通っていない")
            End Sub

            <Test()> Public Sub 異常終了時はreject呼んで例外を渡す_CatchのonRejectedが動く()
                Dim p As New Promise(Sub(resolve, rejected)
                                         rejected(New InvalidCastException("castEx"))
                                     End Sub)
                Dim state As Status
                Dim actual As Exception = Nothing
                p.Then(onFulfilled:=Sub(value)
                                        state = Status.Fulfilled
                                    End Sub) _
                 .Catch(onRejected:=Sub(ex)
                                        state = Status.Rejected
                                        actual = ex
                                    End Sub)

                For i As Integer = 0 To 10
                    Thread.Sleep(10)
                    If state <> Status.Pending Then
                        Assert.That(state, [Is].EqualTo(Status.Rejected))
                        Assert.That(actual.Message, [Is].EqualTo("castEx"))
                        Assert.Pass("OK")
                    End If
                Next
                Assert.Fail("ThenもCatchも通っていない")
            End Sub

            <Test()> Public Sub callback引数省略は例外_引数1つ()
                Try
                    Dim p As New Promise(DirectCast(Nothing, Action(Of Action(Of Object))))
                    Assert.Fail()
                Catch ex As ArgumentException
                    Assert.That(ex.Message, [Is].StringStarting("コールバックは必須"))
                End Try
            End Sub

            <Test()> Public Sub callback引数省略は例外_引数2つ()
                Try
                    Dim p As New Promise(DirectCast(Nothing, Action(Of Action(Of Object), Action(Of Exception))))
                    Assert.Fail()
                Catch ex As ArgumentException
                    Assert.That(ex.Message, [Is].StringStarting("コールバックは必須"))
                End Try
            End Sub

        End Class

        Public Class Then_Catch_を非同期でTest : Inherits PromiseTest

            <Test()> Public Sub 文字列を流す()
                Dim p As New Promise(Sub(resolve, rejected)
                                         resolve("Hoge")
                                     End Sub)
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(actual.Arg, [Is].EqualTo("Hoge"))
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub コンストラクタを同期実行にすれば_outer_の前に_inner_が必ず実行される()
                Dim sequences As New List(Of String)
                Dim p As New Promise(Sub(resolve, rejected)
                                         sequences.Add("inner")
                                         resolve("Then")
                                     End Sub, denyAsyncConstructor:=True)
                p.Then(onFulfilled:=Sub(value)
                                        sequences.Add(value.ToString)
                                    End Sub)
                sequences.Add("outer")

                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(Join(sequences.ToArray, ","), [Is].EqualTo("inner,outer,Then").Or.EqualTo("inner,Then,outer"), _
                            "Thenは非同期なので、Thenとouterの順序は不定（9割方outerが先だけど）")
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

        End Class

        Public Class Then_Catch_を同期させてTest : Inherits PromiseTest

            Public Overrides Sub SetUp()
                MyBase.SetUp()
                Promise.DenyAsync = True
            End Sub

            <Test()> Public Sub メソッドチェーンの順に処理される()
                Dim sequences As New List(Of Integer)
                Dim p As Promise(Of Integer) = Promise.Resolve(0)
                p.Then(Sub() sequences.Add(1)) _
                    .Then(Sub() sequences.Add(2)) _
                    .Catch(Sub(e) sequences.Add(3)) _
                    .Then(Sub() sequences.Add(4))

                Assert.That(Join(sequences.Select(Function(s) s.ToString).ToArray, ","), [Is].EqualTo("1,2,4"), "rejectしてないから3は通らない")
            End Sub

            <Test()> Public Sub Thenでreturnした値は_次のThenの引数になる()
                Dim actual As String = Nothing
                Dim p As Promise(Of Integer) = Promise.Resolve(2)
                Dim p2 As Promise(Of Integer) = p.Then(Function(val) val + 3)
                Dim p3 As Promise(Of String) = p2.Then(Function(val) "*" & val.ToString)
                p3.Then(Sub(val) actual = val, Sub(ex) actual = ex.Message)

                Assert.That(actual, [Is].EqualTo("*5"))
            End Sub

            <Test()> Public Sub Thenは常に新しいインスタンス_独立してるからこの使い方だとデータは流れない()
                Dim actual As Integer = 0
                Dim p As Promise(Of Integer) = Promise.Resolve(100)
                p.Then(Function(value) value * 2)
                p.Then(Function(value) value * 2)
                p.Then(Sub(value) actual = value)

                Assert.That(actual, [Is].EqualTo(100), "初期値100のまま")
            End Sub

            <Test()> Public Sub Thenは常に新しいインスタンス_メソッドチェーンすればデータは流れる()
                Dim actual As Integer = 0
                Dim p As Promise(Of Integer) = Promise.Resolve(100)
                p.Then(Function(value) value * 2) _
                    .Then(Function(value) value * 2) _
                    .Then(Sub(value) actual = value)

                Assert.That(actual, [Is].EqualTo(400), "100*2*2")
            End Sub

            <Test()> Public Sub fulfilでエラー発生時は_同時指定のrejectでcatchできない_次以降のrejectでcatchできる()
                Dim actual As String = Nothing
                Dim p As Promise(Of String) = Promise.Resolve("not a number")
                Dim p2 As Promise(Of String) = p.Then(Function(val)
                                                          Dim s = CInt(val) ' ここで例外
                                                          Return "*" & s
                                                      End Function, _
                                                      Function(ex)
                                                          Return "(Current)" & ex.Message
                                                      End Function)
                p2.Then(Sub(val) actual = val, Sub(ex) actual = "(Next)" & ex.Message)

                Assert.That(actual, [Is].EqualTo("(Next)String ""not a number"" から型 'Integer' への変換は無効です。").Or.EqualTo("(Next)Conversion from string ""not a number"" to type 'Integer' is not valid."), "次のThenでcatchされる")
            End Sub

            <Test()> Public Sub Catchが正常終了なら_Thenで処理できる()
                Dim actual As Object = Nothing
                Dim p As Promise(Of String) = Promise.Resolve("not a number")
                Dim p2 As Promise(Of Integer) = p.Then(Function(val)
                                                           Return CInt(val) ' ここで例外
                                                       End Function)
                Dim p3 As Promise(Of String) = p2.Then(Function(val)
                                                           Return "*" & val ' ここはすっ飛ばされる
                                                       End Function)
                Dim p4 As Promise(Of String) = p3.Catch(Function(ex)
                                                            Return ex.Message
                                                        End Function)
                p4.Then(Sub(val) actual = val, Sub(ex) actual = ex)

                Assert.That(actual, [Is].EqualTo("String ""not a number"" から型 'Integer' への変換は無効です。").Or.EqualTo("Conversion from string ""not a number"" to type 'Integer' is not valid."))
            End Sub

        End Class

        Public Class PromiseをreturnするTest : Inherits PromiseTest

            <Test()> Public Sub ThenでResolveを返せば_onFulfilledが呼ばれる(<Values(False, True)> denyAsync As Boolean)
                Promise.DenyAsync = denyAsync
                Dim p As Promise = Promise.Resolve
                Dim p2 = p.Then(Of String)(Function()
                                               Return Promise.Resolve("ABC")
                                           End Function)
                actual.Wait(p2)

                Assert.That(InspectState(p2).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(actual.Arg, [Is].InstanceOf(GetType(String)))
                Assert.That(actual.Arg, [Is].EqualTo("ABC"))
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub ThenでPromise非総称型のResolveは_型が不明だからPromiseをそのまま返す(<Values(False, True)> denyAsync As Boolean)
                Promise.DenyAsync = denyAsync
                Dim p As Promise = Promise.Resolve
                Dim p2 = p.Then(Function()
                                    ' こうすると無理やり非総称型のPromiseをreturnできる
                                    Return New Promise(Sub(resolve) resolve("ABC"))
                                End Function)
                actual.Wait(p2)

                Assert.That(InspectState(p2).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(actual.Arg, [Is].InstanceOf(GetType(Promise)))
                Assert.That(InspectValue(DirectCast(actual.Arg, Promise)), [Is].EqualTo("ABC"))
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub ThenでRejectを返せば_onRejectが呼ばれる(<Values(False, True)> denyAsync As Boolean)
                Promise.DenyAsync = denyAsync
                Dim p As Promise = Promise.Resolve
                Dim p2 = p.Then(Function()
                                    Return Promise.Reject(New InvalidOperationException("ThenでReject"))
                                End Function)
                actual.Wait(p2)

                Assert.That(InspectState(p2).ToString, [Is].EqualTo(Status.Rejected.ToString))
                Assert.That(actual.Arg, [Is].InstanceOf(GetType(InvalidOperationException)))
                Assert.That(DirectCast(actual.Arg, Exception).Message, [Is].EqualTo("ThenでReject"))
                Assert.That(actual.State, [Is].EqualTo(Status.Rejected))
            End Sub

            <Test()> Public Sub ThenでRejectを返せば_Rejectが呼ばれる_PromiseOfTでも動作する()
                Dim p As Promise(Of String) = Promise.Resolve("a")
                Dim p2 = p.Then(Of String)(Function()
                                               Return New Promise(Of String)(Sub(resolve, reject) reject(New InvalidOperationException("ThenでReject")))
                                           End Function)
                actual.Wait(p2)
                Assert.That(InspectState(p2).ToString, [Is].EqualTo(Status.Rejected.ToString))
                Assert.That(actual.Arg, [Is].InstanceOf(GetType(InvalidOperationException)))
                Assert.That(DirectCast(actual.Arg, Exception).Message, [Is].EqualTo("ThenでReject"))
                Assert.That(actual.State, [Is].EqualTo(Status.Rejected))
            End Sub

        End Class

        Public Class ALL_は引数すべてがresolveした結果を返す_Test : Inherits PromiseTest

            <Test()> Public Sub 同期実行_引数順に結果が返る()
                Promise.DenyAsync = True
                Dim p As Promise(Of Integer()) = Promise.All(Of Integer)(Promise.Resolve(1), Promise.Resolve(2), Promise.Resolve(3))
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(actual.Arg, [Is].InstanceOf(GetType(Integer())))
                Assert.That(Join(DirectCast(actual.Arg, Integer()).Select(Function(i) i.ToString).ToArray, ","), [Is].EqualTo("1,2,3"))
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub 同期実行_引数のどれか一つでもrejectしたら_resolveした結果があっても取得できない()
                Promise.DenyAsync = True
                Dim p As Promise(Of Integer()) = Promise.All(Of Integer)(Promise.Resolve(1), Promise.Reject(Of Integer)(New TestingException("err")), Promise.Resolve(3))
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Rejected.ToString))
                Assert.That(actual.Arg, [Is].InstanceOf(GetType(TestingException)))
                Assert.That(DirectCast(actual.Arg, TestingException).Message, [Is].EqualTo("err"))
                Assert.That(actual.State, [Is].EqualTo(Status.Rejected))
            End Sub

            <Test()> Public Sub 引数の結果は_終了順ではなく_引数順になる()
                Dim p As Promise(Of Integer()) = Promise.All(Of Integer)(New Promise(Of Integer)(Sub(resolve) WaitAndResolve(resolve, 64)), _
                                                                         New Promise(Of Integer)(Sub(resolve) WaitAndResolve(resolve, 32)), _
                                                                         New Promise(Of Integer)(Sub(resolve) WaitAndResolve(resolve, 1)))
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(actual.Arg, [Is].InstanceOf(GetType(Integer())))
                Assert.That(Join(DirectCast(actual.Arg, Integer()).Select(Function(i) i.ToString).ToArray, ","), [Is].EqualTo("64,32,1"), "64ms必要な処理でも配列index=0で返る")
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

        End Class

        Public Class Race_一番最初に_resolve_reject_した結果を返す_Test : Inherits PromiseTest

            <Test()> Public Sub 同期実行だと_第一引数の結果になる_resolve()
                Promise.DenyAsync = True
                Dim p As Promise(Of Integer) = Promise.Race(Of Integer)(Promise.Resolve(3), Promise.Resolve(2), Promise.Resolve(1))
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(actual.Arg, [Is].EqualTo(3), "同期実行だから、常に先頭の3")
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub 同期実行だと_第一引数の結果になる_reject()
                Promise.DenyAsync = True
                Dim p As Promise(Of Integer) = Promise.Race(Of Integer)(Promise.Reject(Of Integer)(New TestingException("ww")), Promise.Resolve(2), Promise.Resolve(1))
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Rejected.ToString))
                Assert.That(actual.Arg, [Is].InstanceOf(GetType(TestingException)))
                Assert.That(DirectCast(actual.Arg, Exception).Message, [Is].EqualTo("ww"), "同期実行だから、常に先頭のreject")
                Assert.That(actual.State, [Is].EqualTo(Status.Rejected))
            End Sub

            <Test()> Public Sub 一番最初の_1msでresolveする処理の結果が返る()
                Dim p As Promise(Of Integer) = Promise.Race(Of Integer)(New Promise(Of Integer)(Sub(resolve) WaitAndResolve(resolve, 64)), _
                                                                        New Promise(Of Integer)(Sub(resolve) WaitAndResolve(resolve, 32)), _
                                                                        New Promise(Of Integer)(Sub(resolve) WaitAndResolve(resolve, 1)))
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(actual.Arg, [Is].EqualTo(1))
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub 一番最初の_2msでresolveする処理の結果が返る_あとの処理がrejectしてても関係無い()
                Dim p As Promise(Of Integer) = Promise.Race(Of Integer)(New Promise(Of Integer)(Sub(resolve, reject) WaitAndReject(reject, 64)), _
                                                                        New Promise(Of Integer)(Sub(resolve, reject) WaitAndReject(reject, 32)), _
                                                                        New Promise(Of Integer)(Sub(resolve, reject) WaitAndResolve(resolve, 2)))
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(actual.Arg, [Is].EqualTo(2))
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub 一番最初の_3msでrejectする処理の結果が返る()
                Dim p As Promise(Of Integer) = Promise.Race(Of Integer)(New Promise(Of Integer)(Sub(resolve, reject) WaitAndReject(reject, 64)), _
                                                                        New Promise(Of Integer)(Sub(resolve, reject) WaitAndReject(reject, 32)), _
                                                                        New Promise(Of Integer)(Sub(resolve, reject) WaitAndReject(reject, 3)))
                actual.Wait(p)

                Assert.That(InspectState(p).ToString, [Is].EqualTo(Status.Rejected.ToString))
                Assert.That(actual.Arg, [Is].InstanceOf(GetType(TestingException)))
                Assert.That(DirectCast(actual.Arg, Exception).Message, [Is].EqualTo("3"))
                Assert.That(actual.State, [Is].EqualTo(Status.Rejected))
            End Sub

        End Class

        Public Class Spec_getify_nativePromiseOnly_Test : Inherits PromiseTest

            <Test()> Public Sub 例外をsubでcatchしたあとは_引数nullでfulfilが動く()
                Dim ex1 As Exception = Nothing
                Dim p As Promise = Promise.Resolve()
                Dim p2 As Promise(Of Object) = p.Then(Sub()
                                                          Throw New TestingException("abc")
                                                      End Sub)
                Dim p3 As Promise(Of Object) = p2.Catch(Sub(ex)
                                                            ex1 = ex
                                                        End Sub)
                actual.Wait(p3)

                Assert.That(InspectState(p2).ToString, [Is].EqualTo(Status.Rejected.ToString))
                Assert.That(InspectState(p3).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(ex1, [Is].InstanceOf(GetType(TestingException)))
                Assert.That(ex1.Message, [Is].EqualTo("abc"))
                Assert.That(actual.Arg, [Is].Null)
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub 例外をcatchしてreturnしたら_引数ありでfulfilが動く()
                Dim ex1 As Exception = Nothing
                Dim p As Promise = Promise.Resolve()
                Dim p2 As Promise(Of Object) = p.Then(Sub()
                                                          Throw New TestingException("abc")
                                                      End Sub)
                Dim p3 As Promise(Of Object) = p2.Catch(Function(ex)
                                                            ex1 = ex
                                                            Return "xyz"
                                                        End Function)
                actual.Wait(p3)

                Assert.That(InspectState(p2).ToString, [Is].EqualTo(Status.Rejected.ToString))
                Assert.That(InspectState(p3).ToString, [Is].EqualTo(Status.Fulfilled.ToString))
                Assert.That(ex1, [Is].InstanceOf(GetType(TestingException)))
                Assert.That(ex1.Message, [Is].EqualTo("abc"))
                Assert.That(actual.Arg, [Is].EqualTo("xyz"))
                Assert.That(actual.State, [Is].EqualTo(Status.Fulfilled))
            End Sub

            <Test()> Public Sub 例外をcatchして_別の例外をrejectでreturnしたら_その例外がRejectされる()
                Dim ex1 As Exception = Nothing
                Dim p As Promise = Promise.Resolve()
                Dim p2 As Promise(Of Object) = p.Then(Sub()
                                                          Throw New TestingException("abc")
                                                      End Sub)
                Dim p3 As Promise(Of Object) = p2.Catch(Function(ex)
                                                            ex1 = ex
                                                            Return Promise.Reject(New TestingException("123"))
                                                        End Function)
                actual.Wait(p3)

                Assert.That(InspectState(p2).ToString, [Is].EqualTo(Status.Rejected.ToString))
                Assert.That(InspectState(p3).ToString, [Is].EqualTo(Status.Rejected.ToString))
                Assert.That(ex1, [Is].InstanceOf(GetType(TestingException)))
                Assert.That(ex1.Message, [Is].EqualTo("abc"))
                Assert.That(actual.Arg, [Is].InstanceOf(GetType(TestingException)))
                Assert.That(DirectCast(actual.Arg, Exception).Message, [Is].EqualTo("123"))
                Assert.That(actual.State, [Is].EqualTo(Status.Rejected))
            End Sub

        End Class

    End Class
End Namespace
