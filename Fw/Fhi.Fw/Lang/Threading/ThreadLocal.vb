Namespace Lang.Threading
    ''' <summary>
    ''' スレッドローカル変数を担うクラス
    ''' </summary>
    ''' <typeparam name="T">変数の型</typeparam>
    ''' <remarks>JavaのThreadLocal相当</remarks>
    Public Class ThreadLocal(Of T)

        <ThreadStatic()> Private Shared _threadLocalValues As Hashtable

        Private Shared ReadOnly Property Values() As Hashtable
            Get
                If _threadLocalValues Is Nothing Then
                    _threadLocalValues = New Hashtable
                End If
                Return _threadLocalValues
            End Get
        End Property

        Private dlgt As InitialValueCallback

        ''' <summary>スレッド毎の初期値を返すデリゲート</summary>
        Public Delegate Function InitialValueCallback() As T

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me.New(Function() Nothing)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="dlgt">スレッド毎の初期値を返すデリゲート</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dlgt As InitialValueCallback)
            Me.dlgt = dlgt
        End Sub

        ''' <summary>
        ''' スレッド毎の初期値を返す
        ''' </summary>
        ''' <returns>スレッド毎の初期値</returns>
        ''' <remarks></remarks>
        Protected Overridable Function InitialValue() As T
            Return dlgt.Invoke
        End Function

        ''' <summary>
        ''' 現スレッドの値を返す
        ''' </summary>
        ''' <returns>現在のスレッドの値</returns>
        ''' <remarks></remarks>
        Public Function [Get]() As T
            If Not Values.ContainsKey(Me) Then
                Values.Add(Me, InitialValue)
            End If
            Return DirectCast(Values(Me), T)
        End Function

        ''' <summary>
        ''' 現スレッドの値を除去する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Remove()
            Values.Remove(Me)
        End Sub

        ''' <summary>
        ''' 現スレッドに値を設定する
        ''' </summary>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Sub [Set](ByVal value As T)
            If Values.ContainsKey(Me) Then
                Values.Remove(Me)
            End If
            Values.Add(Me, value)
        End Sub
    End Class
End Namespace