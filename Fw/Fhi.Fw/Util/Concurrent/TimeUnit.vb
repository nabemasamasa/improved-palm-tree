Imports System.Threading

Namespace Util.Concurrent
    ''' <summary>
    ''' 時間の単位を表すクラス
    ''' </summary>
    ''' <remarks>JavaのTimeUnitクラス相当</remarks>
    Public Class TimeUnit
        ''' <summary>ナノ秒</summary>
        Public Shared ReadOnly NANOSECONDS As New TimeUnit(New NanoSecond)
        ''' <summary>マイクロ秒</summary>
        Public Shared ReadOnly MICROSECONDS As New TimeUnit(New MicroSecond)
        ''' <summary>ミリ秒</summary>
        Public Shared ReadOnly MILLISECONDS As New TimeUnit(New MilliSecond)
        ''' <summary>秒</summary>
        Public Shared ReadOnly SECONDS As New TimeUnit(New Second)
        ''' <summary>分</summary>
        Public Shared ReadOnly MINUTES As New TimeUnit(New Minute)
        ''' <summary>時</summary>
        Public Shared ReadOnly HOURS As New TimeUnit(New Hour)
        ''' <summary>日</summary>
        Public Shared ReadOnly DAYS As New TimeUnit(New Day)

        Private ReadOnly behavior As IBehavior

        Private Sub New(ByVal behavior As IBehavior)
            Me.behavior = behavior
        End Sub

        ''' <summary>
        ''' ナノ秒単位にする
        ''' </summary>
        ''' <param name="duration">時間</param>
        ''' <returns>ナノ秒単位の時間</returns>
        ''' <remarks></remarks>
        Public Function ToNanos(ByVal duration As Long) As Long
            Return behavior.ToNanos(duration)
        End Function

        ''' <summary>
        ''' マイクロ秒単位にする
        ''' </summary>
        ''' <param name="duration">時間</param>
        ''' <returns>マイクロ秒単位の時間</returns>
        ''' <remarks></remarks>
        Public Function ToMicros(ByVal duration As Long) As Long
            Return behavior.ToMicros(duration)
        End Function

        ''' <summary>
        ''' ミリ秒単位にする
        ''' </summary>
        ''' <param name="duration">時間</param>
        ''' <returns>ミリ秒単位の時間</returns>
        ''' <remarks></remarks>
        Public Function ToMillis(ByVal duration As Long) As Long
            Return behavior.ToMillis(duration)
        End Function

        ''' <summary>
        ''' 秒単位にする
        ''' </summary>
        ''' <param name="duration">時間</param>
        ''' <returns>秒単位の時間</returns>
        ''' <remarks></remarks>
        Public Function ToSeconds(ByVal duration As Long) As Long
            Return behavior.ToSeconds(duration)
        End Function

        ''' <summary>
        ''' 分単位にする
        ''' </summary>
        ''' <param name="duration">時間</param>
        ''' <returns>分単位の時間</returns>
        ''' <remarks></remarks>
        Public Function ToMinutes(ByVal duration As Long) As Long
            Return behavior.ToMinutes(duration)
        End Function

        ''' <summary>
        ''' 時単位にする
        ''' </summary>
        ''' <param name="duration">時間</param>
        ''' <returns>時単位の時間</returns>
        ''' <remarks></remarks>
        Public Function ToHours(ByVal duration As Long) As Long
            Return behavior.ToHours(duration)
        End Function

        ''' <summary>
        ''' 日数にする
        ''' </summary>
        ''' <param name="duration">時間</param>
        ''' <returns>日数</returns>
        ''' <remarks></remarks>
        Public Function ToDays(ByVal duration As Long) As Long
            Return behavior.ToDays(duration)
        End Function

        ''' <summary>
        ''' この時間単位にする
        ''' </summary>
        ''' <param name="sourceDuration">時間</param>
        ''' <param name="sourceUnit">単位</param>
        ''' <returns>当インスタンスの単位にした時間</returns>
        ''' <remarks></remarks>
        Public Function Convert(ByVal sourceDuration As Long, ByVal sourceUnit As TimeUnit) As Long
            Return behavior.Convert(sourceDuration, sourceUnit)
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="duration"></param>
        ''' <param name="milliseconds"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Function ExcessNanos(ByVal duration As Long, ByVal milliseconds As Long) As Integer
            Return behavior.ExcessNanos(duration, milliseconds)
        End Function

        ''' <summary>
        ''' この時間単位を使用して、時間指定された Monitor.Wait を実行する
        ''' </summary>
        ''' <param name="obj">待機するObject</param>
        ''' <param name="timeout">待機する最長時間</param>
        ''' <remarks></remarks>
        Public Sub TimedWait(ByVal obj As Object, ByVal timeout As Long)
            If 0 < timeout Then
                Dim millis As Long = ToMillis(timeout)
                'Dim nanos As Integer = ExcessNanos(timeout, ms)
                Monitor.Wait(obj, ti(millis))
            End If
        End Sub

        ''' <summary>
        ''' この時間単位を使用して、時間指定された Thread.Join を実行する
        ''' </summary>
        ''' <param name="aThread">待機するThread</param>
        ''' <param name="timeout">待機する最長時間</param>
        ''' <remarks></remarks>
        Public Sub TimedJoin(ByVal aThread As Thread, ByVal timeout As Long)
            If 0 < timeout Then
                Dim millis As Long = ToMillis(timeout)
                'Dim nanos As Integer = ExcessNanos(timeout, ms)
                aThread.Join(ti(millis))
            End If
        End Sub

        ''' <summary>
        ''' この時間単位を使用して、時間指定された Thread.Sleep を実行する
        ''' </summary>
        ''' <param name="timeout">Sleepの最小時間</param>
        ''' <remarks></remarks>
        Public Sub Sleep(ByVal timeout As Long)
            If 0 < timeout Then
                Dim millis As Long = ToMillis(timeout)
                'Dim nanos As Integer = ExcessNanos(timeout, ms)
                Thread.Sleep(ti(millis))
            End If
        End Sub

        Private Const C0 As Long = 1&
        Private Const C1 As Long = C0 * 1000&
        Private Const C2 As Long = C1 * 1000&
        Private Const C3 As Long = C2 * 1000&
        Private Const C4 As Long = C3 * 60&
        Private Const C5 As Long = C4 * 60&
        Private Const C6 As Long = C5 * 24&

        Private Shared Function x(ByVal d As Long, ByVal m As Double, ByVal over As Double) As Long
            Return x(d, t(m), t(over))
        End Function

        Private Shared Function x(ByVal d As Long, ByVal m As Long, ByVal over As Long) As Long
            If over < d Then
                Return Long.MaxValue
            ElseIf d < -over Then
                Return Long.MinValue
            End If
            Return d * m
        End Function

        Private Shared Function t(ByVal d As Double) As Long
            Return CLng(d)
        End Function

        Private Shared Function ti(ByVal d As Long) As Integer
            Return CInt(d)
        End Function

        Private Interface IBehavior
            Function ToNanos(ByVal duration As Long) As Long
            Function ToMicros(ByVal duration As Long) As Long
            Function ToMillis(ByVal duration As Long) As Long
            Function ToSeconds(ByVal duration As Long) As Long
            Function ToMinutes(ByVal duration As Long) As Long
            Function ToHours(ByVal duration As Long) As Long
            Function ToDays(ByVal duration As Long) As Long
            Function Convert(ByVal sourceDuration As Long, ByVal sourceUnit As TimeUnit) As Long
            Function ExcessNanos(ByVal d As Long, ByVal m As Long) As Integer
        End Interface

        Private Class NanoSecond : Implements IBehavior

            Public Function ToNanos(ByVal duration As Long) As Long Implements IBehavior.ToNanos
                Return duration
            End Function

            Public Function ToMicros(ByVal duration As Long) As Long Implements IBehavior.ToMicros
                Return t(duration / (C1 / C0))
            End Function

            Public Function ToMillis(ByVal duration As Long) As Long Implements IBehavior.ToMillis
                Return t(duration / (C2 / C0))
            End Function

            Public Function ToSeconds(ByVal duration As Long) As Long Implements IBehavior.ToSeconds
                Return t(duration / (C3 / C0))
            End Function

            Public Function ToMinutes(ByVal duration As Long) As Long Implements IBehavior.ToMinutes
                Return t(duration / (C4 / C0))
            End Function

            Public Function ToHours(ByVal duration As Long) As Long Implements IBehavior.ToHours
                Return t(duration / (C5 / C0))
            End Function

            Public Function ToDays(ByVal duration As Long) As Long Implements IBehavior.ToDays
                Return t(duration / (C6 / C0))
            End Function

            Public Function Convert(ByVal sourceDuration As Long, ByVal sourceUnit As TimeUnit) As Long Implements IBehavior.Convert
                Return sourceUnit.ToNanos(sourceDuration)
            End Function

            Public Function ExcessNanos(ByVal d As Long, ByVal m As Long) As Integer Implements IBehavior.ExcessNanos
                Return ti(d - (m * C2))
            End Function
        End Class

        Private Class MicroSecond : Implements IBehavior

            Public Function ToNanos(ByVal duration As Long) As Long Implements IBehavior.ToNanos
                Return x(duration, C1 / C0, Long.MaxValue / (C1 / C0))
            End Function

            Public Function ToMicros(ByVal duration As Long) As Long Implements IBehavior.ToMicros
                Return duration
            End Function

            Public Function ToMillis(ByVal duration As Long) As Long Implements IBehavior.ToMillis
                Return t(duration / (C2 / C1))
            End Function

            Public Function ToSeconds(ByVal duration As Long) As Long Implements IBehavior.ToSeconds
                Return t(duration / (C3 / C1))
            End Function

            Public Function ToMinutes(ByVal duration As Long) As Long Implements IBehavior.ToMinutes
                Return t(duration / (C4 / C1))
            End Function

            Public Function ToHours(ByVal duration As Long) As Long Implements IBehavior.ToHours
                Return t(duration / (C5 / C1))
            End Function

            Public Function ToDays(ByVal duration As Long) As Long Implements IBehavior.ToDays
                Return t(duration / (C6 / C1))
            End Function

            Public Function Convert(ByVal sourceDuration As Long, ByVal sourceUnit As TimeUnit) As Long Implements IBehavior.Convert
                Return sourceUnit.ToMicros(sourceDuration)
            End Function

            Public Function ExcessNanos(ByVal d As Long, ByVal m As Long) As Integer Implements IBehavior.ExcessNanos
                Return ti((d * C1) - (m * C2))
            End Function
        End Class

        Private Class MilliSecond : Implements IBehavior

            Public Function ToNanos(ByVal duration As Long) As Long Implements IBehavior.ToNanos
                Return x(duration, C2 / C0, Long.MaxValue / (C2 / C0))
            End Function

            Public Function ToMicros(ByVal duration As Long) As Long Implements IBehavior.ToMicros
                Return x(duration, C2 / C1, Long.MaxValue / (C2 / C1))
            End Function

            Public Function ToMillis(ByVal duration As Long) As Long Implements IBehavior.ToMillis
                Return duration
            End Function

            Public Function ToSeconds(ByVal duration As Long) As Long Implements IBehavior.ToSeconds
                Return t(duration / (C3 / C2))
            End Function

            Public Function ToMinutes(ByVal duration As Long) As Long Implements IBehavior.ToMinutes
                Return t(duration / (C4 / C2))
            End Function

            Public Function ToHours(ByVal duration As Long) As Long Implements IBehavior.ToHours
                Return t(duration / (C5 / C2))
            End Function

            Public Function ToDays(ByVal duration As Long) As Long Implements IBehavior.ToDays
                Return t(duration / (C6 / C2))
            End Function

            Public Function Convert(ByVal sourceDuration As Long, ByVal sourceUnit As TimeUnit) As Long Implements IBehavior.Convert
                Return sourceUnit.ToMillis(sourceDuration)
            End Function

            Public Function ExcessNanos(ByVal d As Long, ByVal m As Long) As Integer Implements IBehavior.ExcessNanos
                Return 0
            End Function
        End Class

        Private Class Second : Implements IBehavior

            Public Function ToNanos(ByVal duration As Long) As Long Implements IBehavior.ToNanos
                Return x(duration, C3 / C0, Long.MaxValue / (C3 / C0))
            End Function

            Public Function ToMicros(ByVal duration As Long) As Long Implements IBehavior.ToMicros
                Return x(duration, C3 / C1, Long.MaxValue / (C3 / C1))
            End Function

            Public Function ToMillis(ByVal duration As Long) As Long Implements IBehavior.ToMillis
                Return x(duration, C3 / C2, Long.MaxValue / (C3 / C2))
            End Function

            Public Function ToSeconds(ByVal duration As Long) As Long Implements IBehavior.ToSeconds
                Return duration
            End Function

            Public Function ToMinutes(ByVal duration As Long) As Long Implements IBehavior.ToMinutes
                Return t(duration / (C4 / C3))
            End Function

            Public Function ToHours(ByVal duration As Long) As Long Implements IBehavior.ToHours
                Return t(duration / (C5 / C3))
            End Function

            Public Function ToDays(ByVal duration As Long) As Long Implements IBehavior.ToDays
                Return t(duration / (C6 / C3))
            End Function

            Public Function Convert(ByVal sourceDuration As Long, ByVal sourceUnit As TimeUnit) As Long Implements IBehavior.Convert
                Return sourceUnit.ToSeconds(sourceDuration)
            End Function

            Public Function ExcessNanos(ByVal d As Long, ByVal m As Long) As Integer Implements IBehavior.ExcessNanos
                Return 0
            End Function
        End Class

        Private Class Minute : Implements IBehavior

            Public Function ToNanos(ByVal duration As Long) As Long Implements IBehavior.ToNanos
                Return x(duration, C4 / C0, Long.MaxValue / (C4 / C0))
            End Function

            Public Function ToMicros(ByVal duration As Long) As Long Implements IBehavior.ToMicros
                Return x(duration, C4 / C1, Long.MaxValue / (C4 / C1))
            End Function

            Public Function ToMillis(ByVal duration As Long) As Long Implements IBehavior.ToMillis
                Return x(duration, C4 / C2, Long.MaxValue / (C4 / C2))
            End Function

            Public Function ToSeconds(ByVal duration As Long) As Long Implements IBehavior.ToSeconds
                Return x(duration, C4 / C3, Long.MaxValue / (C4 / C3))
            End Function

            Public Function ToMinutes(ByVal duration As Long) As Long Implements IBehavior.ToMinutes
                Return duration
            End Function

            Public Function ToHours(ByVal duration As Long) As Long Implements IBehavior.ToHours
                Return t(duration / (C5 / C4))
            End Function

            Public Function ToDays(ByVal duration As Long) As Long Implements IBehavior.ToDays
                Return t(duration / (C6 / C4))
            End Function

            Public Function Convert(ByVal sourceDuration As Long, ByVal sourceUnit As TimeUnit) As Long Implements IBehavior.Convert
                Return sourceUnit.ToMinutes(sourceDuration)
            End Function

            Public Function ExcessNanos(ByVal d As Long, ByVal m As Long) As Integer Implements IBehavior.ExcessNanos
                Return 0
            End Function
        End Class

        Private Class Hour : Implements IBehavior

            Public Function ToNanos(ByVal duration As Long) As Long Implements IBehavior.ToNanos
                Return x(duration, C5 / C0, Long.MaxValue / (C5 / C0))
            End Function

            Public Function ToMicros(ByVal duration As Long) As Long Implements IBehavior.ToMicros
                Return x(duration, C5 / C1, Long.MaxValue / (C5 / C1))
            End Function

            Public Function ToMillis(ByVal duration As Long) As Long Implements IBehavior.ToMillis
                Return x(duration, C5 / C2, Long.MaxValue / (C5 / C2))
            End Function

            Public Function ToSeconds(ByVal duration As Long) As Long Implements IBehavior.ToSeconds
                Return x(duration, C5 / C3, Long.MaxValue / (C5 / C3))
            End Function

            Public Function ToMinutes(ByVal duration As Long) As Long Implements IBehavior.ToMinutes
                Return x(duration, C5 / C4, Long.MaxValue / (C5 / C4))
            End Function

            Public Function ToHours(ByVal duration As Long) As Long Implements IBehavior.ToHours
                Return duration
            End Function

            Public Function ToDays(ByVal duration As Long) As Long Implements IBehavior.ToDays
                Return t(duration / (C6 / C5))
            End Function

            Public Function Convert(ByVal sourceDuration As Long, ByVal sourceUnit As TimeUnit) As Long Implements IBehavior.Convert
                Return sourceUnit.ToHours(sourceDuration)
            End Function

            Public Function ExcessNanos(ByVal d As Long, ByVal m As Long) As Integer Implements IBehavior.ExcessNanos
                Return 0
            End Function
        End Class

        Private Class Day : Implements IBehavior

            Public Function ToNanos(ByVal duration As Long) As Long Implements IBehavior.ToNanos
                Return x(duration, C6 / C0, Long.MaxValue / (C6 / C0))
            End Function

            Public Function ToMicros(ByVal duration As Long) As Long Implements IBehavior.ToMicros
                Return x(duration, C6 / C1, Long.MaxValue / (C6 / C1))
            End Function

            Public Function ToMillis(ByVal duration As Long) As Long Implements IBehavior.ToMillis
                Return x(duration, C6 / C2, Long.MaxValue / (C6 / C2))
            End Function

            Public Function ToSeconds(ByVal duration As Long) As Long Implements IBehavior.ToSeconds
                Return x(duration, C6 / C3, Long.MaxValue / (C6 / C3))
            End Function

            Public Function ToMinutes(ByVal duration As Long) As Long Implements IBehavior.ToMinutes
                Return x(duration, C6 / C4, Long.MaxValue / (C6 / C4))
            End Function

            Public Function ToHours(ByVal duration As Long) As Long Implements IBehavior.ToHours
                Return x(duration, C6 / C5, Long.MaxValue / (C6 / C5))
            End Function

            Public Function ToDays(ByVal duration As Long) As Long Implements IBehavior.ToDays
                Return duration
            End Function

            Public Function Convert(ByVal sourceDuration As Long, ByVal sourceUnit As TimeUnit) As Long Implements IBehavior.Convert
                Return sourceUnit.ToDays(sourceDuration)
            End Function

            Public Function ExcessNanos(ByVal d As Long, ByVal m As Long) As Integer Implements IBehavior.ExcessNanos
                Return 0
            End Function
        End Class

    End Class
End Namespace