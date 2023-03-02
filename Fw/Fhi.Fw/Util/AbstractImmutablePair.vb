Namespace Util
    ''' <summary>
    ''' ペアでの不変オブジェクト
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class AbstractImmutablePair(Of TA, TB) : Implements IEquatable(Of AbstractImmutablePair(Of TA, TB))
        ''' <summary>ペア値A</summary>
        Protected ReadOnly PairA As TA
        ''' <summary>ペア値B</summary>
        Protected ReadOnly PairB As TB

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="pairA">ペア値A</param>
        ''' <param name="pairB">ペア値B</param>
        ''' <remarks></remarks>
        Protected Sub New(ByVal pairA As TA, ByVal pairB As TB)
            Me.PairA = pairA
            Me.PairB = pairB
        End Sub

        Public Overrides Function Equals(ByVal obj As Object) As Boolean
            If obj Is Nothing OrElse Not obj.GetType Is Me.GetType Then
                Return False
            End If
            Return PerformEquals(DirectCast(obj, AbstractImmutablePair(Of TA, TB)))
        End Function

        Public Overloads Function Equals(ByVal other As AbstractImmutablePair(Of TA, TB)) As Boolean Implements IEquatable(Of AbstractImmutablePair(Of TA, TB)).Equals
            If other Is Nothing Then
                Return False
            End If
            Return PerformEquals(other)
        End Function

        Private Function PerformEquals(ByVal other As AbstractImmutablePair(Of TA, TB)) As Boolean

            Return EzUtil.IsEqualIfNull(Me.PairA, other.PairA) AndAlso EzUtil.IsEqualIfNull(Me.PairB, other.PairB)
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return If(PairA Is Nothing, 0, PairA.GetHashCode) Xor If(PairB Is Nothing, 0, PairB.GetHashCode)
        End Function

    End Class
End Namespace