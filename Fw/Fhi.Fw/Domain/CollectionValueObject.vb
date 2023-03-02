Namespace Domain
    ''' <summary>
    ''' 不変のコレクションオブジェクトを担うクラス
    ''' </summary>
    ''' <typeparam name="T">値オブジェクト</typeparam>
    ''' <remarks></remarks>
    Public Class CollectionValueObject(Of T As ValueObject) : Inherits CollectionObject(Of T)

        Public Sub New()
        End Sub

        Public Sub New(ByVal initialList As IEnumerable(Of T))
            MyBase.New(initialList)
        End Sub

        Public Sub New(ByVal src As CollectionObject(Of T))
            MyBase.New(src)
        End Sub

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj Is Nothing OrElse Not (obj.GetType Is Me.GetType) Then
                Return False
            End If
            Dim other As CollectionValueObject(Of T) = DirectCast(obj, CollectionValueObject(Of T))
            If Me.Count <> other.Count Then
                Return False
            End If
            Return Enumerable.Range(0, Count).All(Function(i) Item(i).Equals(other.Item(i)))
        End Function

        Public Overrides Function GetHashCode() As Integer
            If CollectionUtil.IsEmpty(InternalList) Then
                Return 0
            End If
            Return InternalList.Select(Function(x) If(x IsNot Nothing, x.GetHashCode, 0)).Aggregate(Function(x, y) x Xor y)
        End Function

        ''' <summary>
        ''' 新しい値オブジェクト型のコレクションオブジェクトにする
        ''' </summary>
        ''' <typeparam name="TResult">新しい値オブジェクト型</typeparam>
        ''' <param name="selector"></param>
        ''' <returns>新しい値オブジェクト型のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Public Function [Select](Of TResult As ValueObject)(ByVal selector As Func(Of T, TResult)) As CollectionValueObject(Of TResult)
            Return New CollectionValueObject(Of TResult)(InternalList.Select(selector))
        End Function
        ''' <summary>
        ''' 新しい値オブジェクト型のコレクションオブジェクトにする
        ''' </summary>
        ''' <typeparam name="TResult">新しい値オブジェクト型</typeparam>
        ''' <param name="selector"></param>
        ''' <returns>新しい値オブジェクト型のコレクションオブジェクト</returns>
        ''' <remarks></remarks>
        Public Function [Select](Of TResult As ValueObject)(ByVal selector As Func(Of T, Integer, TResult)) As CollectionValueObject(Of TResult)
            Return New CollectionValueObject(Of TResult)(InternalList.Select(selector))
        End Function

    End Class
End Namespace