Namespace Domain
    ''' <summary>
    ''' 一意の識別子IDをもつEntity
    ''' </summary>
    ''' <typeparam name="T">idの型</typeparam>
    ''' <remarks></remarks>
    Public MustInherit Class Entity(Of T)

        Private _id As Object
        Private _hashCode As Integer?
        ''' <summary>識別子</summary>
        Protected Overridable Property Id As T
            Get
                Return DirectCast(_id, T)
            End Get
            Set(value As T)
                If _id IsNot Nothing Then
                    If _id.Equals(value) Then
                        Return
                    End If
                    Throw New InvalidOperationException("設定済みのEntityIDは変更できない")
                ElseIf _hashCode.HasValue Then
                    Throw New InvalidOperationException("EntityID設定前にHashCode利用がある. 先にEntityIDを設定すべき")
                End If
                _id = value
                If value IsNot Nothing Then
                    _hashCode = value.GetHashCode Xor 31
                End If
            End Set
        End Property

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj Is Nothing OrElse Not (obj.GetType Is Me.GetType) Then
                Return False
            End If
            If ReferenceEquals(Me, obj) Then
                Return True
            End If
            Dim other As Entity(Of T) = DirectCast(obj, Entity(Of T))
            If Me.GetHashCode = obj.GetHashCode Then
                Return True
            End If
            If Me.Id Is Nothing AndAlso other.Id Is Nothing Then
                Return Me Is other
            End If
            Return Me.Id.Equals(other.Id)
        End Function

        Public Overrides Function GetHashCode() As Integer
            If Not _hashCode.HasValue Then
                _hashCode = MyBase.GetHashCode
            End If
            Return _hashCode.Value
        End Function

    End Class
    ''' <summary>
    ''' 一意の識別子IDをもつEntity
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class Entity : Inherits Entity(Of Object)

    End Class
End Namespace