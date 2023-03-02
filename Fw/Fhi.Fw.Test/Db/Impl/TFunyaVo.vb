Namespace Db.Impl
    Public Class TFunyaVo
        Private _FunyaId As Nullable(Of Long)
        Private _FunyaName As String
        Private _UpdatedUserId As String
        Private _UpdatedDate As String
        Private _UpdatedTime As String
        Public Property FunyaId() As Nullable(Of Long)
            Get
                Return _FunyaId
            End Get
            Set(ByVal value As Nullable(Of Long))
                _FunyaId = value
            End Set
        End Property
        Public Property FunyaName() As String
            Get
                Return _FunyaName
            End Get
            Set(ByVal value As String)
                _FunyaName = value
            End Set
        End Property
        Public Property UpdatedUserId() As String
            Get
                Return _UpdatedUserId
            End Get
            Set(ByVal value As String)
                _UpdatedUserId = value
            End Set
        End Property
        Public Property UpdatedDate() As String
            Get
                Return _UpdatedDate
            End Get
            Set(ByVal value As String)
                _UpdatedDate = value
            End Set
        End Property
        Public Property UpdatedTime() As String
            Get
                Return _UpdatedTime
            End Get
            Set(ByVal value As String)
                _UpdatedTime = value
            End Set
        End Property
    End Class
End Namespace