Namespace Db.Sql
    Class HogeVo
        Private _id As Nullable(Of Integer)
        Private _name As String
        Private _last As Nullable(Of DateTime)
        Private _exName7 As String
        Public Property Id() As Nullable(Of Integer)
            Get
                Return _id
            End Get
            Set(ByVal value As Nullable(Of Integer))
                _id = value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(ByVal value As String)
                _name = value
            End Set
        End Property
        Public Property Last() As Nullable(Of DateTime)
            Get
                Return _last
            End Get
            Set(ByVal value As Nullable(Of DateTime))
                _last = value
            End Set
        End Property
        Public Property ExName7() As String
            Get
                Return _exName7
            End Get
            Set(ByVal value As String)
                _exName7 = value
            End Set
        End Property

    End Class
End Namespace