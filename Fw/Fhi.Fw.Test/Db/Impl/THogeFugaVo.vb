Namespace Db.Impl
    Public Class THogeFugaVo
        Public Enum Piyo
            ONE = 1
            TWO = 2
        End Enum
        Private _hogeId As Nullable(Of Int32)
        Private _hogeSub As String
        Private _hogeName As String
        Private _HogeDate As Nullable(Of DateTime)
        Private _HogeDecimal As Nullable(Of Decimal)
        Private _IsHoge As Nullable(Of Boolean)
        Private _HogeEnum As Nullable(Of Piyo)
        Private _UpdatedUserId As String
        Private _UpdatedDate As String
        Private _UpdatedTime As String
        Public Property HogeId() As Nullable(Of Int32)
            Get
                Return _hogeId
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _hogeId = value
            End Set
        End Property
        Public Property HogeSub() As String
            Get
                Return _hogeSub
            End Get
            Set(ByVal value As String)
                _hogeSub = value
            End Set
        End Property
        Public Property HogeName() As String
            Get
                Return _hogeName
            End Get
            Set(ByVal value As String)
                _hogeName = value
            End Set
        End Property
        Public Property HogeDate() As Nullable(Of DateTime)
            Get
                Return _HogeDate
            End Get
            Set(ByVal value As Nullable(Of DateTime))
                _HogeDate = value
            End Set
        End Property
        Public Property HogeDecimal() As Nullable(Of Decimal)
            Get
                Return _HogeDecimal
            End Get
            Set(ByVal value As Nullable(Of Decimal))
                _HogeDecimal = value
            End Set
        End Property
        Public Property IsHoge() As Boolean?
            Get
                Return _IsHoge
            End Get
            Set(ByVal value As Boolean?)
                _IsHoge = value
            End Set
        End Property

        Public Property HogeEnum() As Piyo?
            Get
                Return _HogeEnum
            End Get
            Set(ByVal value As Piyo?)
                _HogeEnum = value
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