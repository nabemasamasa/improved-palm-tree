Namespace Util.LabelValue
    Public Class FakeSampleVo
        Private _hogeId As Nullable(Of Int32)
        Private _hogeName As String
        Private _HogeDate As Nullable(Of DateTime)
        Private _HogeDecimal As Nullable(Of Decimal)
        Public Sub New()
        End Sub
        Public Sub New(ByVal hogeId As Nullable(Of Int32), ByVal hogeName As String, ByVal HogeDate As Nullable(Of DateTime), ByVal HogeDecimal As Nullable(Of Decimal))
            _hogeId = hogeId
            _hogeName = hogeName
            _HogeDate = HogeDate
            _HogeDecimal = HogeDecimal
        End Sub

        Public Property HogeId() As Nullable(Of Int32)
            Get
                Return _hogeId
            End Get
            Set(ByVal value As Nullable(Of Int32))
                _hogeId = value
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

    End Class
End Namespace