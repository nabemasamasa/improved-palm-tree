Namespace Db
    Public Enum SampleEnum
        A = 1
        B = 2
    End Enum

    Public Class SampleVo
        Public Sub New()
        End Sub
        Public Sub New(hogeId As Integer?, hogeName As String, HogeDate As DateTime?, HogeDecimal As Decimal?)
            Me.HogeId = hogeId
            Me.HogeName = hogeName
            Me.HogeDate = HogeDate
            Me.HogeDecimal = HogeDecimal
        End Sub

        Public Property HogeId As Integer?
        Public Property HogeName As String
        Public Property HogeDate As DateTime?
        Public Property HogeDecimal As Decimal?
        Public Property HogeEnum As SampleEnum?
    End Class
End Namespace