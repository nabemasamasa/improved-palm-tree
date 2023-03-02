Namespace Db.Impl.LinkServer
    ''' <summary>
    ''' IBMAS13_CYJTESTF にあるテストテーブル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Testpf01Vo
        Private _fld001 As String
        Private _fld002 As String

        Public Property Fld001() As String
            Get
                Return _fld001
            End Get
            Set(ByVal value As String)
                _fld001 = value
            End Set
        End Property

        Public Property Fld002() As String
            Get
                Return _fld002
            End Get
            Set(ByVal value As String)
                _fld002 = value
            End Set
        End Property
    End Class
End Namespace