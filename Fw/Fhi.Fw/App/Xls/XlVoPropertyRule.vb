Namespace App.Xls
    ''' <summary>
    ''' Excel⇔Vo相互変換時の1プロパティに対するルール
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class XlVoPropertyRule
        ''' <summary>Vo変換時の装飾処理</summary>
        Public ReadOnly Property ToVoDecolator As Func(Of Object, Object)
            Get
                Return _ToVoDecolator
            End Get
        End Property

        ''' <summary>Excel出力時の装飾処理</summary>
        Public ReadOnly Property ToXlsDecolator As Func(Of Object, Object)
            Get
                Return _ToXlsDecolator
            End Get
        End Property

        ''' <summary>プロパティ名</summary>
        Public ReadOnly Property PropertyName As String
            Get
                Return _PropertyName
            End Get
        End Property

        Private ReadOnly _ToVoDecolator As Func(Of Object, Object)
        Private ReadOnly _ToXlsDecolator As Func(Of Object, Object)
        Private ReadOnly _PropertyName As String

        Public Sub New(propertyName As String, Optional toVoDecolator As Func(Of Object, Object) = Nothing, Optional toXlsDecolator As Func(Of Object, Object) = Nothing)
            _PropertyName = propertyName
            _ToVoDecolator = toVoDecolator
            _ToXlsDecolator = toXlsDecolator
        End Sub
    End Class
End Namespace