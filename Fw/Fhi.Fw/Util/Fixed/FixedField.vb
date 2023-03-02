Namespace Util.Fixed
    ''' <summary>
    ''' 固定長の文字列を表わすクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FixedField : Implements IFixedEntry

        Private ReadOnly _Name As String
        Private ReadOnly _Length As Integer
        Private ReadOnly _repeat As Integer
        Private ReadOnly isZenkaku As Boolean

        ''' <summary>Folder内で先頭からのoffset位置</summary>
        Private _offset As Integer

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="isZenkaku">全角の場合、true</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal length As Integer, ByVal isZenkaku As Boolean)
            Me.New(name, length, isZenkaku, 1)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="isZenkaku">全角の場合、true</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal length As Integer, ByVal isZenkaku As Boolean, ByVal repeat As Integer)
            EzUtil.AssertParameterIsNotNull(name, "name")
            Me._Name = name.ToUpper
            Me._Length = length
            Me.isZenkaku = isZenkaku
            Me._repeat = repeat
        End Sub

        Public ReadOnly Property Offset() As Integer Implements IFixedEntry.Offset
            Get
                Return _offset
            End Get
        End Property

        Public ReadOnly Property Length() As Integer Implements IFixedEntry.Length
            Get
                Return _Length
            End Get
        End Property

        Public ReadOnly Property Name() As String Implements IFixedEntry.Name
            Get
                Return _Name
            End Get
        End Property

        Public ReadOnly Property Repeat() As Integer Implements IFixedEntry.Repeat
            Get
                Return _repeat
            End Get
        End Property

        Public Function ContainsName(ByVal childName As String) As Boolean Implements IFixedEntry.ContainsName
            Return False
        End Function

        Public Function GetChlid(ByVal childName As String) As IFixedEntry Implements IFixedEntry.GetChlid
            Return Nothing
        End Function

        Public Sub InitializeOffset(ByVal offset As Integer) Implements IFixedEntry.InitializeOffset
            Me._offset = offset
        End Sub

        ''' <summary>
        ''' 値を固定長文字列にする
        ''' </summary>
        ''' <param name="value">値</param>
        ''' <returns>固定長文字列</returns>
        ''' <remarks></remarks>
        Public Function Format(ByVal value As Object) As String Implements IFixedEntry.Format
            Dim str As String = StringUtil.ToString(value) & StringUtil.Repeat(" "c, _Length)
            Return Left(If(isZenkaku, StringUtil.ToZenkaku(str), StringUtil.ToHankaku(str)), _Length)
        End Function

        ''' <summary>
        ''' 固定長文字列を値にする
        ''' </summary>
        ''' <param name="fixedString">固定長文字列</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Public Function Parse(ByVal fixedString As String) As Object Implements IFixedEntry.Parse
            If fixedString Is Nothing Then
                Return Nothing
            End If
            Return fixedString.TrimEnd
        End Function
    End Class
End Namespace