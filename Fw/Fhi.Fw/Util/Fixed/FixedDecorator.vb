Namespace Util.Fixed
    Public Class FixedDecorator : Implements IFixedEntry
        Private ReadOnly _Name As String
        Private ReadOnly _Length As Integer
        Private ReadOnly _repeat As Integer
        Private ReadOnly isZenkaku As Boolean
        Private ReadOnly decorateToVo As Func(Of String, Object)
        Private ReadOnly decorateToString As Func(Of Object, String)

        ''' <summary>Folder内で先頭からのoffset位置</summary>
        Private _offset As Integer

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="isZenkaku">全角の場合、true</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <param name="decorateToVo">固定長文字列からVoへ変換時の装飾処理</param>
        ''' <param name="decorateToString">Voから固定長文字列へ変換時の装飾処理</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal isZenkaku As Boolean, ByVal repeat As Integer, _
                       Optional decorateToVo As Func(Of String, Object) = Nothing, Optional decorateToString As Func(Of Object, String) = Nothing)
            Me.New(name, AbstractFixedDefine.LENGTH_AS_OTHERS, isZenkaku, repeat, decorateToVo, decorateToString)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="isZenkaku">全角の場合、true</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <param name="decorateToVo">固定長文字列からVoへ変換時の装飾処理</param>
        ''' <param name="decorateToString">Voから固定長文字列へ変換時の装飾処理</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal length As Integer, ByVal isZenkaku As Boolean, ByVal repeat As Integer, _
                       Optional decorateToVo As Func(Of String, Object) = Nothing, Optional decorateToString As Func(Of Object, String) = Nothing)
            EzUtil.AssertParameterIsNotNull(name, "name")
            Me._Name = name.ToUpper
            Me._Length = length
            Me.isZenkaku = isZenkaku
            Me._repeat = repeat
            Me.decorateToString = decorateToString
            Me.decorateToVo = decorateToVo
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
            Dim funcOfToString As Func(Of Object, String) = If(decorateToString, Function(v) StringUtil.ToString(v))

            Dim str As String = funcOfToString.Invoke(value)
            If _Length <> AbstractFixedDefine.LENGTH_AS_OTHERS Then
                str &= StringUtil.Repeat(" "c, _Length)
                str = Left(str, _Length)
            End If
            Return If(isZenkaku, StringUtil.ToZenkaku(str), StringUtil.ToHankaku(str))
        End Function

        ''' <summary>
        ''' 固定長文字列を値にする
        ''' </summary>
        ''' <param name="fixedString">固定長文字列</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Public Function Parse(ByVal fixedString As String) As Object Implements IFixedEntry.Parse
            Dim str As String = StringUtil.TrimEnd(fixedString)
            If decorateToVo IsNot Nothing Then
                Return decorateToVo.Invoke(str)
            End If
            Return str
        End Function
    End Class
End Namespace