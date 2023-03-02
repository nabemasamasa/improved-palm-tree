Imports Fhi.Fw.Domain

Namespace Util.Fixed
    ''' <summary>
    ''' 固定長の数値を表わすクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FixedNumber : Implements IFixedEntry

        Private ReadOnly _Name As String
        Private ReadOnly _Length As Integer
        Private ReadOnly _scale As Integer
        Private ReadOnly aType As Type
        Private ReadOnly _repeat As Integer

        ''' <summary>Folder内で先頭からのoffset位置</summary>
        Private _offset As Integer

        ''' <summary>数値型がNullのときzero埋めするか？</summary>
        Public IsZeroPaddingIfNull As Boolean

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="aType">型Type</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="scale">文字数のうち小数桁数</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal aType As Type, ByVal length As Integer, ByVal scale As Integer)
            Me.New(name, aType, length, scale, 1)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="name">属性名</param>
        ''' <param name="aType">型Type</param>
        ''' <param name="length">桁数（文字数）</param>
        ''' <param name="scale">文字数のうち小数桁数</param>
        ''' <param name="repeat">繰返し数</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal aType As Type, ByVal length As Integer, ByVal scale As Integer, ByVal repeat As Integer)
            EzUtil.AssertParameterIsNotNull(name, "name")
            Me._Name = name.ToUpper
            Me._Length = length
            Me._scale = scale
            Me._repeat = repeat
            Me.aType = aType
        End Sub

        ''' <summary>Group内の開始位置offset</summary>
        Public ReadOnly Property Offset() As Integer Implements IFixedEntry.Offset
            Get
                Return _offset
            End Get
        End Property

        ''' <summary>固定長の長さ（Groupの場合、内包する長さ）</summary>
        Public ReadOnly Property Length() As Integer Implements IFixedEntry.Length
            Get
                Return _Length
            End Get
        End Property

        ''' <summary>名前</summary>
        Public ReadOnly Property Name() As String Implements IFixedEntry.Name
            Get
                Return _Name
            End Get
        End Property

        ''' <summary>繰り返し数</summary>
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
            If value Is Nothing Then
                If IsZeroPaddingIfNull Then
                    Return ZeroPadding(value)
                Else
                    Return StringUtil.Repeat(" "c, _Length)
                End If
            End If
            If TypeOf value Is PrimitiveValueObject Then
                Return Format(value.ToString())
            End If
            If Not IsNumeric(value) Then
                Throw New ArgumentException(String.Format("{0} は数値型（桁数{1}、うち小数{2}桁）だが、数値以外が指定された value='{3}'", _Name, _Length, _scale, value), "value")
            End If
            If _scale = 0 Then
                Return ZeroPadding(value)
            End If
            Return ZeroPadding(Decimal.Truncate(Decimal.Multiply(CDec(value), CDec(10 ^ _scale))))
        End Function

        Private Function ZeroPadding(ByVal value As Object) As String
            If value Is Nothing Then
                Return StringUtil.Repeat("0"c, _Length)
            End If
            Return Right(StringUtil.Repeat("0"c, _Length) & StringUtil.ToString(value), _Length)
        End Function

        ''' <summary>
        ''' 固定長文字列を値にする
        ''' </summary>
        ''' <param name="fixedString">固定長文字列</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Public Function Parse(ByVal fixedString As String) As Object Implements IFixedEntry.Parse
            If StringUtil.IsEmpty(fixedString) Then
                If IsZeroPaddingIfNull Then
                    Return 0
                Else
                    Return Nothing
                End If
            End If
            If Not IsNumeric(fixedString) Then
                Throw New ArgumentException("値は数値文字列であるべき", "fixedString")
            End If

            Dim t As Type = If(aType.IsArray, aType.GetElementType, If(TypeUtil.IsTypeCollection(aType), TypeUtil.DetectElementType(aType), aType))
            Return VoUtil.ResolveValue(Decimal.Divide(CDec(fixedString), CDec(10 ^ _scale)), t)
        End Function
    End Class
End Namespace