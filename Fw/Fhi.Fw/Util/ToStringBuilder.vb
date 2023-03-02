Imports System.Text

Namespace Util

    ''' <summary>
    ''' ToStringの文字列表示を作成するための簡易クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ToStringBuilder
        Private _result As New List(Of String)

        ''' <summary>
        ''' ToString表示用の情報を追加する
        ''' </summary>
        ''' <param name="name">値の名前</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Sub Add(ByVal name As String, ByVal value As String)
            _result.Add(BuildToString(name, value))
        End Sub

        ''' <summary>
        ''' ToString表示用の情報を追加する
        ''' </summary>
        ''' <param name="name">値の名前</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Sub Add(Of T As Structure)(ByVal name As String, ByVal value As Nullable(Of T))
            _result.Add(BuildToString(name, value))
        End Sub

        ''' <summary>
        ''' ToString表示用の情報を追加する
        ''' </summary>
        ''' <param name="name">値の名前</param>
        ''' <param name="obj">値</param>
        ''' <remarks></remarks>
        Public Sub Add(ByVal name As String, ByVal obj As Object)

            _result.Add(BuildToString(name, obj))
        End Sub

        ''' <summary>
        ''' ToString表示用の情報を作成する
        ''' </summary>
        ''' <param name="name">値の名前</param>
        ''' <param name="value">値</param>
        ''' <returns>（文字列）情報</returns>
        ''' <remarks></remarks>
        Public Shared Function BuildToString(Of T As Structure)(ByVal name As String, ByVal value As Nullable(Of T)) As String
            Dim sb As New StringBuilder
            sb.Append(name).Append("="c)
            If value Is Nothing Then
                sb.Append("<nothing>")
            Else
                sb.Append(value)
            End If
            Return sb.ToString
        End Function

        ''' <summary>
        ''' ToString表示用の情報を作成する
        ''' </summary>
        ''' <param name="name">値の名前</param>
        ''' <param name="value">値</param>
        ''' <returns>（文字列）情報</returns>
        ''' <remarks></remarks>
        Public Shared Function BuildToString(ByVal name As String, ByVal value As Object) As String
            Dim sb As New StringBuilder
            sb.Append(name).Append("="c)
            If TypeOf value Is ICollection Then
                Dim index As Integer = 0
                Dim result As New List(Of String)
                For Each obj As Object In DirectCast(value, ICollection)
                    result.Add(BuildToString(String.Format("({0})", index), obj))
                    index += 1
                Next
                sb.Append("{"c).Append(Join(result.ToArray, ", ")).Append("}"c)
            Else
                sb.Append(ValueToString(value))
            End If
            Return sb.ToString
        End Function

        Private Shared Function ValueToString(ByVal value As Object) As String
            Dim sb As New StringBuilder
            If value Is Nothing Then
                sb.Append("<nothing>")
            ElseIf TypeOf value Is String Then
                sb.Append("'"c).Append(value).Append("'"c)
            ElseIf TypeUtil.GetTypeIfNullable(value.GetType).IsValueType Then
                sb.Append(value)
            Else
                sb.Append("{"c).Append(value.ToString).Append("}"c)
            End If
            Return sb.ToString
        End Function

        Public ReadOnly Property Result() As String
            Get
                Return Join(_result.ToArray, ", ")
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return Result
        End Function

        ''' <summary>
        ''' ToStringBuilderで作成した文字列の書式を整形して返す
        ''' </summary>
        ''' <param name="toString">ToStringBuilderで作成した文字列</param>
        ''' <returns>整形した文字列</returns>
        ''' <remarks></remarks>
        Public Shared Function Format(ByVal toString As String) As String
            Dim result As New StringBuilder
            Dim indent As Integer = 0

            Dim toCharArray As Char() = toString.ToCharArray()
            Dim enterFlag As Boolean
            For i As Integer = 0 To toCharArray.Length - 1
                Dim c As Char = toCharArray(i)
                If c = "{"c Then
                    If i < toCharArray.Length AndAlso toCharArray(i + 1) = "}"c Then
                        i += 1
                        result.Append("{}")
                    Else
                        indent += 1
                        result.Append(c).Append(vbCrLf).Append(Repeat(vbTab, indent))
                        enterFlag = True
                        Continue For
                    End If
                ElseIf c = "}"c Then
                    indent -= 1
                    result.Append(vbCrLf).Append(Repeat(vbTab, indent)).Append(c)
                ElseIf c = ","c Then
                    result.Append(c).Append(vbCrLf).Append(Repeat(vbTab, indent))
                    enterFlag = True
                    Continue For
                Else
                    If enterFlag AndAlso IsWhitespace(c) Then
                        Continue For
                    End If
                    result.Append(c)
                End If
                enterFlag = False
            Next
            Return result.ToString
        End Function
        Private Const WHITESPACE As String = vbCrLf & vbTab & " "
        Private Shared Function IsWhitespace(ByVal c As Char) As Boolean
            Return 0 <= WHITESPACE.IndexOf(c)
        End Function
        Private Shared Function Repeat(ByVal c As String, ByVal count As Integer) As String
            Dim result As New StringBuilder
            For i As Integer = 0 To count - 1
                result.Append(c)
            Next
            Return result.ToString
        End Function
    End Class
End Namespace