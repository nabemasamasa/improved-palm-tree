Imports System.Text

Namespace Db.Sql
    ''' <summary>
    ''' XML定義に習い、不等号を、CDATA で囲むクラス
    ''' </summary>
    ''' <remarks>但し、指定のXMLタグを許容するため、不等号全てが CDATA で囲まれるわけではない</remarks>
    Public Class XmlCDataEncloser
        Private Const XML_TAG_NAME_NEXT_CHARS As String = " >" & vbCrLf & vbTab
        ' 許容するタグ名
        Private allowTagNames As String()

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="allowTagNames">許容するタグ名</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal ParamArray allowTagNames As String())
            Me.allowTagNames = allowTagNames
        End Sub

        ''' <summary>
        ''' タグ名から始まっているかを返す
        ''' </summary>
        ''' <param name="str">判定文字列</param>
        ''' <returns>始まっている場合、true</returns>
        ''' <remarks></remarks>
        Private Function StartsWithTag(ByVal str As String) As Boolean
            For Each tagName As String In allowTagNames
                If tagName.Length < str.Length AndAlso str.ToLower.StartsWith(tagName.ToLower) AndAlso IsTagNameNextChar(str.Chars(tagName.Length)) Then
                    Return True
                End If
                Dim endTagName As String = "/" & tagName
                If endTagName.Length < str.Length AndAlso str.ToLower.StartsWith(endTagName.ToLower) AndAlso IsTagNameNextChar(str.Chars(endTagName.Length)) Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' タグ名の次に来るべき文字かを返す
        ''' </summary>
        ''' <param name="c">判定文字</param>
        ''' <returns>次の文字なら、true</returns>
        ''' <remarks></remarks>
        Private Function IsTagNameNextChar(ByVal c As Char) As Boolean
            Return 0 <= XML_TAG_NAME_NEXT_CHARS.IndexOf(c)
        End Function

        ''' <summary>
        ''' タグとは関係の無い不等号を、&lt;![CDATA[&lt;=]]&gt; な感じにする
        ''' </summary>
        ''' <param name="sql">SQL</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Enclose(ByVal sql As String) As String
            Dim chars As Char() = sql.ToCharArray
            Dim result As New StringBuilder
            Dim dqing As Boolean
            Dim sqing As Boolean
            Dim bracket As Integer = 0
            For i As Integer = 0 To chars.Length - 1
                Dim c As Char = chars(i)
                If c = """"c AndAlso Not sqing Then
                    dqing = Not dqing
                ElseIf c = "'"c AndAlso Not dqing Then
                    sqing = Not sqing
                ElseIf Not dqing AndAlso Not sqing AndAlso 0 <= "<>".IndexOf(c) Then
                    If c = "<"c Then
                        If StartsWithTag(sql.Substring(i + 1)) Then
                            bracket += 1
                            If 1 < bracket Then
                                Throw New ArgumentException("開いたままのタグがあるのに、またタグを開いています. " & sql.Substring(i))
                            End If
                        Else
                            result.Append(Enclose(chars, i))
                            i += If(IsNextAPartOfSign(chars, i), 1, 0)
                            Continue For
                        End If
                    Else ' ">"
                        If 0 < bracket Then
                            bracket -= 1
                        Else
                            result.Append(Enclose(chars, i))
                            i += If(IsNextAPartOfSign(chars, i), 1, 0)
                            Continue For
                        End If
                    End If
                End If
                result.Append(c)
            Next
            Return result.ToString
        End Function

        ''' <summary>
        ''' 不等号の次の文字が不等号の一部かを返す
        ''' </summary>
        ''' <param name="chars">SQLのchar配列</param>
        ''' <param name="index">不等号の位置index</param>
        ''' <returns>一部なら、true</returns>
        ''' <remarks></remarks>
        Private Function IsNextAPartOfSign(ByVal chars As Char(), ByVal index As Integer) As Boolean
            Return index + 1 < chars.Length - 1 AndAlso (chars(index + 1) = "="c OrElse "<>".Equals(New String(chars, index, 2)))
        End Function

        ''' <summary>
        ''' 不等号をCDATAで囲んで返す
        ''' </summary>
        ''' <param name="chars">SQLのchar配列</param>
        ''' <param name="index">不等号の位置index</param>
        ''' <returns>CDATAで囲んだ不等号</returns>
        ''' <remarks></remarks>
        Private Function Enclose(ByVal chars As Char(), ByVal index As Integer) As String
            Dim sign As String = chars(index)
            If IsNextAPartOfSign(chars, index) Then
                sign &= chars(index + 1)
            End If
            Return String.Format("<![CDATA[{0}]]>", sign)
        End Function
    End Class
End Namespace