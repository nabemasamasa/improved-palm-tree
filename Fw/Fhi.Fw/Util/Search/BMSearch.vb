Namespace Util.Search
    ''' <summary>
    ''' Boyer-Moore法による文字列検索処理
    ''' </summary>
    ''' <remarks></remarks>
    Public Class BMSearch
        ''' <summary>
        ''' Skipテーブルを作成する（文字が末尾から数えて何番目か、のテーブル）
        ''' </summary>
        ''' <param name="pattern">検索文字列</param>
        ''' <returns>Skipテーブル</returns>
        ''' <remarks>末尾の文字を0番目とするが、0番目の文字はテーブルに含まれない</remarks>
        Public Shared Function BuildSkipTable(ByVal pattern As String) As Dictionary(Of Char, Integer)
            Dim result As New Dictionary(Of Char, Integer)
            Dim plen As Integer = pattern.Length
            For i As Integer = 0 To plen - 2
                result.Add(pattern.Chars(i), plen - i - 1)
            Next
            Return result
        End Function

        ''' <summary>
        ''' 文字列検索を行う
        ''' </summary>
        ''' <param name="text">文字列</param>
        ''' <param name="pattern">検索文字列</param>
        ''' <returns>一致箇所の数</returns>
        ''' <remarks></remarks>
        Public Shared Function Search(ByVal text As String, ByVal pattern As String) As Integer
            Dim skipTable As Dictionary(Of Char, Integer) = BuildSkipTable(pattern)
            Return PerformSearch(text, pattern, skipTable)
        End Function

        Private Shared Function PerformSearch(ByVal text As String, ByVal pattern As String, ByVal skipTable As Dictionary(Of Char, Integer)) As Integer
            Dim textIndex As Integer = 0
            Dim textLength As Integer = text.Length
            Dim patternLength As Integer = pattern.Length
            Dim result As Integer = 0
            Dim mismatch As Boolean = False

            While textIndex + patternLength < textLength
                For subtrahend As Integer = 1 To patternLength
                    Dim patternIndex As Integer = patternLength - subtrahend
                    If text.Chars(textIndex + patternIndex) = pattern.Chars(patternIndex) Then
                        Continue For
                    End If
                    If skipTable.ContainsKey(text.Chars(textIndex + patternIndex)) Then
                        Dim s As Integer = skipTable(text.Chars(textIndex + patternIndex)) - (patternLength - 1 - patternIndex)
                        If 0 < s Then
                            textIndex += s
                        Else
                            ' 照合開始位置が前に戻るので、1文字だけズラす
                            textIndex += 1
                        End If
                    Else
                        If patternIndex < patternLength - 1 Then
                            textIndex += patternLength - 1 - patternIndex
                        Else
                            textIndex += patternLength
                        End If
                    End If
                    mismatch = True
                    Exit For
                Next

                ' 照合失敗
                If mismatch Then
                    mismatch = False
                    Continue While
                End If

                ' 照合成功
                If skipTable.ContainsKey(text.Chars(textIndex + patternLength - 1)) Then
                    textIndex += skipTable(text.Chars(textIndex + patternLength - 1))
                Else
                    textIndex += patternLength
                End If

                result += 1
            End While

            Return result
        End Function

    End Class
End Namespace