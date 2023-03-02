Imports System.Text

Namespace Util
    ''' <summary>
    ''' ランダムな文字列を生成するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RandomStringBuilder

        Private Const CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
        Private ReadOnly r As Random
        Private ReadOnly generatedStrings As New Dictionary(Of Integer, List(Of String))

        Private Class Key
            Private Shared ReadOnly instance As New Key
            Public Shared Function Increment() As Integer
                Return instance.PerformIncrement
            End Function

            Private uniqueValue As Integer = System.Environment.TickCount
            Private Function PerformIncrement() As Integer
                SyncLock Me
                    uniqueValue += 1
                    Return uniqueValue
                End SyncLock
            End Function
        End Class

        Public Sub New()
            r = New Random(Key.Increment)
        End Sub

        ''' <summary>
        ''' ランダムな文字列を生成する
        ''' </summary>
        ''' <param name="length">文字列の長さ</param>
        ''' <returns>ランダムな文字列</returns>
        ''' <remarks></remarks>
        Public Function CreateRandomString(ByVal length As Integer) As String
            If generatedStrings.ContainsKey(length) _
                AndAlso CHARS.Length ^ length <= generatedStrings(length).Count Then
                Throw New InvalidOperationException(String.Format("これ以上{0}桁で重複のない文字列を抽出できません", length.ToString))
            End If
            Dim sb As New StringBuilder
            For i As Integer = 1 To length
                sb.Append(CHARS.Substring(r.Next(0, CHARS.Length), 1))
            Next
            Dim result As String = sb.ToString
            If generatedStrings.ContainsKey(length) _
                AndAlso generatedStrings(length).Contains(result) Then
                Return CreateRandomString(length)
            End If

            If Not generatedStrings.ContainsKey(length) Then
                generatedStrings.Add(length, New List(Of String))
            End If
            generatedStrings(length).Add(result)
            Return result
        End Function

    End Class
End Namespace
