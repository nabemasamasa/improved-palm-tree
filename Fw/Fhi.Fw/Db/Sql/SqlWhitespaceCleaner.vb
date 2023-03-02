Imports System.Text

Namespace Db.Sql
    ''' <summary>
    ''' SQL文中の連続する無駄なホワイトスペースをお掃除するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SqlWhitespaceCleaner

        Private Const SQ As Char = "'"c
        Private Const DQ As Char = """"c
        Private Const SPACE As Char = " "c

        ''' <summary>
        ''' SQL文中の連続する無駄なホワイトスペースを除去する
        ''' </summary>
        ''' <param name="sql">SQL文</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Clean(ByVal sql As String) As String
            Dim modeSq As Boolean = False
            Dim modeDq As Boolean = False
            Dim modeWhitespace As Boolean = False
            Dim result As New StringBuilder
            For Each c As Char In sql.Trim.ToCharArray
                If modeSq Then
                    If c = SQ Then
                        modeSq = False
                    End If
                ElseIf modeDq Then
                    If c = DQ Then
                        modeDq = False
                    End If
                ElseIf c = SPACE OrElse c = vbTab OrElse c = vbCrLf OrElse c = vbCr OrElse c = vbLf Then
                    If Not modeWhitespace Then
                        result.Append(SPACE)
                        modeWhitespace = True
                    End If
                    Continue For
                Else
                    modeWhitespace = False
                    If c = SQ Then
                        modeSq = True
                    ElseIf c = DQ Then
                        modeDq = True
                    End If
                End If
                result.Append(c)
            Next
            Return result.ToString
        End Function
    End Class
End Namespace