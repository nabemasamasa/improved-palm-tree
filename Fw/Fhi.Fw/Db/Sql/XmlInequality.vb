Imports System.Text

Namespace Db.Sql
    ''' <summary>
    ''' XMLの属性にもつ、不等号を、&gt; に変更する
    ''' </summary>
    ''' <remarks></remarks>
    Public Class XmlInequality
        Private Const SQ As Char = "'"c
        Private Const DQ As Char = """"c
        Private Const LT As Char = "<"c
        Private Const GT As Char = ">"c
        Private Const ESCAPED_LT As String = "&lt;"
        Private Const ESCAPED_GT As String = "&gt;"

        Private Sub New()
        End Sub

        ''' <summary>
        ''' 引用符内の不等号を、エスケープ(&gt; or &lt;)する
        ''' </summary>
        ''' <param name="xmlStr">エスケープ前のXML文字列</param>
        ''' <returns>エスケープ後のXML文字列</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvInequality(ByVal xmlStr As String) As String
            Dim result As New StringBuilder
            Dim dqing As Boolean
            Dim sqing As Boolean
            For Each c As Char In xmlStr.ToCharArray
                If c = DQ Then
                    If Not sqing Then
                        dqing = Not dqing
                    End If
                ElseIf c = SQ Then
                    If Not dqing Then
                        sqing = Not sqing
                    End If
                ElseIf dqing OrElse sqing Then
                    If c = LT Then
                        result.Append(ESCAPED_LT)
                        Continue For
                    ElseIf c = GT Then
                        result.Append(ESCAPED_GT)
                        Continue For
                    End If
                End If
                result.Append(c)
            Next
            Return result.ToString
        End Function
    End Class
End Namespace