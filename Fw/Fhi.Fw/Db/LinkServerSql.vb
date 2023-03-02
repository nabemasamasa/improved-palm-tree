Imports System.Text

Namespace Db
    ''' <summary>
    ''' リンクサーバー向けのSQLを担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class LinkServerSql

        Private ReadOnly linkServerName As String

        Public Sub New(ByVal linkServerName As String)
            EzUtil.AssertParameterIsNotNull(linkServerName, "linkServerName")
            Me.linkServerName = linkServerName
        End Sub

        ''' <summary>
        ''' リンクサーバー向けOPENQUERYに変換する
        ''' </summary>
        ''' <param name="boundSql">値埋め込み済のSelect文</param>
        ''' <returns>OPENQUERY SQL</returns>
        ''' <remarks></remarks>
        Public Function ConvertOpenqueryFrom(ByVal boundSql As String) As String
            Dim upperBoundSql As String = boundSql.ToUpper
            If upperBoundSql.StartsWith("SELECT ") Then
                Return String.Format("SELECT * FROM OPENQUERY( {0}, '{1}')", linkServerName, boundSql.Replace("'", "''"))

            ElseIf upperBoundSql.StartsWith("INSERT ") AndAlso Not upperBoundSql.StartsWith("INSERT SELECT") Then
                Dim indexOfValues As Integer = upperBoundSql.IndexOf("VALUES")
                Dim indexOfFirstOpenBracket As Integer = boundSql.IndexOf("(")
                Dim indexOfStartTable As Integer = upperBoundSql.IndexOf("INTO") + "INTO ".Length
                Dim tableName As String = StringUtil.Trim(boundSql.Substring(indexOfStartTable, Math.Min(indexOfValues, indexOfFirstOpenBracket) - indexOfStartTable))
                Dim columnsPart As String = If(indexOfValues < indexOfFirstOpenBracket, "*", StringUtil.Trim(boundSql.Substring(indexOfFirstOpenBracket + 1, boundSql.IndexOf(")") - 1 - indexOfFirstOpenBracket)))
                Dim valuesPart As String = StringUtil.Trim(boundSql.Substring(indexOfValues))
                Return String.Format("INSERT OPENQUERY( {0}, 'SELECT {1} FROM {2} WHERE 1=0') {3}", linkServerName, columnsPart, tableName, valuesPart)

            ElseIf upperBoundSql.StartsWith("UPDATE ") Then
                Dim indexOfWhere As Integer = upperBoundSql.IndexOf("WHERE")
                Dim indexOfTable As Integer = "UPDATE ".Length
                Dim indexOfSet As Integer = upperBoundSql.IndexOf("SET")
                Dim tableName As String = StringUtil.Trim(boundSql.Substring(indexOfTable, indexOfSet - indexOfTable))
                Dim wherePart As String = StringUtil.Trim(boundSql.Substring(indexOfWhere))
                Dim setPart As String = StringUtil.Trim(boundSql.Substring(indexOfSet, indexOfWhere - indexOfSet))
                Return String.Format("UPDATE OPENQUERY( {0}, 'SELECT * FROM {1} {2}') {3}", linkServerName, tableName, wherePart.Replace("'", "''"), setPart)

            ElseIf upperBoundSql.StartsWith("DELETE ") Then
                Dim indexOfFromNext As Integer = upperBoundSql.IndexOf("FROM ") + "FROM ".Length
                Dim indexOfWhere As Integer = upperBoundSql.IndexOf("WHERE")
                Dim tableName As String = StringUtil.Trim(boundSql.Substring(indexOfFromNext, indexOfWhere - indexOfFromNext))
                Dim wherePart As String = StringUtil.Trim(boundSql.Substring(indexOfWhere))
                Return String.Format("DELETE OPENQUERY( {0}, 'SELECT * FROM {1} {2}')", linkServerName, tableName, wherePart.Replace("'", "''"))
            End If
            Throw New InvalidProgramException(String.Format("このSQL構文は未対応 {0}", StringUtil.OmitIfLengthOver(boundSql, 20)))
        End Function

        ''' <summary>
        ''' リンクサーバー向け４部構成(linked_server_name.catalog.schema.object_name)に変換する
        ''' </summary>
        ''' <param name="boundSql">値埋め込み済のSelect文</param>
        ''' <returns>OPENQUERY SQL</returns>
        ''' <remarks></remarks>
        Public Function ConvertFourPartNamesFrom(ByVal boundSql As String) As String
            Dim upperBoundSql As String = boundSql.ToUpper
            If upperBoundSql.StartsWith("SELECT ") OrElse upperBoundSql.StartsWith("DELETE ") Then
                Dim indexOfFromNext As Integer = upperBoundSql.IndexOf("FROM ") + "FROM ".Length

                Dim fromIndex As Integer = 0
                While 0 < upperBoundSql.IndexOf("FROM ", fromIndex)
                    fromIndex = upperBoundSql.IndexOf("FROM ", fromIndex) + "FROM ".Length
                    While 0 <= (" " & vbTab & vbCrLf).IndexOf(upperBoundSql.Chars(fromIndex))
                        fromIndex += 1
                    End While
                    indexOfFromNext = fromIndex
                End While

                Dim indices As New List(Of Integer)
                indices.Add(indexOfFromNext)
                If upperBoundSql.StartsWith("SELECT ") Then
                    Dim idx As Integer = indexOfFromNext
                    While 0 < upperBoundSql.IndexOf("JOIN ", idx)
                        idx = upperBoundSql.IndexOf("JOIN ", idx) + "JOIN ".Length
                        While 0 <= (" " & vbTab & vbCrLf).IndexOf(upperBoundSql.Chars(idx))
                            idx += 1
                        End While
                        indices.Add(idx)
                    End While
                End If

                Dim result As New StringBuilder(boundSql.Substring(0, indices(0)))
                For i As Integer = 0 To indices.Count - 1
                    result.Append(GetPrefixOfTableForUpdating)
                    If indices.Count - 1 <= i Then
                        result.Append(boundSql.Substring(indices(i)))
                    Else
                        result.Append(boundSql.Substring(indices(i), indices(i + 1) - indices(i)))
                    End If
                Next
                Return result.ToString

            ElseIf upperBoundSql.StartsWith("UPDATE ") Then
                Dim indexOfTable As Integer = "UPDATE ".Length
                Return New StringBuilder(boundSql.Substring(0, indexOfTable)).Append(GetPrefixOfTableForUpdating).Append(boundSql.Substring(indexOfTable)).ToString

            ElseIf upperBoundSql.StartsWith("INSERT ") AndAlso Not upperBoundSql.StartsWith("INSERT SELECT") Then
                Dim indexOfStartTable As Integer = upperBoundSql.IndexOf("INTO") + "INTO ".Length
                Return New StringBuilder(boundSql.Substring(0, indexOfStartTable)).Append(GetPrefixOfTableForUpdating).Append(boundSql.Substring(indexOfStartTable)).ToString
            End If
            Throw New InvalidProgramException(String.Format("このSQL構文は未対応 {0}", StringUtil.OmitIfLengthOver(boundSql, 20)))
        End Function

        ''' <summary>
        ''' リンクサーバー向け４部構成(linked_server_name.catalog.schema.object_name)の接頭後を返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetPrefixOfTableForUpdating() As String
            If StringUtil.IsEmpty(linkServerName) Then
                Return ""
            End If
            Dim strings As String() = Split(linkServerName, "_")
            If strings.Length <> 2 Then
                Throw New NotSupportedException("未知のリンクサーバー名です. " & linkServerName)
            End If
            Return String.Format("[{0}].[{1}].[{2}].", linkServerName, strings(0), strings(1))
        End Function

    End Class
End Namespace