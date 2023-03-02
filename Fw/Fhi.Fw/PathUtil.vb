Imports System.IO

''' <summary>
''' パスに関するユーティリティクラス
''' </summary>
''' <remarks></remarks>
Public Class PathUtil

    ''' <summary>
    ''' パスを結合する
    ''' </summary>
    ''' <param name="paths">パス</param>
    ''' <returns>結合されたパス</returns>
    ''' <remarks></remarks>
    Public Shared Function Combine(ByVal ParamArray paths As String()) As String
        If CollectionUtil.IsEmpty(paths) Then
            Throw New ArgumentNullException("paths")
        End If

        Dim result As String = Nothing
        For i = 0 To paths.Length - 1
            Dim p As String = StringUtil.ToString(paths(i))
            result = If(i = 0, p, Path.Combine(result, p))
        Next
        Return result
    End Function

    ''' <summary>
    ''' ファイルパスを中略して返す
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <param name="maxLength"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function OmitPath(filePath As String, maxLength As Integer) As String
        If filePath.Length <= maxLength Then
            Return filePath
        End If
        Dim tmpFilePath As String = filePath
        Dim splitPaths As New List(Of String)({"...", Path.GetFileName(tmpFilePath)})
        Dim thisPath As String = Combine(splitPaths.ToArray)
        If maxLength < thisPath.Length Then
            Return thisPath
        End If
        Dim lastPath As String = thisPath
        Dim pathRoot As String = Path.GetPathRoot(tmpFilePath)
        splitPaths.Insert(0, pathRoot)
        Do
            thisPath = Combine(splitPaths.ToArray)
            If maxLength < thisPath.Length Then
                Return lastPath
            End If
            lastPath = thisPath
            tmpFilePath = Path.GetDirectoryName(tmpFilePath)
            splitPaths.Insert(2, Path.GetFileName(tmpFilePath))
        Loop While pathRoot.Length < tmpFilePath.Length
        Throw New InvalidProgramException
    End Function

End Class
