Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

''' <summary>
''' ファイル操作に関するユーティリティ集
''' </summary>
''' <remarks></remarks>
Public Class FileUtil

    <DllImport("kernel32.dll", EntryPoint:="CreateHardLinkA", CharSet:=CharSet.Ansi, SetLastError:=True, ExactSpelling:=True)> _
    Private Shared Function CreateHardLink(ByVal lpFileName As String, ByVal lpExistingFilename As String, ByVal lpSecurityAttirbute As Integer) As Long
    End Function

    ''' <summary>出力時の文字エンコード</summary>
    Public Shared ReadOnly DEFAULT_ENCODING As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")

    ''' <summary>
    ''' ファイルを作成する
    ''' </summary>
    ''' <param name="fileName">パスつきのファイル名</param>
    ''' <param name="value">ファイルの内容</param>
    ''' <param name="isOverride">上書きする場合、true</param>
    ''' <remarks></remarks>
    Public Overloads Shared Sub WriteFile(ByVal fileName As String, ByVal value As String, Optional ByVal isOverride As Boolean = True)
        PerformWriteFile(fileName, value, DEFAULT_ENCODING, If(isOverride, WriteFileMode.OVERRIDE, WriteFileMode.NON_OVERRIDE))
    End Sub

    ''' <summary>
    ''' ファイルを作成する
    ''' </summary>
    ''' <param name="fileName">パスつきのファイル名</param>
    ''' <param name="value">ファイルの内容</param>
    ''' <param name="anEncoding">文字エンコード</param>
    ''' <param name="isOverride">上書きする場合、true</param>
    ''' <remarks></remarks>
    Public Overloads Shared Sub WriteFile(ByVal fileName As String, ByVal value As String, ByVal anEncoding As String, Optional ByVal isOverride As Boolean = True)
        PerformWriteFile(fileName, value, System.Text.Encoding.GetEncoding(anEncoding), If(isOverride, WriteFileMode.OVERRIDE, WriteFileMode.NON_OVERRIDE))
    End Sub

    ''' <summary>
    ''' ファイルに追記する
    ''' </summary>
    ''' <param name="fileName">パスつきのファイル名</param>
    ''' <param name="value">追記する内容</param>
    ''' <remarks></remarks>
    Public Shared Sub AppendFile(ByVal fileName As String, ByVal value As String)
        PerformWriteFile(fileName, value, DEFAULT_ENCODING, WriteFileMode.APPEND)
    End Sub

    Private Enum WriteFileMode
        NON_OVERRIDE
        OVERRIDE
        APPEND
    End Enum
    Private Shared Sub PerformWriteFile(ByVal fileName As String, ByVal value As String, ByVal anEncoding As Encoding, ByVal aMode As WriteFileMode)
        If System.IO.File.Exists(fileName) Then
            If aMode = WriteFileMode.OVERRIDE Then
                System.IO.File.Delete(fileName)
            ElseIf aMode = WriteFileMode.NON_OVERRIDE Then
                Throw New System.InvalidOperationException(fileName & " はすでに存在.")
            End If
        End If

        Dim appends As Boolean = aMode = WriteFileMode.APPEND
        Using writer As New System.IO.StreamWriter(fileName, appends, anEncoding)
            writer.WriteLine(value)
        End Using
    End Sub

    ''' <summary>
    ''' ファイルを読み込み内容を返す
    ''' </summary>
    ''' <param name="fileName">パスつきのファイル名</param>
    ''' <returns>ファイルの内容　※配列の各要素が、内容の1行を表す</returns>
    ''' <remarks></remarks>
    Public Shared Function ReadFileAsArray(ByVal fileName As String) As String()
        Dim result As New List(Of String)
        Using reader As New System.IO.StreamReader(fileName, DEFAULT_ENCODING)
            While 0 < reader.Peek
                result.Add(reader.ReadLine)
            End While
        End Using
        Return result.ToArray
    End Function

    ''' <summary>
    ''' ファイルを読み込み内容を返す
    ''' </summary>
    ''' <param name="fileName">パスつきのファイル名</param>
    ''' <returns>ファイルの内容</returns>
    ''' <remarks></remarks>
    Public Shared Function ReadFile(ByVal fileName As String) As String
        Return Join(ReadFileAsArray(fileName), vbCrLf)
    End Function

    ''' <summary>
    ''' ファイルの存在チェック処理及び削除処理
    ''' </summary>
    ''' <param name="fileFullPath"></param>
    ''' <remarks></remarks>
    Public Shared Sub DeleteFileIfExist(ByVal fileFullPath As String)

        'ファイルの存在チェック
        If Not System.IO.File.Exists(fileFullPath) Then
            Exit Sub
        End If

        'ファイルの削除処理
        System.IO.File.Delete(fileFullPath)

    End Sub

    ''' <summary>
    ''' フォルダの存在チェック処理及び削除処理
    ''' </summary>
    ''' <param name="foldrePath"></param>
    ''' <remarks></remarks>
    Public Shared Sub DeleteFolderIfExist(ByVal foldrePath As String, Optional ByVal recursive As Boolean = False)

        'フォルダの存在チェック
        If Not System.IO.Directory.Exists(foldrePath) Then
            Exit Sub
        End If

        'フォルダの削除処理
        System.IO.Directory.Delete(foldrePath, recursive)

    End Sub

    ''' <summary>
    ''' ファイルの存在チェック処理及びリネーム処理
    ''' </summary>
    ''' <param name="sourceFileName">元のファイル名</param>
    ''' <param name="destFileName">変更後のファイル名</param>
    ''' <remarks>変更後のファイル名があればリネームしない</remarks>
    Public Shared Sub RenameFileIfExist(ByVal sourceFileName As String, ByVal destFileName As String)

        '元ファイルの存在チェック
        If Not System.IO.File.Exists(sourceFileName) Then
            Exit Sub
        End If

        '変更後のファイルの存在チェック
        If System.IO.File.Exists(destFileName) Then
            Exit Sub
        End If

        'ファイルのリネーム処理
        System.IO.File.Move(sourceFileName, destFileName)

    End Sub


    Private Shared ReadOnly CSV_SHOULD_ESCAPE_CHARS As Char() = ("""," & vbCrLf).ToCharArray
    Private Shared ReadOnly TSV_SHOULD_ESCAPE_CHARS As Char() = ("""" & vbTab & vbCrLf).ToCharArray

    ''' <summary>
    ''' 値を必要ならエスケープして返す
    ''' </summary>
    ''' <param name="value">値</param>
    ''' <returns>結果</returns>
    ''' <remarks></remarks>
    Private Shared Function EscapeValue(ByVal value As String, ByVal escapeChars As Char()) As String
        If value IsNot Nothing AndAlso 0 <= value.IndexOfAny(escapeChars) Then
            Return """" & value.Replace("""", """""") & """"
        End If
        Return value
    End Function

    ''' <summary>
    ''' CSV値を必要ならエスケープして返す
    ''' </summary>
    ''' <param name="value">CSV値</param>
    ''' <returns>結果</returns>
    ''' <remarks></remarks>
    Public Shared Function EscapeCsvValue(ByVal value As String) As String
        Return EscapeValue(value, CSV_SHOULD_ESCAPE_CHARS)
    End Function

    ''' <summary>
    ''' TSV値を必要ならエスケープして返す
    ''' </summary>
    ''' <param name="value">CSV値</param>
    ''' <returns>結果</returns>
    ''' <remarks></remarks>
    Public Shared Function EscapeTsvValue(ByVal value As String) As String
        Return EscapeValue(value, TSV_SHOULD_ESCAPE_CHARS)
    End Function

    ''' <summary>
    ''' CSV出力行情報を、必要ならエスケープして返す
    ''' </summary>
    ''' <param name="columns">出力行情報</param>
    ''' <returns>エスケープ済み行情報</returns>
    ''' <remarks></remarks>
    Public Shared Function OutputCsvRow(ByVal columns As ICollection(Of String)) As String
        Return OutputCsvRow(columns, ",")
    End Function

    ''' <summary>
    ''' CSV出力行情報を、必要ならエスケープして返す
    ''' </summary>
    ''' <param name="columns">出力行情報</param>
    ''' <param name="token">CSVの区切り文字</param>
    ''' <returns>エスケープ済み行情報</returns>
    ''' <remarks></remarks>
    Public Shared Function OutputCsvRow(ByVal columns As ICollection(Of String), ByVal token As String) As String
        Dim result As New List(Of String)
        If vbTab.Equals(token) Then
            For Each column As String In columns
                result.Add(EscapeTsvValue(column))
            Next
        Else
            For Each column As String In columns
                result.Add(EscapeCsvValue(column))
            Next
        End If
        Return Join(result.ToArray, token)
    End Function

    ''' <summary>
    ''' もしダブルクォートで囲まれたパスだったら、ダブルクォートを外して返す
    ''' </summary>
    ''' <param name="pathFileName">パス</param>
    ''' <returns>ダブルクォートを外したパス</returns>
    ''' <remarks></remarks>
    Private Shared Function RemoveIfEncloseDq(ByVal pathFileName As String) As String
        If pathFileName Is Nothing Then
            Return pathFileName
        End If
        If pathFileName.Length <= 1 Then
            Return pathFileName
        End If
        If Not (pathFileName.StartsWith("""") AndAlso pathFileName.EndsWith("""")) Then
            Return pathFileName
        End If
        Return pathFileName.Substring(1, pathFileName.Length - 2)
    End Function

    ''' <summary>
    ''' ファイルの存在チェックをして返す
    ''' </summary>
    ''' <param name="fileNames">存在チェックするパス付きのファイル</param>
    ''' <returns>一つでも存在しなければ、false</returns>
    ''' <remarks></remarks>
    Public Shared Function ExistsFiles(ByVal fileNames As String()) As Boolean
        For Each fileName As String In fileNames
            If Not File.Exists(RemoveIfEncloseDq(fileName)) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' ハードリンクを作成する
    ''' </summary>
    ''' <param name="srcPathFileName">リンク元ファイルパス名</param>
    ''' <param name="destPathFileName"></param>
    ''' <remarks></remarks>
    Public Shared Sub CreateHardLink(ByVal srcPathFileName As String, ByVal destPathFileName As String)

        Dim srcInfo As New FileInfo(srcPathFileName)
        Dim destInfo As New FileInfo(destPathFileName)
        If Not Path.GetPathRoot(srcInfo.DirectoryName).Equals(Path.GetPathRoot(destInfo.DirectoryName), StringComparison.CurrentCultureIgnoreCase) Then
            Throw New ArgumentException("違うドライブへハードリンクを作成できません.")
        End If
        CreateHardLink(destPathFileName, srcPathFileName, 0)
        'Process.Start("fsutil", String.Format("hardlink create ""{0}"" ""{1}""", destPathFileName, srcPathFileName))
    End Sub

    ''' <summary>
    ''' フォルダを作成する（フォルダが無い場合）
    ''' </summary>
    ''' <param name="folderPath">フォルダのパス</param>
    ''' <remarks></remarks>
    Public Shared Sub MakeFolderIfNotExists(ByVal folderPath As String)

        If Not Directory.Exists(folderPath) Then
            ' 仮に c:\a フォルダが無くても、 c:\a\b\c を作成できる
            Directory.CreateDirectory(folderPath)
        End If
    End Sub

    Private Class FileNameEscape
        Public ReadOnly [Char] As Char
        Public ReadOnly EscapeChar As Char
        Public Sub New(ByVal [Char] As Char, ByVal EscapeChar As Char)
            Me.[Char] = [Char]
            Me.EscapeChar = EscapeChar
        End Sub
    End Class

    Private Class EscapeHolder
        Public Shared ReadOnly ESCAPES As FileNameEscape()

        Shared Sub New()
            ESCAPES = New FileNameEscape() {New FileNameEscape("\"c, "￥"c), New FileNameEscape("/"c, "／"c), _
                                            New FileNameEscape(":"c, "："c), New FileNameEscape("*"c, "＊"c), _
                                            New FileNameEscape("?"c, "？"c), New FileNameEscape(""""c, "’"c), _
                                            New FileNameEscape("<"c, "＜"c), New FileNameEscape(">"c, "＞"c)}
        End Sub
    End Class

    ''' <summary>
    ''' ファイル名に使用できない文字を代替文字に差替えて返す
    ''' </summary>
    ''' <param name="fileName">ファイル名</param>
    ''' <returns>差替えたファイル名</returns>
    ''' <remarks></remarks>
    Public Shared Function EscapeFileName(ByVal fileName As String) As String
        Dim result As String = fileName
        For Each escape As FileNameEscape In EscapeHolder.ESCAPES
            If 0 <= result.IndexOf(escape.Char) Then
                result = result.Replace(escape.Char, escape.EscapeChar)
            End If
        Next
        Return result
    End Function

    ''' <summary>
    ''' 特定ディレクトリ配下のファイルパス[]を取得する
    ''' </summary>
    ''' <param name="target">探索させたいフォルダのパス</param>
    ''' <param name="isRecursive">引数フォルダ配下のフォルダ内ファイルも取得したいならTrue (初期値:False)</param>
    ''' <returns>ファイルパス[]</returns>
    ''' <remarks></remarks>
    Public Shared Function GetFilePathsUnderDirectory(target As String, Optional isRecursive As Boolean = False) As String()
        Dim detectDepthCallback As Func(Of String, Integer) = Function(filePath As String) Replace(filePath, target, "").ToCharArray.Where(Function(c) "\"c.Equals(c)).Count
        Dim results As List(Of String) = RecurGetFilePathUnderDirectory(target, isRecursive)
        results.Sort(Function(a, b)
                         Dim result As Integer = EzUtil.CompareDescNullsLast(detectDepthCallback(a), detectDepthCallback(b))
                         If result <> 0 Then
                             Return result
                         End If
                         Return EzUtil.CompareNullsLast(a, b)
                     End Function)
        Return results.ToArray
    End Function

    Private Shared Function RecurGetFilePathUnderDirectory(target As String, isRecursive As Boolean) As List(Of String)
        Dim results As New List(Of String)
        If isRecursive Then
            Dim directories As String() = Directory.GetDirectories(target)
            For Each directory As String In directories
                results.AddRange(RecurGetFilePathUnderDirectory(directory, isRecursive))
            Next
        End If
        results.AddRange(Directory.GetFiles(target))
        Return results
    End Function

End Class
