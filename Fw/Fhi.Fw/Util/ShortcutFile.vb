Namespace Util
    ''' <summary>
    ''' Windowsショートカットファイルを担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ShortcutFile : Implements IDisposable

        ''' <summary>ショートカットファイルのリンク先</summary>
        <Obsolete()> Public WriteOnly Property LinkPathOrLinkUrl() As String
            Set(ByVal value As String)
                _TargetPath = value
            End Set
        End Property
        ''' <summary>ショートカットファイルのパスファイル名</summary>
        <Obsolete()> Public WriteOnly Property FileName() As String
            Set(ByVal value As String)
                _FullName = value
            End Set
        End Property

        ''' <summary>実行時の大きさ</summary>
        Public Enum Style
            ''' <summary>通常のウィンドウ</summary>
            Normal = 1
            ''' <summary>最大化</summary>
            Maximum = 3
            ''' <summary>最小化</summary>
            Minimum = 7
        End Enum

        Private ReadOnly wshShellType As Type = GetWshShellType()
        Private shortcut As Object

#Region "Public properties..."
        Private _TargetPath As String
        ''' <summary>リンク先のファイル名</summary>
        Public Property TargetPath() As String
            Get
                If shortcut Is Nothing Then
                    Return _TargetPath
                End If
                Return StringUtil.ToString(wshShellType.InvokeMember("TargetPath", System.Reflection.BindingFlags.GetProperty, _
                                                                     Nothing, shortcut, Nothing))
            End Get
            Set(ByVal value As String)
                If shortcut Is Nothing Then
                    _TargetPath = value
                    Return
                End If
                wshShellType.InvokeMember("TargetPath", System.Reflection.BindingFlags.SetProperty, _
                                          Nothing, shortcut, New Object() {value})
            End Set
        End Property

        Private _FullName As String
        ''' <summary>ショートカットファイルのフルパス</summary>
        Public Property FullName() As String
            Get
                If shortcut Is Nothing Then
                    Return _FullName
                End If
                Return StringUtil.ToString(wshShellType.InvokeMember("FullName", System.Reflection.BindingFlags.GetProperty, _
                                                                     Nothing, shortcut, Nothing))
            End Get
            Set(ByVal value As String)
                If shortcut Is Nothing Then
                    _FullName = value
                    Return
                End If
                Throw New InvalidOperationException("#Create()後はFullNameを変更できない")
            End Set
        End Property

        ''' <summary>ショートカットのアイコン</summary>
        Public Property IconLocation() As String
            Get
                AssertShortcutExist()
                Return StringUtil.ToString(wshShellType.InvokeMember("IconLocation", System.Reflection.BindingFlags.GetProperty, _
                                                                     Nothing, shortcut, Nothing))
            End Get
            Set(ByVal value As String)
                AssertShortcutExist()
                wshShellType.InvokeMember("IconLocation", System.Reflection.BindingFlags.SetProperty, _
                                          Nothing, shortcut, New Object() {value})
            End Set
        End Property

        ''' <summary>ショートカットの説明</summary>
        Public Property Description() As String
            Get
                AssertShortcutExist()
                Return StringUtil.ToString(wshShellType.InvokeMember("Description", System.Reflection.BindingFlags.GetProperty, _
                                                                     Nothing, shortcut, Nothing))
            End Get
            Set(ByVal value As String)
                AssertShortcutExist()
                wshShellType.InvokeMember("Description", System.Reflection.BindingFlags.SetProperty, _
                                          Nothing, shortcut, New Object() {value})
            End Set
        End Property

        ''' <summary>実行ファイルに渡すパラメータ</summary>
        Public Property Arguments() As String
            Get
                AssertShortcutExist()
                Return StringUtil.ToString(wshShellType.InvokeMember("Arguments", System.Reflection.BindingFlags.GetProperty, _
                                                                     Nothing, shortcut, Nothing))
            End Get
            Set(ByVal value As String)
                AssertShortcutExist()
                wshShellType.InvokeMember("Arguments", System.Reflection.BindingFlags.SetProperty, _
                                          Nothing, shortcut, New Object() {value})
            End Set
        End Property

        ''' <summary>作業フォルダ</summary>
        Public Property WorkingDirectory() As String
            Get
                AssertShortcutExist()
                Return StringUtil.ToString(wshShellType.InvokeMember("WorkingDirectory", System.Reflection.BindingFlags.GetProperty, _
                                                                     Nothing, shortcut, Nothing))
            End Get
            Set(ByVal value As String)
                AssertShortcutExist()
                wshShellType.InvokeMember("WorkingDirectory", System.Reflection.BindingFlags.SetProperty, _
                                          Nothing, shortcut, New Object() {value})
            End Set
        End Property

        ''' <summary>実行時の大きさ</summary>
        Public Property WindowStyle() As Style
            Get
                AssertShortcutExist()
                Dim value As Object = wshShellType.InvokeMember("WindowStyle", System.Reflection.BindingFlags.GetProperty, _
                                                                Nothing, shortcut, Nothing)
                Return DirectCast([Enum].ToObject(gettype(Style), value), style) 
            End Get
            Set(ByVal value As Style)
                AssertShortcutExist()
                wshShellType.InvokeMember("WindowStyle", System.Reflection.BindingFlags.SetProperty, _
                                          Nothing, shortcut, New Object() {value})
            End Set
        End Property

        ''' <summary>キーボード・ショートカット</summary>
        Public Property Hotkey() As String
            Get
                AssertShortcutExist()
                Return StringUtil.ToString(wshShellType.InvokeMember("Hotkey", System.Reflection.BindingFlags.GetProperty, _
                                                                     Nothing, shortcut, Nothing))
            End Get
            Set(ByVal value As String)
                AssertShortcutExist()
                wshShellType.InvokeMember("Hotkey", System.Reflection.BindingFlags.SetProperty, _
                                          Nothing, shortcut, New Object() {value})
            End Set
        End Property
#End Region

        Private Sub New()
        End Sub

        Public Sub New(ByVal fullName As String)
            Me.New(fullName, Nothing)
        End Sub

        Public Sub New(ByVal fullName As String, ByVal targetPath As String)
            If Not String.IsNullOrEmpty(targetPath) Then
                ' `C:\` はlnk にして `http:` はurlにしたい
                Dim extension As String = If(2 <= targetPath.IndexOf(":"), "url", "lnk")
                If Not fullName.EndsWith("." & extension) Then
                    fullName &= "." & extension
                End If
                Me.TargetPath = targetPath
            End If
            Me.FullName = fullName
            Open()
        End Sub

        ''' <summary>
        ''' （内部使用）ショートカットオブジェクトを生成する
        ''' </summary>
        ''' <param name="pathFileName">ショートカットファイルのパスファイル名</param>
        ''' <remarks></remarks>
        Private Sub InternalCreateShortcut(ByVal pathFileName As String)
            ReleaseComObject(shortcut)
            Dim shell As Object = VoUtil.NewInstance(wshShellType)
            Try
                shortcut = wshShellType.InvokeMember("CreateShortcut", System.Reflection.BindingFlags.InvokeMethod, _
                                                     Nothing, shell, New Object() {pathFileName})
            Finally
                ReleaseComObject(shell)
            End Try
        End Sub

        ''' <summary>
        ''' ショートカットファイルを開く
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Open()
            InternalCreateShortcut(_FullName)
            If _TargetPath Is Nothing Then
                Return
            End If
            TargetPath = _TargetPath
            If _FullName.ToLower.EndsWith(".lnk") Then
                IconLocation = _TargetPath & ",0"
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Close()
        End Sub

        ''' <summary>
        ''' ショートカットファイルを閉じる
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Close()
            ReleaseComObject(shortcut)
        End Sub

        ''' <summary>
        ''' 保存する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Save()
            AssertShortcutExist()
            wshShellType.InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, Nothing, shortcut, Nothing)
        End Sub

        Private Sub AssertShortcutExist()
            If shortcut Is Nothing Then
                Throw New InvalidOperationException("#Open()してから処理すべき")
            End If
        End Sub

        ''' <summary>
        ''' ショートカットファイルを作成する
        ''' </summary>
        ''' <param name="pathFileName">ショートカットファイルの作成先パスファイル名 ※*.lnkまたは*.urlであること</param>
        ''' <param name="linkPathOrLinkUrl">ショートカットファイルのリンク先 ※PATHかURL</param>
        ''' <remarks></remarks>
        Public Shared Sub Create(ByVal pathFileName As String, ByVal linkPathOrLinkUrl As String)
            Using shortcut As New ShortcutFile With {.FullName = pathFileName, .TargetPath = linkPathOrLinkUrl}
                shortcut.Open()
                shortcut.Save()
                shortcut.Close()
            End Using
        End Sub

        ''' <summary>
        ''' WshShell型を返す
        ''' </summary>
        ''' <returns>WshShell型</returns>
        ''' <remarks></remarks>
        Private Shared Function GetWshShellType() As Type
            Return Type.GetTypeFromCLSID(New Guid("72C24DD5-D70A-438B-8A42-98424B88AFB8"))
        End Function

        Private Shared Sub ReleaseComObject(Of T As Class)(ByRef comObject As T, Optional ByVal force As Boolean = False)
            If comObject Is Nothing Then
                Return
            End If
            If System.Runtime.InteropServices.Marshal.IsComObject(comObject) Then
                If force Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comObject)
                Else
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject)
                End If
            End If
            comObject = Nothing
        End Sub

    End Class
End Namespace