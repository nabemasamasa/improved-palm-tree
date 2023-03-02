Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports Microsoft.Win32
Imports System.Text

Public Class OsUtil
    ''' <summary>
    ''' エラーログをイベントビュアーに出力
    ''' </summary>
    ''' <param name="e">Exceptionオブジェクト</param>
    ''' <param name="sourceName">ソース名</param>
    ''' <remarks></remarks>
    Public Shared Sub OutputErrorLog(ByVal e As Exception, ByVal sourceName As String)

        Try
            'ソースが存在していない時は、作成する
            If Not System.Diagnostics.EventLog.SourceExists(sourceName) Then
                'ログ名を空白にすると、"Application"となる

                System.Diagnostics.EventLog.CreateEventSource(sourceName, "")
            End If
            'テスト用にイベントログエントリに付加するデータを適当に作る
            Dim myData() As Byte = {}
            Dim msg As String = "例外：" & vbNewLine & e.Message & vbNewLine & "例外スタックトレース" & vbNewLine & e.StackTrace & vbNewLine
            If e.InnerException IsNot Nothing Then
                msg = msg & "InnerException:" & vbNewLine & e.InnerException.Message & vbNewLine & "InnerExceptionスタックトレース:" & vbNewLine & e.InnerException.StackTrace
            End If
            'イベントログにエントリを書き込む
            'ここではエントリの種類をエラー、イベントIDを1、分類を1000とする
            System.Diagnostics.EventLog.WriteEntry(sourceName, msg, System.Diagnostics.EventLogEntryType.Error, 1, 1000, myData)
        Catch ignore As Exception
            '' ここの例外は無視 by T.Homma
        End Try
    End Sub

    Private Class ShellGetFileInfo
        Private Declare Ansi Function SHGetFileInfo Lib "shell32.dll" (ByVal pszPath As String, ByVal dwFileAttiributes As Integer, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As Integer) As IntPtr
        Private Structure SHFILEINFO
            Public hIcon As IntPtr
            Public iIcon As IntPtr
            Public dwAttributes As Integer
            <MarshalAs(UnmanagedType.ByValTStr, sizeconst:=260)> Public szDisplayName As String
            <MarshalAs(UnmanagedType.ByValTStr, sizeconst:=80)> Public szTypeName As String
        End Structure
        Private Const SHGFI_ICON As Integer = &H100
        Private Const SHGFI_LARGEICON As Integer = &H0
        Private Const SHGFI_SMALLICON As Integer = &H1
        Private Const SHGFI_TYPENAME As Integer = &H400

        Public Function ExtractSmallIcon(ByVal pathFileName As String) As Icon
            Dim shInfo As New SHFILEINFO
            Dim hSuccess As IntPtr = SHGetFileInfo(pathFileName, 0, shInfo, Marshal.SizeOf(shInfo), SHGFI_ICON Or SHGFI_SMALLICON)
            If hSuccess.Equals(IntPtr.Zero) Then
                Return Nothing
            End If
            Return Icon.FromHandle(shInfo.hIcon)
        End Function

        Public Function ExtractTypeName(ByVal pathFileName As String) As String
            Dim shInfo As New SHFILEINFO
            Dim hSuccess As IntPtr = SHGetFileInfo(pathFileName, 0, shInfo, Marshal.SizeOf(shInfo), SHGFI_TYPENAME)
            If hSuccess.Equals(IntPtr.Zero) Then
                Return Nothing
            End If
            Return shInfo.szTypeName
        End Function
    End Class

    ''' <summary>
    ''' ファイルに関連付けられた（小さい）アイコンを返す
    ''' </summary>
    ''' <param name="pathFileName">ファイルパス</param>
    ''' <returns>（小さい）アイコン</returns>
    ''' <remarks></remarks>
    Public Shared Function ExtractSmallIcon(ByVal pathFileName As String) As Icon
        Dim shell32 As New ShellGetFileInfo
        Return shell32.ExtractSmallIcon(pathFileName)
    End Function

    ''' <summary>
    ''' ファイルに関連付けられた"種類"名を返す
    ''' </summary>
    ''' <param name="pathFileName">ファイルパス</param>
    ''' <returns>"種類"名</returns>
    ''' <remarks></remarks>
    Public Shared Function ExtractTypeName(ByVal pathFileName As String) As String
        Dim shell32 As New ShellGetFileInfo
        Return shell32.ExtractTypeName(pathFileName)
    End Function

    ''' <summary>
    ''' 「ファイルを開くプログラムの選択」ダイアログを表示し、選んだプログラムでファイルを開く
    ''' </summary>
    ''' <param name="pathFileName">ファイル名(フルパス)</param>
    ''' <remarks></remarks>
    Public Shared Sub ShowOpenAs(ByVal pathFileName As String)

        ' 次のコマンドを実行して、「ファイルを開くプログラムの選択」を出す
        ' rundll32.exe shell32.dll,OpenAs_RunDLL ファイル
        Dim openAsProcess As New Process()
        openAsProcess.StartInfo.FileName = "rundll32.exe"
        openAsProcess.StartInfo.Arguments = "shell32.dll,OpenAs_RunDLL " & pathFileName
        openAsProcess.StartInfo.Verb = "Open"
        openAsProcess.Start()
    End Sub

    ''' <summary>
    ''' 「ファイルを開く」を関連付ける
    ''' </summary>
    ''' <param name="extension">拡張子（ドットを含める ".xxx"の書式）</param>
    ''' <param name="executeFilePath">実行ファイル名（フルパス）</param>
    ''' <param name="fileType">ファイルタイプ名（レジストリ内の関連付けキー名）</param>
    ''' <param name="description">エクスプローラに表示される「種類」</param>
    ''' <remarks>ファイルの関連付け = File assosiation</remarks>
    Public Shared Sub AssociateFile(ByVal extension As String, ByVal executeFilePath As String, ByVal fileType As String, ByVal description As String)
        Dim commandLine As String = """" & executeFilePath & """ ""%1"""
        Dim verb As String = "open"
        Dim verbDescription As String = "開く(&O)"
        Dim iconPath As String = executeFilePath
        Dim iconIndex As Integer = 0

        Const DEFAULT_KEY As String = ""

        Using extensionKey As RegistryKey = Registry.ClassesRoot.CreateSubKey(extension)
            SetValueIfNecessary(extensionKey, DEFAULT_KEY, fileType)
        End Using

        Using fileTypeKey As RegistryKey = Registry.ClassesRoot.CreateSubKey(fileType)
            SetValueIfNecessary(fileTypeKey, DEFAULT_KEY, description)

            '動詞とその説明
            Using verbKey As RegistryKey = fileTypeKey.CreateSubKey("shell\" & verb)
                SetValueIfNecessary(verbKey, DEFAULT_KEY, verbDescription)

                Using cmdKey As RegistryKey = verbKey.CreateSubKey("command")
                    SetValueIfNecessary(cmdKey, DEFAULT_KEY, commandLine)
                End Using
            End Using

            'アイコンの登録
            Using iconKey As RegistryKey = fileTypeKey.CreateSubKey("DefaultIcon")
                SetValueIfNecessary(iconKey, DEFAULT_KEY, String.Format("{0},{1}", iconPath, iconIndex))
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 値が存在しない、もしくは違う時に書き込む
    ''' </summary>
    ''' <param name="regKey">書き込み先のRegistryKey</param>
    ''' <param name="name">キー名</param>
    ''' <param name="value">値</param>
    ''' <remarks></remarks>
    Private Shared Sub SetValueIfNecessary(ByVal regKey As RegistryKey, ByVal name As String, ByVal value As String)

        If CollectionUtil.Contains(regKey.GetValueNames, name) AndAlso value.Equals(regKey.GetValue(name)) Then
            Return
        End If
        regKey.SetValue(name, value)
    End Sub

    ''' <summary>
    ''' 「ファイルを開く」に関連づいたコマンドを返す
    ''' </summary>
    ''' <param name="extension">拡張子（ドットを含める ".xxx"の書式）</param>
    ''' <returns>コマンド</returns>
    ''' <remarks></remarks>
    Public Shared Function GetAssosiatedCommand(ByVal extension As String) As String

        Using extensionKey As RegistryKey = Registry.ClassesRoot.OpenSubKey(extension)
            If extensionKey Is Nothing Then
                Return Nothing
            End If
            Dim extensionValue As Object = extensionKey.GetValue("")

            If StringUtil.IsEmpty(extensionValue) Then
                Return Nothing
            End If

            Using commandKey As RegistryKey = Registry.ClassesRoot.OpenSubKey(extensionValue.ToString & "\shell\open\command")
                If commandKey Is Nothing Then
                    Return Nothing
                End If
                Return StringUtil.ToString(commandKey.GetValue(""))
            End Using
        End Using
    End Function

    ''' <summary>
    ''' ウィンドウを最前面にする
    ''' </summary>
    ''' <param name="hWnd">ウィンドウハンドル</param>
    ''' <returns>成功したらtrue(?)</returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll")> Public Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    ''' <summary>
    ''' プロセス情報を集めたクラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class FoundProcessInfo

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="windowText">ウィンドウのタイトル</param>
        ''' <param name="className">ウィンドウが属するクラス名</param>
        ''' <remarks></remarks>
        Public Sub New(windowText As String, className As String)
            _windowText = windowText
            _className = className
        End Sub

        Private _foundProcesses As ArrayList = Nothing
        ''' <summary>プロセス情報</summary>
        Public ReadOnly Property FoundProcesses() As ArrayList
            Get
                If _foundProcesses Is Nothing Then
                    _foundProcesses = New ArrayList
                End If
                Return _foundProcesses
            End Get
        End Property

        Private ReadOnly _windowText As String
        ''' <summary>ウィンドウのタイトル</summary>
        Public ReadOnly Property WindowText() As String
            Get
                Return _windowText
            End Get
        End Property

        Private ReadOnly _className As String
        ''' <summary>ウィンドウが属するクラス名</summary>
        Public ReadOnly Property ClassName() As String
            Get
                Return _className
            End Get
        End Property
    End Class

    ''' <summary>
    ''' 指定された文字列をウィンドウのタイトルとクラス名に含んでいるプロセスをすべて取得する。
    ''' </summary>
    ''' <param name="windowText">ウィンドウのタイトルに含むべき文字列。
    ''' nullを指定すると、classNameだけで検索する。</param>
    ''' <param name="className">ウィンドウが属するクラス名に含むべき文字列。
    ''' nullを指定すると、windowTextだけで検索する。</param>
    ''' <returns>見つかったプロセスの配列。</returns>
    Public Shared Function GetProcessesByWindow(windowText As String, className As String) As Process()

        Dim processInfo As New FoundProcessInfo(windowText, className)

        'ウィンドウを列挙して、対象のプロセスを探す
        EnumWindows(New EnumWindowsDelegate(AddressOf EnumWindowCallBack), processInfo)

        '結果を返す
        Return DirectCast(processInfo.FoundProcesses.ToArray(GetType(Process)), Process())
    End Function

    Private Shared Function EnumWindowCallBack(hWnd As IntPtr, processInfo As FoundProcessInfo) As Boolean
        If processInfo.WindowText IsNot Nothing Then
            'ウィンドウのタイトルの長さを取得する
            Dim textLen As Integer = GetWindowTextLength(hWnd)
            If textLen = 0 Then
                '次のウィンドウを検索
                Return True
            End If
            'ウィンドウのタイトルを取得する
            Dim tsb As New StringBuilder(textLen + 1)
            GetWindowText(hWnd, tsb, tsb.Capacity)
            'タイトルに指定された文字列を含むか
            If tsb.ToString().IndexOf(processInfo.WindowText) < 0 Then
                '含んでいない時は、次のウィンドウを検索
                Return True
            End If
        End If

        If processInfo.ClassName IsNot Nothing Then
            'ウィンドウのクラス名を取得する
            Dim csb As New StringBuilder(256)
            GetClassName(hWnd, csb, csb.Capacity)
            'クラス名に指定された文字列を含むか
            If csb.ToString().IndexOf(processInfo.ClassName) < 0 Then
                '含んでいない時は、次のウィンドウを検索
                Return True
            End If
        End If

        Dim foundProcessIds As New ArrayList()

        'プロセスのIDを取得する
        Dim processId As Integer
        GetWindowThreadProcessId(hWnd, processId)
        '今まで見つかったプロセスでは無いことを確認する
        If Not foundProcessIds.Contains(processId) Then
            foundProcessIds.Add(processId)
            'プロセスIDをからProcessオブジェクトを作成する
            processInfo.FoundProcesses.Add(Process.GetProcessById(processId))
        End If

        Return True
    End Function

    Private Delegate Function EnumWindowsDelegate(hWnd As IntPtr, processInfo As FoundProcessInfo) As Boolean

    <DllImport("user32.dll")> _
    Private Shared Function EnumWindows(lpEnumFunc As EnumWindowsDelegate, processInfo As FoundProcessInfo) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function GetClassName(hWnd As IntPtr, lpClassName As StringBuilder, nMaxCount As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

    ''' <summary>
    ''' ウィンドウタイトルの長さを取得する
    ''' </summary>
    ''' <param name="hWnd">ウィンドウハンドル</param>
    ''' <returns>ウィンドウタイトルの長さ</returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function GetWindowTextLength(hWnd As IntPtr) As Integer
    End Function

    ''' <summary>
    ''' ウィンドウタイトルを取得する
    ''' </summary>
    ''' <param name="hWnd">ウィンドウハンドル</param>
    ''' <param name="lpString">テキストバッファ</param>
    ''' <param name="nMaxCount">バッファにコピーする文字の最大数</param>
    ''' <returns>コピーされた文字列の最大数</returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxCount As Integer) As Integer
    End Function

    Public Enum OsType
        OTHER_OS
        WINDOWS_95
        WINDOWS_98
        WINDOWS_ME
        WINDOWS_NT_3
        WINDOWS_NT_31
        WINDOWS_NT_35
        WINDOWS_NT_351
        WINDOWS_NT_40
        WINDOWS_2000
        WINDOWS_XP
        WINDOWS_SERVER_2003
        WINDOWS_VISTA
        WINDOWS_7
        WINDOWS_8
        WINDOWS_81
        WINDOWS_10
        WIN_32_S
        WINDOWS_CE
        UNIX
        '.NET Framework 3.5以降
        MAC_OS_X
    End Enum

    ''' <summary>OSタイプ（簡易）</summary>
    Public Enum EzOsType
        OTHER_OS
        XP_OR_2000_OR_2003SERVER
        VISTA_OR_7
    End Enum

    ''' <summary>
    ''' OSタイプを特定する
    ''' </summary>
    ''' <returns>OSタイプ</returns>
    ''' <remarks></remarks>
    Public Shared Function DetectOsType() As OsType
        '一部OSバージョンはそもそも.NET非対応のため、全くのナンセンスです

        'プラットフォームの取得
        Dim os As System.OperatingSystem = System.Environment.OSVersion

        Select Case os.Platform
            Case System.PlatformID.Win32Windows
                ' OSは Windows 95 以降
                If os.Version.Major >= 4 Then
                    Select Case os.Version.Minor
                        Case 0
                            Return OsType.WINDOWS_95
                        Case 10
                            Return OsType.WINDOWS_98
                        Case 90
                            Return OsType.WINDOWS_ME
                    End Select
                End If

            Case System.PlatformID.Win32NT
                ' OSは Windows NT 以降
                Select Case os.Version.Major
                    Case 3
                        Select Case os.Version.Minor
                            Case 0
                                Return OsType.WINDOWS_NT_3
                            Case 1
                                Return OsType.WINDOWS_NT_31
                            Case 5
                                Return OsType.WINDOWS_NT_35
                            Case 51
                                Return OsType.WINDOWS_NT_351
                        End Select

                    Case 4
                        If os.Version.Minor = 0 Then
                            Return OsType.WINDOWS_NT_40
                        End If

                    Case 5
                        Select Case os.Version.Minor
                            Case 0
                                Return OsType.WINDOWS_2000
                            Case 1
                                Return OsType.WINDOWS_XP
                            Case 2
                                Return OsType.WINDOWS_SERVER_2003
                        End Select

                    Case 6
                        Select Case os.Version.Minor
                            Case 0
                                Return OsType.WINDOWS_VISTA
                            Case 1
                                Return OsType.WINDOWS_7
                            Case 2
                                'マニフェストが設定されてないとWindows8.1以降でもOSVersionがWindows8と同じになる
                                Return DetectAfterWindows8()
                            Case 3
                                Return OsType.WINDOWS_81

                        End Select
                    Case 10
                        If os.Version.Minor = 0 Then
                            Return OsType.WINDOWS_10
                        End If

                End Select

            Case System.PlatformID.Win32S
                Return OsType.WIN_32_S

            Case System.PlatformID.WinCE
                Return OsType.WINDOWS_CE

            Case System.PlatformID.Unix
                '.NET Framework 2.0以降
                Return OsType.UNIX

            Case System.PlatformID.Xbox
                '.NET Framework 3.5以降
                Return OsType.OTHER_OS

            Case System.PlatformID.MacOSX
                '.NET Framework 3.5以降
                Return OsType.MAC_OS_X
        End Select

        Return OsType.OTHER_OS
    End Function

    Private Shared Function DetectAfterWindows8() As OsType
        Dim rk As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")

        'CurrentMajorVersionNumberはWindows10から追加された値
        Dim major As Integer? = StringUtil.ToInteger(rk.GetValue("CurrentMajorVersionNumber"))
        If major.HasValue Then
            Dim minor As Integer? = StringUtil.ToInteger(rk.GetValue("CurrentMinorVersionNumber"))
            If major.Value = 10 AndAlso minor.HasValue AndAlso minor.Value = 0 Then
                Return OsType.WINDOWS_10
            End If
            Return OsType.OTHER_OS
        End If

        Dim version As String = StringUtil.ToString(rk.GetValue("CurrentVersion"))
        If "6.3".Equals(version) Then
            Return OsType.WINDOWS_81
        End If

        Return OsType.WINDOWS_8
    End Function

    ''' <summary>
    ''' OSタイプ（簡易）を特定する
    ''' </summary>
    ''' <returns>OSタイプ（簡易）</returns>
    ''' <remarks></remarks>
    Public Shared Function DetectEzOsType() As EzOsType

        'プラットフォームの取得
        Dim os As System.OperatingSystem = System.Environment.OSVersion

        If os.Platform = PlatformID.Win32NT Then
            Select Case os.Version.Major
                Case 5
                    Return EzOsType.XP_OR_2000_OR_2003SERVER
                Case 6
                    Return EzOsType.VISTA_OR_7
            End Select
        End If
        Return EzOsType.OTHER_OS
    End Function

    'SendInput 関数でマウス操作に関する動作等を指定する MOUSEINPUT 構造体
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Private Structure MOUSEINPUT
        ''' <summary>マウス位置の指定座標. 絶対座標/相対座標で指定</summary>
        Public dx As Integer
        ''' <summary></summary>
        Public dy As Integer
        ''' <summary>ホイールの移動量を設定</summary>
        Public mouseData As Integer
        ''' <summary>マウスの動作を指定するフラグを設定</summary>
        Public dwFlags As Integer
        ''' <summary>タイムスタンプ(このメンバは無視されます)</summary>
        Public time As Integer
        ''' <summary>追加情報(このメンバは無視されます)</summary>
        Public dwExtraInfo As IntPtr
    End Structure

    ''' <summary>
    ''' SendInput 関数の設定に使用する INPUT 構造体
    ''' </summary>
    ''' <remarks></remarks>
    <StructLayout(LayoutKind.Sequential, Size:=28)> _
    Private Structure INPUT
        ''' <summary>SendInput 関数の使用目的 0=マウス　1=キーボード　2=ハードウェア</summary>
        Public Type As Integer
        ''' <summary>KEYBDINPUT 構造体</summary>
        Public ki As MOUSEINPUT
    End Structure

    Private Class InputType
        Public Const MOUSE As Integer = 0
        Public Const KEYBOARD As Integer = 1
        Public Const HARDWARE As Integer = 2
    End Class

    Private Class MouseEventFlag
        ''' <summary>マウスを移動する</summary>
        Public Const MOUSE_MOVED As Integer = &H1
        ''' <summary>移動時、絶対座標を指定</summary>
        Public Const MOUSEEVENTF_ABSOLUTE As Integer = &H8000&
        ''' <summary>X ボタンDown</summary>
        Public Const MOUSEEVENTF_XDOWN As Integer = &H100
        ''' <summary>X ボタンUP</summary>
        Public Const MOUSEEVENTF_XUP As Integer = &H200
        ''' <summary>ホイールが回転したことを示し、移動量は、dwData パラメータで指定</summary>
        Public Const MOUSEEVENTF_WHEEL As Integer = &H80
        ''' <summary>左ボタンUP</summary>
        Public Const MOUSEEVENTF_LEFTUP As Integer = &H4
        ''' <summary>左ボタンDown</summary>
        Public Const MOUSEEVENTF_LEFTDOWN As Integer = &H2
        ''' <summary>中央ボタンDown</summary>
        Public Const MOUSEEVENTF_MIDDLEDOWN As Integer = &H20
        ''' <summary>中央ボタンUP</summary>
        Public Const MOUSEEVENTF_MIDDLEUP As Integer = &H40
        ''' <summary>右ボタンDown</summary>
        Public Const MOUSEEVENTF_RIGHTDOWN As Integer = &H8
        ''' <summary>右ボタンUP</summary>
        Public Const MOUSEEVENTF_RIGHTUP As Integer = &H10
    End Class

    Private Class MouseEvent
        Public X As Integer
        Public Y As Integer
        Public EventFlag As Integer
    End Class

    ''' <summary>
    ''' マウス左クリックを送信する
    ''' </summary>
    ''' <param name="x">X座標</param>
    ''' <param name="y">Y座標</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SendMouseLeftClick(ByVal x As Integer, ByVal y As Integer) As Integer
        Dim events As MouseEvent() = {New MouseEvent With {.X = x, .Y = y, .EventFlag = MouseEventFlag.MOUSEEVENTF_LEFTDOWN}, _
                                      New MouseEvent With {.X = x, .Y = y, .EventFlag = MouseEventFlag.MOUSEEVENTF_LEFTUP}}
        Return SedMouseInput(events)
    End Function

    Private Shared Function SedMouseInput(ByVal events As MouseEvent()) As Integer
        Dim eventN As Integer = events.Length - 1
        Dim arrayINPUT(eventN) As INPUT
        For i As Integer = 0 To eventN
            arrayINPUT(i).Type = InputType.MOUSE
            With arrayINPUT(i).ki
                .dx = events(i).X               'マウスの移動時のX方向への移動量
                .dy = events(i).Y               'マウスの移動時のY方向への移動量
                .mouseData = 0                '必要としません
                .dwFlags = events(i).EventFlag        'マウスイベント
                .time = 0                     'デフォルトの設定
                .dwExtraInfo = IntPtr.Zero    '必要としません
            End With
        Next
        '関数の実行(連続でマウスの操作を実施）'個々のマウスの操作の間に割り込みが入らない。
        Return SendInput(arrayINPUT.Length, arrayINPUT, Marshal.SizeOf(GetType(INPUT)))
    End Function

    ''' <summary>
    ''' キーストローク、マウスの動き、ボタンのクリックなどを合成します。
    ''' </summary>
    ''' <param name="nInputs">入力イベントの数</param>
    ''' <param name="pInputs">挿入する入力イベントの配列</param>
    ''' <param name="cbsize">構造体のサイズ</param>
    ''' <returns>挿入することができたイベントの数を返す。ブロックされている場合は 0 を返す</returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Private Shared Function SendInput(ByVal nInputs As Integer, ByVal pInputs As INPUT(), ByVal cbsize As Integer) As Integer
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function OpenPrinter( _
        ByVal pPrinterName As String, ByRef hPrinter As IntPtr, _
        ByVal pDefault As IntPtr) As Boolean
    End Function

    <DllImport("winspool.drv", SetLastError:=True)> _
    Private Shared Function ClosePrinter( _
        ByVal hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.drv", SetLastError:=True)> _
    Private Shared Function GetPrinter( _
        ByVal hPrinter As IntPtr, ByVal dwLevel As Integer, _
        ByVal pPrinter As IntPtr, ByVal cbBuf As Integer, _
        ByRef pcbNeeded As Integer) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Public Structure PRINTER_INFO_2
        Public pServerName As String
        Public pPrinterName As String
        Public pShareName As String
        Public pPortName As String
        Public pDriverName As String
        Public pComment As String
        Public pLocation As String
        Public pDevMode As IntPtr
        Public pSepFile As String
        Public pPrintProcessor As String
        Public pDatatype As String
        Public pParameters As String
        Public pSecurityDescriptor As IntPtr
        Public Attributes As System.UInt32
        Public Priority As System.UInt32
        Public DefaultPriority As System.UInt32
        Public StartTime As System.UInt32
        Public UntilTime As System.UInt32
        Public Status As System.UInt32
        Public cJobs As System.UInt32
        Public AveragePPM As System.UInt32
    End Structure

    ''' <summary>
    ''' プリンタの情報をPRINTER_INFO_2で取得する
    ''' </summary>
    ''' <param name="printerName">プリンタ名</param>
    ''' <returns>プリンタの情報</returns>
    Public Shared Function GetPrinterInfo( _
        ByVal printerName As String) As PRINTER_INFO_2
        'プリンタのハンドルを取得する
        Dim hPrinter As IntPtr
        If Not OpenPrinter(printerName, hPrinter, IntPtr.Zero) Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If

        Dim pPrinterInfo As IntPtr = IntPtr.Zero
        Try
            '必要なバイト数を取得する
            Dim needed As Integer
            GetPrinter(hPrinter, 2, IntPtr.Zero, 0, needed)
            If needed <= 0 Then
                Throw New Exception("失敗しました。")
            End If
            'メモリを割り当てる
            pPrinterInfo = Marshal.AllocHGlobal(needed)

            'プリンタ情報を取得する
            Dim temp As Integer
            If Not GetPrinter(hPrinter, 2, pPrinterInfo, needed, temp) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            'PRINTER_INFO_2型にマーシャリングする
            Dim printerInfo As PRINTER_INFO_2 = _
                CType(Marshal.PtrToStructure( _
                pPrinterInfo, GetType(PRINTER_INFO_2)), PRINTER_INFO_2)

            With printerInfo
                .pServerName = EncodePrinterInfo(.pServerName)
                .pPrinterName = EncodePrinterInfo(.pPrinterName)
                .pShareName = EncodePrinterInfo(.pShareName)
                .pPortName = EncodePrinterInfo(.pPortName)
                .pDriverName = EncodePrinterInfo(.pDriverName)
                .pComment = EncodePrinterInfo(.pComment)
                .pLocation = EncodePrinterInfo(.pLocation)
                .pSepFile = EncodePrinterInfo(.pSepFile)
                .pPrintProcessor = EncodePrinterInfo(.pPrintProcessor)
                .pDatatype = EncodePrinterInfo(.pDatatype)
                .pParameters = EncodePrinterInfo(.pParameters)
            End With

            '結果を返す
            Return printerInfo
        Finally
            '後始末をする
            ClosePrinter(hPrinter)
            Marshal.FreeHGlobal(pPrinterInfo)
        End Try
    End Function

    Private Shared Function EncodePrinterInfo(ByVal value As String) As String
        If value Is Nothing Then
            Return Nothing
        End If
        Return System.Text.Encoding.GetEncoding(932).GetString(System.Text.Encoding.Unicode.GetBytes(value))
    End Function

    ''' <summary>
    ''' システムが和暦設定か？を返す
    ''' </summary>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsJapaneseCalendar() As Boolean
        Return DateUtil.IsJapaneseCalendarOnSystem
    End Function

    ''' <summary>
    ''' .NETで解釈できないシステム日付書式か？を返す
    ''' </summary>
    ''' <returns>判定結果</returns>
    ''' <remarks>
    ''' コントロールパネル | 地域と言語 > 形式タブ | 追加の設定 > 日付タブ | データ形式 短い形式
    ''' ここに「yyyy/MM/dd' 'ddd」を設定するとtrueになる
    ''' </remarks>
    Public Shared Function IsInvalidDateFormat() As Boolean
        Const SAMPLE As Date = #6/17/2019 1:23:45 AM#
        Try
            Return CDate(CStr(SAMPLE)) <> SAMPLE
        Catch ex As Exception
            Return True
        End Try
    End Function

    ''' <summary>
    ''' OSはWin10か？
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function IsWin10() As Boolean
        Return My.Computer.Info.OSFullName.Contains("Windows 10")
    End Function

    ''' <summary>
    ''' OSはWin10でバージョン20H2(ビルド19042)以降か？
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function IsWin10AndAfterVer20H2() As Boolean
        If Not IsWin10() Then
            Return False
        End If
        Const WIN10_20H2_BUILD As Integer = 19042
        Dim build As String = StringUtil.ToString(Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion", False).GetValue("CurrentBuild"))
        Return IsNumeric(build) AndAlso WIN10_20H2_BUILD <= CInt(build)
    End Function

End Class
