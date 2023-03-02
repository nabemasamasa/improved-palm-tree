Imports System.Diagnostics
Imports System.Text

Namespace Sys

    ''' <summary>
    ''' 外部プログラム実行クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Executer

        ''' <summary>Window表示する外部アプリの場合、true</summary>
        Private _IsShowWindowApp As Boolean

        ''' <summary>実行時の表示Window状態</summary>
        Public ProcessWindowStyle As ProcessWindowStyle = ProcessWindowStyle.Hidden

        ''' <summary>プロンプト処理の場合、true  ※初期値true</summary>
        Public IsPrompt As Boolean = True

        ''' <summary>Window表示する外部アプリの場合、true</summary>
        Public Property IsShowWindowApp() As Boolean
            Get
                Return _IsShowWindowApp
            End Get
            Set(ByVal value As Boolean)
                _IsShowWindowApp = value
                If _IsShowWindowApp Then
                    ProcessWindowStyle = Diagnostics.ProcessWindowStyle.Normal
                    IsWaitForExit = False
                    IsPrompt = False
                Else
                    ProcessWindowStyle = Diagnostics.ProcessWindowStyle.Hidden
                    IsWaitForExit = True
                End If
            End Set
        End Property

        ''' <summary>Excelアドインファイルの場合、true</summary>
        Public Property IsExcelXla() As Boolean
            Get
                Return IsShowWindowApp
            End Get
            Set(ByVal value As Boolean)
                IsShowWindowApp = value
            End Set
        End Property

        ''' <summary>実行終了まで待機する場合、true  ※初期値true</summary>
        Public Property IsWaitForExit() As Boolean = True

        ''' <summary>
        ''' 実行中の同名プロセスを終了するか？  ※初期値false<br/>
        ''' ※編集中の画面でも即時終了させるので利用には注意すること
        ''' </summary>
        Public Property IsCloseRunningProcess As Boolean = False

        ''' <summary>結果を出力する場合、true  ※初期値false</summary>
        Public Property UsesRedirectStandardOutput As Boolean = False

        Private _ResultRedirectStandardOutput As String
        ''' <summary> 出力結果</summary>
        Public ReadOnly Property ResultRedirectStandardOutput As String
            Get
                If Not UsesRedirectStandardOutput Then
                    Throw New InvalidOperationException("UsesRedirectStandardOutputをTrueにしないと利用できない")
                End If
                Return _ResultRedirectStandardOutput
            End Get
        End Property

        ''' <summary>
        ''' 外部プログラム実行
        ''' </summary>
        ''' <param name="execName">実行プログラムフルパス指定</param>
        ''' <param name="processArgs">実行時引数</param>
        ''' <param name="workingDirectory">プログラムを起動するディレクトリ</param>
        ''' <param name="verb">ドキュメントを開くときに使用する動詞</param>
        ''' <remarks></remarks>
        Public Sub Exec(ByVal execName As String, _
                        Optional ByVal processArgs As String = "", _
                        Optional ByVal workingDirectory As String = "",
                        Optional verb As String = "")

            If IsCloseRunningProcess Then
                KillProcess(execName)
            End If

            Dim hProcess As New ProcessStartInfo

            '実行exe名
            hProcess.FileName = execName
            If StringUtil.IsNotEmpty(workingDirectory) Then
                hProcess.WorkingDirectory = workingDirectory
            End If

            If processArgs <> "" Then

                '起動引数
                hProcess.Arguments = processArgs.Replace(vbCrLf, Space(1)).Trim
            End If

            'ウインドウ作成
            hProcess.CreateNoWindow = IsPrompt

            'shell起動
            hProcess.UseShellExecute = Not UsesRedirectStandardOutput

            'エラーダイアログ表示
            hProcess.ErrorDialog = True

            '画面状態
            hProcess.WindowStyle = Me.ProcessWindowStyle
            ' Minimizedだと、それにフォーカスが移動してしまい、Asyncで実行していると、処理後にBackgroundUiからフォーカスが外れる事があるので。

            '実行結果を出力する
            hProcess.RedirectStandardOutput = UsesRedirectStandardOutput

            hProcess.Verb = verb

            'プロセス開始
            LogoutCommand(hProcess)
            Dim runProcess As Process = Process.Start(hProcess)

            If UsesRedirectStandardOutput Then
                '出力結果を受け取る
                _ResultRedirectStandardOutput = runProcess.StandardOutput.ReadToEnd
            End If

            If runProcess Is Nothing Then
                ' ファイルやURLだけが、execNameに指定された場合（exeを直接実行しない場合）
                Return
            End If
            Try
                If IsWaitForExit Then
                    'ウエイト
                    runProcess.WaitForExit()

                    If runProcess.ExitCode <> 0 Then
                        Throw New InvalidOperationException("exe起動に失敗.  > " & execName & " " & processArgs)
                    End If
                End If

            Finally
                'クローズ
                runProcess.Close()

                '初期化
                runProcess.Dispose()
            End Try
        End Sub

        ''' <summary>
        ''' プロセスを即時中断する
        ''' </summary>
        ''' <param name="execName">実行プログラム名</param>
        ''' <remarks></remarks>
        Private Sub KillProcess(ByVal execName As String)
            Dim execFileName As String = System.IO.Path.GetFileNameWithoutExtension(execName)
            For Each aProcess As Process In Process.GetProcessesByName(execFileName)
                aProcess.Kill()
                aProcess.WaitForExit()
            Next
        End Sub

        Private Shared Sub LogoutCommand(ByVal hProcess As ProcessStartInfo)
            Const PREFIX As String = "> "
            Const SPACE As String = " "
            Dim sb As New StringBuilder
            sb.Append(PREFIX).Append(hProcess.FileName).Append(SPACE).Append(hProcess.Arguments)
            EzUtil.logDebug(sb.ToString)
        End Sub
    End Class
End Namespace