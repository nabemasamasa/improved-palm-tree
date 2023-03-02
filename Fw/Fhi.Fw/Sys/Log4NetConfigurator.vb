Imports System.IO
Imports log4net.Appender

Namespace Sys
    ''' <summary>
    ''' Log4Netの設定を担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Log4NetConfigurator

        ''' <summary>MDCに持たせるUserIdキー</summary>
        Private Const MDC_KEY_USER_ID As String = "id"

        Private Shared ReadOnly _instance As New Log4NetConfigurator

        Private ReadOnly rfa As New log4net.Appender.RollingFileAppender
        Private ca As AppenderSkeleton = New ConsoleAppender

        Private Sub New()
        End Sub

        ''' <summary>
        ''' 唯一のインスタンスを返す
        ''' </summary>
        ''' <returns>インスタンス</returns>
        ''' <remarks></remarks>
        Public Shared Function GetInstance() As Log4NetConfigurator
            Return _instance
        End Function

#Region "Public properties..."
        ''' <summary>ユーザーIDを出力する場合、true ※初期値 false</summary>
        Public IsOutputUserId As Boolean = False

        ''' <summary>ファイル出力</summary>
        Public ReadOnly Property RollingFileAppender() As RollingFileAppender
            Get
                Return rfa
            End Get
        End Property

        ''' <summary>コンソール出力</summary>
        Public Property ConsoleAppender() As AppenderSkeleton
            Get
                Return ca
            End Get
            Set(ByVal value As AppenderSkeleton)
                ca = value
            End Set
        End Property
#End Region

        ''' <summary>
        ''' 初期化する
        ''' </summary>
        ''' <param name="aLevel">出力レベル</param>
        ''' <remarks></remarks>
        Public Sub Initialize(ByVal aLevel As log4net.Core.Level, ByVal assemblyName As String)
            Initialize(aLevel, assemblyName, AssemblyUtil.GetPath)
        End Sub

        ''' <summary>
        ''' 初期化する
        ''' </summary>
        ''' <param name="aLevel">出力レベル</param>
        ''' <remarks></remarks>
        Public Sub Initialize(ByVal aLevel As log4net.Core.Level, ByVal assemblyName As String, ByVal aPath As String)

            ' RollingFileAppender

            rfa.Threshold = aLevel
            rfa.Layout = New log4net.Layout.PatternLayout("%d{yyyy-MM-dd HH:mm:ss.fff} %-5p " & If(IsOutputUserId, "[%X{" & MDC_KEY_USER_ID & "}] ", "") & "%m%n")
            rfa.Name = "simple_Log4netForConsole"

            rfa.ImmediateFlush = True

            rfa.File = Path.Combine(aPath, Path.GetFileNameWithoutExtension(assemblyName) & ".log")

            rfa.AppendToFile = True
            rfa.Encoding = System.Text.Encoding.UTF8
            rfa.SecurityContext = log4net.Core.SecurityContextProvider.DefaultProvider.CreateSecurityContext(Nothing)
            rfa.LockingModel = New log4net.Appender.FileAppender.ExclusiveLock

            rfa.MaxFileSize = 1024 * 1024
            rfa.DatePattern = "'.'yyyy---m-m--dd"
            rfa.MaxSizeRollBackups = 4
            rfa.CountDirection = -1
            rfa.RollingStyle = log4net.Appender.RollingFileAppender.RollingMode.Size
            rfa.StaticLogFileName = False

            rfa.ActivateOptions()

            log4net.Config.BasicConfigurator.Configure(rfa)


            'ConsoleAppender

            ca.Threshold = aLevel
            ca.Layout = New log4net.Layout.PatternLayout("%d{HH:mm:ss.fff} %-5p " & If(IsOutputUserId, "[%X{" & MDC_KEY_USER_ID & "}] ", "") & "%m%n")
            log4net.Config.BasicConfigurator.Configure(ca)

        End Sub

        ''' <summary>
        ''' コンソール出力の出力レイアウトを設定する
        ''' </summary>
        ''' <param name="pattern">出力レイアウト</param>
        ''' <remarks></remarks>
        Public Sub SetLayoutToConsoleAppender(ByVal pattern As String)
            ca.Layout = New log4net.Layout.PatternLayout(pattern)
        End Sub

        ''' <summary>
        ''' MDCへUserIdを設定する
        ''' </summary>
        ''' <param name="userId">UserId</param>
        ''' <remarks></remarks>
        Public Shared Sub SetUserId(ByVal userId As String)
            log4net.MDC.Set(MDC_KEY_USER_ID, userId)
        End Sub
    End Class

End Namespace