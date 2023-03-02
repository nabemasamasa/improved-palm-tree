Imports log4net.Core
Imports log4net

''' <summary>
''' 偽装ログクラス
''' </summary>
''' <remarks></remarks>
Public Class FakeLogger : Implements ILog
    Private ReadOnly stockLog As New List(Of String)
    Public ReadOnly Property GetLog() As String()
        Get
            Return stockLog.ToArray()
        End Get
    End Property

    Private Sub AddLog(level As String, message As String)
        stockLog.Add(String.Format("{0, -5} {1}", level, message))
    End Sub

    Public ReadOnly Property Logger() As ILogger Implements ILoggerWrapper.Logger
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Sub Debug(ByVal message As Object) Implements ILog.Debug
        AddLog("DEBUG", message.ToString())
    End Sub

    Public Sub Debug(ByVal message As Object, ByVal exception As Exception) Implements ILog.Debug
        Debug(String.Format("{0}{1}{2}", message, vbCrLf, exception))
    End Sub

    Public Sub DebugFormat(ByVal format As String, ByVal ParamArray args As Object()) Implements ILog.DebugFormat
        Debug(String.Format(format, args))
    End Sub

    Public Sub DebugFormat(ByVal format As String, ByVal arg0 As Object) Implements ILog.DebugFormat
        DebugFormat(format, {arg0})
    End Sub

    Public Sub DebugFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object) Implements ILog.DebugFormat
        DebugFormat(format, {arg0, arg1})
    End Sub

    Public Sub DebugFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object, ByVal arg2 As Object) Implements ILog.DebugFormat
        DebugFormat(format, {arg0, arg1, arg2})
    End Sub

    Public Sub DebugFormat(ByVal provider As IFormatProvider, ByVal format As String, ByVal ParamArray args As Object()) Implements ILog.DebugFormat
        DebugFormat(format, args)
    End Sub

    Public Sub Info(ByVal message As Object) Implements ILog.Info
        AddLog("INFO", message.ToString())
    End Sub

    Public Sub Info(ByVal message As Object, ByVal exception As Exception) Implements ILog.Info
        Info(String.Format("{0}{1}{2}", message, vbCrLf, exception))
    End Sub

    Public Sub InfoFormat(ByVal format As String, ByVal ParamArray args As Object()) Implements ILog.InfoFormat
        Info(String.Format(format, args))
    End Sub

    Public Sub InfoFormat(ByVal format As String, ByVal arg0 As Object) Implements ILog.InfoFormat
        InfoFormat(format, {arg0})
    End Sub

    Public Sub InfoFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object) Implements ILog.InfoFormat
        InfoFormat(format, {arg0, arg1})
    End Sub

    Public Sub InfoFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object, ByVal arg2 As Object) Implements ILog.InfoFormat
        InfoFormat(format, {arg0, arg1, arg2})
    End Sub

    Public Sub InfoFormat(ByVal provider As IFormatProvider, ByVal format As String, ByVal ParamArray args As Object()) Implements ILog.InfoFormat
        InfoFormat(format, args)
    End Sub

    Public Sub Warn(ByVal message As Object) Implements ILog.Warn
        AddLog("WARN", message.ToString())
    End Sub

    Public Sub Warn(ByVal message As Object, ByVal exception As Exception) Implements ILog.Warn
        Warn(String.Format("{0}{1}{2}", message, vbCrLf, exception))
    End Sub

    Public Sub WarnFormat(ByVal format As String, ByVal ParamArray args As Object()) Implements ILog.WarnFormat
        Warn(String.Format(format, args))
    End Sub

    Public Sub WarnFormat(ByVal format As String, ByVal arg0 As Object) Implements ILog.WarnFormat
        WarnFormat(format, {arg0})
    End Sub

    Public Sub WarnFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object) Implements ILog.WarnFormat
        WarnFormat(format, {arg0, arg1})
    End Sub

    Public Sub WarnFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object, ByVal arg2 As Object) Implements ILog.WarnFormat
        WarnFormat(format, {arg0, arg1, arg2})
    End Sub

    Public Sub WarnFormat(ByVal provider As IFormatProvider, ByVal format As String, ByVal ParamArray args As Object()) Implements ILog.WarnFormat
        WarnFormat(format, args)
    End Sub

    Public Sub [Error](ByVal message As Object) Implements ILog.[Error]
        AddLog("ERROR", message.ToString())
    End Sub

    Public Sub [Error](ByVal message As Object, ByVal exception As Exception) Implements ILog.[Error]
        [Error](String.Format("{0}{1}{2}", message, vbCrLf, exception))
    End Sub

    Public Sub ErrorFormat(ByVal format As String, ByVal ParamArray args As Object()) Implements ILog.ErrorFormat
        [Error](String.Format(format, args))
    End Sub

    Public Sub ErrorFormat(ByVal format As String, ByVal arg0 As Object) Implements ILog.ErrorFormat
        ErrorFormat(format, {arg0})
    End Sub

    Public Sub ErrorFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object) Implements ILog.ErrorFormat
        ErrorFormat(format, {arg0, arg1})
    End Sub

    Public Sub ErrorFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object, ByVal arg2 As Object) Implements ILog.ErrorFormat
        ErrorFormat(format, {arg0, arg1, arg2})
    End Sub

    Public Sub ErrorFormat(ByVal provider As IFormatProvider, ByVal format As String, ByVal ParamArray args As Object()) Implements ILog.ErrorFormat
        ErrorFormat(format, args)
    End Sub

    Public Sub Fatal(ByVal message As Object) Implements ILog.Fatal
        AddLog("FATAL", message.ToString())
    End Sub

    Public Sub Fatal(ByVal message As Object, ByVal exception As Exception) Implements ILog.Fatal
        Fatal(String.Format("{0}{1}{2}", message, vbCrLf, exception))
    End Sub

    Public Sub FatalFormat(ByVal format As String, ByVal ParamArray args As Object()) Implements ILog.FatalFormat
        Fatal(String.Format(format, args))
    End Sub

    Public Sub FatalFormat(ByVal format As String, ByVal arg0 As Object) Implements ILog.FatalFormat
        FatalFormat(format, {arg0})
    End Sub

    Public Sub FatalFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object) Implements ILog.FatalFormat
        FatalFormat(format, {arg0, arg1})
    End Sub

    Public Sub FatalFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object, ByVal arg2 As Object) Implements ILog.FatalFormat
        FatalFormat(format, {arg0, arg1, arg2})
    End Sub

    Public Sub FatalFormat(ByVal provider As IFormatProvider, ByVal format As String, ByVal ParamArray args As Object()) Implements ILog.FatalFormat
        FatalFormat(format, args)
    End Sub

    Public ReadOnly Property IsDebugEnabled() As Boolean Implements ILog.IsDebugEnabled
        Get
            Return True
        End Get
    End Property

    Public ReadOnly Property IsInfoEnabled() As Boolean Implements ILog.IsInfoEnabled
        Get
            Return True
        End Get
    End Property

    Public ReadOnly Property IsWarnEnabled() As Boolean Implements ILog.IsWarnEnabled
        Get
            Return True
        End Get
    End Property

    Public ReadOnly Property IsErrorEnabled() As Boolean Implements ILog.IsErrorEnabled
        Get
            Return True
        End Get
    End Property

    Public ReadOnly Property IsFatalEnabled() As Boolean Implements ILog.IsFatalEnabled
        Get
            Return True
        End Get
    End Property
End Class
