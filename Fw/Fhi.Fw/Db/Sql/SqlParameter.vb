Imports System.Text

Namespace Db.Sql
    Public Class SqlParameter
#Region "Nested Classes..."
        Private Class DbParameterVo
            Friend ReadOnly ParamName As String
            Friend ReadOnly Value As Object
            Friend ReadOnly DbType As DbType?
            ''' <summary>
            ''' コンストラクタ
            ''' </summary>
            ''' <param name="paramName2">パラメータ名</param>
            ''' <param name="value">パラメータ値</param>
            ''' <param name="dbtype">パラメータ型</param>
            ''' <remarks></remarks>
            Friend Sub New(ByVal paramName2 As String, ByVal value As Object, ByVal dbType As DbType?)
                Me.ParamName = paramName2
                Me.Value = value
                Me.DbType = dbType
            End Sub

            Public Overrides Function ToString() As String
                Return String.Format("paramName = {0}, value = {1}, dbType = {2}", ParamName, Value, DbType.ToString)
            End Function
            Public Function ToLogDebugFormat() As String
                Dim sb As New StringBuilder
                sb.Append(ParamName).Append(":")
                If Value Is Nothing OrElse Value.ToString Is Nothing Then
                    sb.Append("<null>")
                ElseIf Value.GetType.IsEnum Then
                    sb.AppendFormat("{0} ({1})", CInt(Value), Value.ToString)
                Else
                    sb.Append(Value.ToString)
                End If
                Return sb.ToString
            End Function
        End Class
        Public Interface IBehavior
            Sub LogDebug(message As String)
        End Interface
        Private Class DefaultBehavior : Implements IBehavior
            Public Sub LogDebug(message As String) Implements IBehavior.LogDebug
                DbAccessHelper.logDebug(message)
            End Sub
        End Class
#End Region

        Private ReadOnly parameterNames As New List(Of String)
        Private ReadOnly parameters As New List(Of DbParameterVo)
        Private ReadOnly behavior As IBehavior

        Public Sub New()
            Me.New(Nothing)
        End Sub

        Public Sub New(behavior As IBehavior)
            Me.behavior = If(behavior, New DefaultBehavior)
        End Sub

        ''' <summary>
        ''' パラメーターを追加する
        ''' </summary>
        ''' <param name="paramName">パラメーター名</param>
        ''' <param name="value">パラメーター値</param>
        ''' <remarks></remarks>
        Public Sub Add(ByVal paramName As String, ByVal value As Object)
            Call Add(paramName, value, Nothing)
        End Sub

        ''' <summary>
        ''' パラメーターを追加する
        ''' </summary>
        ''' <param name="paramName">パラメーター名</param>
        ''' <param name="value">パラメーター値</param>
        ''' <param name="dbtype">パラメータ型</param>
        ''' <remarks></remarks>
        Public Sub Add(ByVal paramName As String, ByVal value As Object, ByVal dbtype As DbType?)
            parameterNames.Add(paramName)
            parameters.Add(New DbParameterVo(paramName, value, dbtype))
        End Sub

        ''' <summary>
        ''' パラメータをクリアする
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Clear()
            parameterNames.Clear()
            parameters.Clear()
        End Sub

        ''' <summary>パラメータの件数</summary>
        ''' <returns>パラメータの件数</returns>
        Public ReadOnly Property Count() As Integer
            Get
                Return parameters.Count
            End Get
        End Property

        ''' <summary>
        ''' パラメータ名が含まれているかを返す
        ''' </summary>
        ''' <param name="paramName">パラメータ名</param>
        ''' <returns>含まれている場合、true</returns>
        ''' <remarks></remarks>
        Public Function ContainsName(ByVal paramName As String) As Boolean
            Return parameterNames.Contains(paramName)
        End Function

        ''' <summary>
        ''' パラメータを SqlAccess に追加設定する
        ''' </summary>
        ''' <param name="db">パラメータを追加する SqlAccess</param>
        ''' <remarks></remarks>
        Public Sub AddAllTo(ByVal db As DbAccess)
            For Each param As DbParameterVo In parameters
                behavior.LogDebug(param.ToLogDebugFormat)
                If param.DbType.HasValue Then
                    db.AddParameter(param.ParamName, param.Value, param.DbType.Value)
                Else
                    db.AddParameter(param.ParamName, param.Value)
                End If
            Next
        End Sub

        ''' <summary>
        ''' パラメータ情報を List(Of String) に追加する
        ''' </summary>
        ''' <param name="results">パラメータ情報を追加する List</param>
        ''' <remarks></remarks>
        Public Sub AddAllTo(ByVal results As List(Of String))
            For Each vo As DbParameterVo In parameters
                results.Add(vo.ToString)
            Next
        End Sub

        ''' <summary>
        ''' SQL文にバインドパラメータを直接書き込む
        ''' </summary>
        ''' <param name="sql"></param>
        ''' <returns>埋め込み後のSQL文</returns>
        ''' <remarks></remarks>
        Public Function BindToSql(ByVal sql As String) As String
            Dim result As String = sql
            Dim aParams As New List(Of DbParameterVo)(parameters)
            aParams.Sort(Function(a, b) a.ParamName.CompareTo(b.ParamName))
            For subtrahend As Integer = 1 To aParams.Count
                Dim param As DbParameterVo = aParams(aParams.Count - subtrahend)
                If param.Value Is Nothing Then
                    result = result.Replace(param.ParamName, "NULL")
                ElseIf TypeOf param.Value Is String Then
                    result = result.Replace(param.ParamName, String.Format("'{0}'", SqlUtil.EscapeParameter(param.Value.ToString)))
                ElseIf TypeOf param.Value Is DateTime OrElse TypeOf param.Value Is DateTime? Then
                    result = result.Replace(param.ParamName, String.Format("'{0}'", DateUtil.ToFixedString(DirectCast(param.Value, DateTime?))))
                Else
                    result = result.Replace(param.ParamName, param.Value.ToString)
                End If
            Next
            Return result
        End Function
    End Class
End Namespace
