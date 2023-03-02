Imports Fhi.Fw.Domain
Imports System.Text

Namespace Db.Sql
    ''' <summary>
    ''' "@Hoge"というバインド文字列に該当するプロパティ値を抽出する
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SqlBindAnalyzer
        ''' <summary>Int32/DateTime/Stringなどがパラメータだった時の、埋め込みパラメータ名</summary>
        Public Const VALUE_BIND_PARAM_NAME As String = "@Value"

        Private ReadOnly sql As String
        Private ReadOnly bindParamValue As Object
        Private ReadOnly parameter As New SqlParameter
        Private ReadOnly dbTypeOfString As DbType

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="sql">SQL(バインド文字有)</param>
        ''' <param name="bindParamValue">バインドする値、もしくはプロパティ値としてもつ値</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal sql As String, ByVal bindParamValue As Object)
            Me.New(sql, bindParamValue, DbType.AnsiString)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="sql">SQL(バインド文字有)</param>
        ''' <param name="bindParamValue">バインドする値、もしくはプロパティ値としてもつ値</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal sql As String, ByVal bindParamValue As Object, ByVal dbTypeOfString As DbType)
            Me.sql = sql
            Me.bindParamValue = bindParamValue
            Me.dbTypeOfString = dbTypeOfString
        End Sub

        ''' <summary>
        ''' "@Value"に埋め込むパラメータ型かどうかを返す
        ''' </summary>
        ''' <returns>該当するパラメータ型なら、true</returns>
        ''' <remarks></remarks>
        Private Function IsValueBindParameter() As Boolean
            If bindParamValue Is Nothing Then
                Return False
            End If
            Dim paramType As Type = bindParamValue.GetType
            Return paramType.IsValueType OrElse paramType Is GetType(String)
        End Function

        Private Function UsesCriteriaParameter() As Boolean
            Return TypeOf bindParamValue Is CriteriaBinder
        End Function

        ''' <summary>
        ''' プロパティ値を抽出する為に、SQL文を解析する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Analyze()
            Dim chars As Char() = (sql & " ").ToCharArray
            Dim dqing As Boolean
            Dim sqing As Boolean
            Dim bindNameBuilder As New StringBuilder
            Dim bindNaming As Boolean
            parameter.Clear()

            Dim IsValueBindParam As Boolean = IsValueBindParameter()
            Dim usesCriteria As Boolean = UsesCriteriaParameter()
            For Each c As Char In chars
                If bindNaming AndAlso SqlBindUtil.IsInternalBindNameChar(c) Then
                    bindNameBuilder.Append(c)
                Else
                    If bindNaming Then
                        Dim bindName As String = bindNameBuilder.ToString
                        If bindParamValue Is Nothing Then
                            Throw New ArgumentException("バインド変数名 " & bindName & " を検知したが、そもそもバインドする値がnull.")
                        End If
                        Dim value As Object
                        If IsValueBindParam Then
                            If Not VALUE_BIND_PARAM_NAME.Equals(bindName) Then
                                Throw New ArgumentException("バインドする値の型 " & bindParamValue.GetType.Name & " は、@Valueを使用すべき. " & bindName & " は使えない.")
                            End If
                            value = bindParamValue
                        ElseIf usesCriteria Then
                            value = DirectCast(bindParamValue, CriteriaBinder).GetValueByIdentifyName(SqlBindUtil.ConvBindNameToPropertyName(bindName))
                        Else
                            value = VoUtil.GetValue(bindParamValue, SqlBindUtil.ConvBindNameToPropertyName(bindName))
                        End If
                        If TypeOf value Is PrimitiveValueObject Then
                            value = DirectCast(value, PrimitiveValueObject).Value
                        End If
                        If Not parameter.ContainsName(bindName) Then
                            If TypeOf value Is String Then
                                '          |SqlServer|CE|  AS400 LinkServer  |
                                'AnsiString|   OK    |NG|         OK         |
                                'String    |   OK    |OK|NG(SQL結果がtimeout)|
                                parameter.Add(bindName, value, dbTypeOfString)
                                'parameter.Add(bindName, value, DbType.String)
                            Else
                                parameter.Add(bindName, value)
                            End If
                        End If
                        bindNameBuilder = Nothing
                        bindNaming = False
                    End If
                    If c = """"c Then
                        dqing = Not dqing
                    ElseIf c = "'"c Then
                        sqing = Not sqing
                    ElseIf Not dqing AndAlso Not sqing AndAlso c = "@"c Then
                        bindNaming = True
                        bindNameBuilder = New StringBuilder
                        bindNameBuilder.Append(c)
                    End If
                End If
            Next
        End Sub

        ''' <summary>
        ''' パラメータを SqlAccess に追加設定する
        ''' </summary>
        ''' <param name="db">パラメータを追加する SqlAccess</param>
        ''' <remarks></remarks>
        Public Sub AddAllTo(ByVal db As DbAccess)
            parameter.AddAllTo(db)
        End Sub

        ''' <summary>
        ''' パラメータ情報を List(Of String) に追加する
        ''' </summary>
        ''' <param name="aList">パラメータ情報を追加する List</param>
        ''' <remarks></remarks>
        Public Sub AddAllTo(ByVal aList As List(Of String))
            parameter.AddAllTo(aList)
        End Sub

        ''' <summary>
        ''' バインドパラメータを直接書き込んだSQL文を作成する
        ''' </summary>
        ''' <returns>埋め込み後のSQL文</returns>
        ''' <remarks></remarks>
        Public Function MakeParametersBoundSql() As String
            Return parameter.BindToSql(sql)
        End Function
    End Class
End Namespace