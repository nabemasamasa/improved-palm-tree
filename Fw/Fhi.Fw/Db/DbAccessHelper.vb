Imports System.Reflection
Imports System.Data.Common
Imports Fhi.Fw.Db.Sql
Imports System.Text
Imports Fhi.Fw.Domain

Namespace Db
    Public Class DbAccessHelper

        Private Class Holder
            Public Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(GetType(DbAccessHelper))
        End Class

        Private Shared ReadOnly TYPE_OF_PRIMITIVE_VALUE_OBJECT As Type = GetType(PrimitiveValueObject)

#Region "Public properties..."
        ''' <summary>デバッグ出力にデータベース名を出力する場合、true</summary>
        Public Shared IsShowDatabaseName As Boolean
        Private _linkServerName As String
        ''' <summary>リンクサーバー経由の場合、リンクサーバー名を設定する</summary>
        Public Property LinkServerName As String
            Get
                Return _linkServerName
            End Get
            Set(value As String)
                _linkServerName = value
                If StringUtil.IsNotEmpty(value) AndAlso IsShowDatabaseName Then
                    DirectCast(logger, DebugDbShow).DatabaseName = value
                End If
            End Set
        End Property

        ''' <summary>埋め込みパラメータ使用を許可しない場合、true ※初期値=false</summary>
        Public DenyBindParameter As Boolean = False

        ''' <summary>強制的に項目値をTrimする場合、true ※初期値=false</summary>
        Public TrimsEndOfValueToForce As Boolean = False

        ''' <summary>OPENQUERYを使用する場合、true</summary>
        Public DenyOpenQuery As Boolean = False
#End Region
#Region "Nested classes..."
        Private Interface IAdapter
            Function ContainsName(ByVal columnName As String) As Boolean
            Function GetFirstValue() As Object
            Function GetValue(ByVal columnName As String) As Object
        End Interface

        Private Class DataReaderAdapter : Implements IAdapter

            Private ReadOnly aDataReader As DbDataReader
            Private columnNames As List(Of String)

            Public Sub New(ByVal aDataReader As DbDataReader)
                Me.aDataReader = aDataReader
            End Sub

            Public Function ContainsName(ByVal columnName As String) As Boolean Implements IAdapter.ContainsName
                If columnNames Is Nothing Then
                    columnNames = New List(Of String)
                    For i As Integer = 0 To aDataReader.FieldCount - 1
                        columnNames.Add(aDataReader.GetName(i).ToUpper)
                    Next
                End If
                Return columnNames.Contains(columnName)
            End Function

            Public Overridable Function GetFirstValue() As Object Implements IAdapter.GetFirstValue
                Return aDataReader.Item(0)
            End Function

            Public Overridable Function GetValue(ByVal columnName As String) As Object Implements IAdapter.GetValue
                Return aDataReader.Item(columnName)
            End Function
        End Class

        Private Class DataReaderAdapterOnTrim : Inherits DataReaderAdapter

            Public Sub New(ByVal aDataReader As DbDataReader)
                MyBase.New(aDataReader)
            End Sub

            Public Overrides Function GetFirstValue() As Object
                Return TrimEndIfString(MyBase.GetFirstValue)
            End Function

            Public Overrides Function GetValue(ByVal columnName As String) As Object
                Return TrimEndIfString(MyBase.GetValue(columnName))
            End Function
        End Class

        Private Class DataTableAdapter : Implements IAdapter

            Public aDataRow As DataRow

            Private ReadOnly aDataColumns As DataColumnCollection

            Public Sub New(ByVal aDataColumns As DataColumnCollection)
                Me.aDataColumns = aDataColumns
            End Sub

            Public Function ContainsName(ByVal columnName As String) As Boolean Implements IAdapter.ContainsName
                Return aDataColumns.Contains(columnName)
            End Function

            Public Overridable Function GetFirstValue() As Object Implements IAdapter.GetFirstValue
                Return aDataRow(0)
            End Function

            Public Overridable Function GetValue(ByVal columnName As String) As Object Implements IAdapter.GetValue
                Return aDataRow(columnName)
            End Function
        End Class

        Private Class DataTableAdapterOnTrim : Inherits DataTableAdapter

            Public Sub New(ByVal aDataColumns As DataColumnCollection)
                MyBase.New(aDataColumns)
            End Sub

            Public Overrides Function GetFirstValue() As Object
                Return TrimEndIfString(MyBase.GetFirstValue)
            End Function

            Public Overrides Function GetValue(ByVal columnName As String) As Object
                Return TrimEndIfString(MyBase.GetValue(columnName))
            End Function
        End Class

        Private Interface ILogDebug
            ''' <summary>
            ''' SQLをデバッグ出力する
            ''' </summary>
            ''' <param name="sql">出力内容</param>
            ''' <remarks></remarks>
            Sub DebugSql(ByVal sql As String)

            ''' <summary>
            ''' DB設定をデバッグ出力する
            ''' </summary>
            ''' <param name="command">出力内容</param>
            ''' <remarks></remarks>
            Sub DebugSetting(ByVal command As String)

            ''' <summary>
            ''' プロシージャをデバッグ出力する
            ''' </summary>
            ''' <param name="procedureName">出力内容</param>
            ''' <remarks></remarks>
            Sub DebugProcedure(ByVal procedureName As String)
        End Interface

        Private Class DebugDbShow : Implements ILogDebug

            Friend DatabaseName As String

            Public Sub New(ByVal databaseName As String)
                Me.DatabaseName = databaseName
            End Sub

            Public Sub DebugSql(ByVal sql As String) Implements ILogDebug.DebugSql
                PerformLogout(String.Empty, sql)
            End Sub

            Public Sub DebugSetting(ByVal command As String) Implements ILogDebug.DebugSetting
                PerformLogout(String.Empty, command)
            End Sub

            Public Sub DebugProcedure(ByVal procedureName As String) Implements ILogDebug.DebugProcedure
                Const prefix2 As String = "(PROCEDURE)"
                PerformLogout(prefix2, procedureName)
            End Sub

            Private Sub PerformLogout(ByVal prefix As String, ByVal message As String)
                Dim sb As New StringBuilder(DatabaseName)
                sb.Append(prefix).Append(">").Append(message)
                logDebug(sb.ToString)
            End Sub
        End Class

        Private Class DebugDbDontShow : Implements ILogDebug

            Public Sub DebugSql(ByVal sql As String) Implements ILogDebug.DebugSql
                Const Prefix As String = "SQL:"
                PerformLogout(Prefix, sql)
            End Sub

            Public Sub DebugSetting(ByVal command As String) Implements ILogDebug.DebugSetting
                PerformLogout(String.Empty, command)
            End Sub

            Public Sub DebugProcedure(ByVal procedureName As String) Implements ILogDebug.DebugProcedure
                Const Prefix As String = "PROCEDURE:"
                PerformLogout(Prefix, procedureName)
            End Sub

            Private Sub PerformLogout(ByVal prefix As String, ByVal message As String)
                Dim sb As New StringBuilder()
                sb.Append(prefix).Append(message)
                logDebug(sb.ToString)
            End Sub
        End Class

#End Region

        Private ReadOnly db As DbAccess
        Private ReadOnly logger As ILogDebug

        Public Sub New(ByVal db As DbAccess)
            Me.db = db
            Me.logger = If(IsShowDatabaseName, DirectCast(New DebugDbShow(db.Database), ILogDebug), DirectCast(New DebugDbDontShow, ILogDebug))
        End Sub

        ''' <summary>
        ''' 参照クエリを実行する
        ''' </summary>
        ''' <typeparam name="T">戻り値の型</typeparam>
        ''' <param name="sql">参照クエリ</param>
        ''' <param name="param">パラメータ</param>
        ''' <param name="irregularDbFieldNameByPropertyName">Camelizeルール対象外のルール</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function QueryForList(Of T)(ByVal sql As String, ByVal param As Object, ByVal irregularDbFieldNameByPropertyName As Dictionary(Of String, String)) As List(Of T)

            AssertAnonymous(param, "param")
            Dim sqlAnalyzer As New SqlAnalyzer(sql, param, db.DbTypeOfString)
            sqlAnalyzer.Analyze()

            Dim newSql As String
            If String.IsNullOrEmpty(LinkServerName) AndAlso Not DenyBindParameter Then
                newSql = sqlAnalyzer.AnalyzedSql
            ElseIf Not String.IsNullOrEmpty(LinkServerName) Then
                Dim boundSql As String = sqlAnalyzer.MakeParametersBoundSql
                If DenyOpenQuery Then
                    newSql = (New LinkServerSql(LinkServerName)).ConvertFourPartNamesFrom(boundSql)
                Else
                    newSql = (New LinkServerSql(LinkServerName)).ConvertOpenqueryFrom(boundSql)
                End If
            Else 'If DenyBindParameter Then
                newSql = sqlAnalyzer.MakeParametersBoundSql
            End If
            logger.DebugSql(newSql)
            If String.IsNullOrEmpty(LinkServerName) AndAlso Not DenyBindParameter Then
                sqlAnalyzer.AddParametersTo(db)
            End If

            Dim dtData As New DataTable
            Try
                db.Fill(newSql, dtData)
            Catch ex As Exception
                HandleException(ex)
                Throw
            End Try
            Return ConvDataTableToList(Of T)(dtData, irregularDbFieldNameByPropertyName, TrimsEndOfValueToForce)
        End Function

        ''' <summary>
        ''' 更新クエリを実行する
        ''' </summary>
        ''' <param name="sql">参照クエリ</param>
        ''' <param name="param">パラメータ</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Execute(ByVal sql As String, ByVal param As Object) As Integer
            AssertAnonymous(param, "param")
            Dim sqlAnalyzer As New SqlAnalyzer(sql, param, db.DbTypeOfString)
            sqlAnalyzer.Analyze()

            Dim resolvedSql As String
            If String.IsNullOrEmpty(LinkServerName) AndAlso Not DenyBindParameter Then
                resolvedSql = sqlAnalyzer.AnalyzedSql
                logger.DebugSql(resolvedSql)
                sqlAnalyzer.AddParametersTo(db)
            Else
                resolvedSql = sqlAnalyzer.MakeParametersBoundSql
                If Not String.IsNullOrEmpty(LinkServerName) Then
                    If DenyOpenQuery Then
                        resolvedSql = (New LinkServerSql(LinkServerName)).ConvertFourPartNamesFrom(resolvedSql)
                    Else
                        resolvedSql = (New LinkServerSql(LinkServerName)).ConvertOpenqueryFrom(resolvedSql)
                    End If
                End If
                logger.DebugSql(resolvedSql)
            End If
            Try
                Return db.ExecuteNonQuery(resolvedSql)
            Catch ex As Exception
                HandleException(ex)
                Throw
            End Try
        End Function

        ''' <summary>
        ''' プロシージャ実行のパラメータを追加する
        ''' </summary>
        ''' <param name="paramaterName">パラメータ名</param>
        ''' <param name="value">値</param>
        ''' <param name="direction">パラメータの型</param>
        ''' <param name="size">桁数</param>
        ''' <remarks></remarks>
        Public Sub AddProcedureParameter(ByVal paramaterName As String, ByVal value As Object, ByVal direction As ParameterDirection, ByVal size As Integer)
            db.AddParameter(paramaterName, value, direction:=direction, size:=size)
        End Sub

        ''' <summary>
        ''' プロシージャを実行する
        ''' </summary>
        ''' <param name="procedureName">プロシージャ名</param>
        ''' <returns>結果[]</returns>
        ''' <remarks></remarks>
        Public Function ExecuteProcedure(ByVal procedureName As String) As List(Of Object)
            logger.DebugProcedure(procedureName)
            Try
                Return db.ExecuteProcedure(procedureName)
            Catch ex As Exception
                HandleException(ex)
                Throw
            End Try
        End Function

        Private Sub AssertAnonymous(ByVal value As Object, ByVal paramName As String)

            If value Is Nothing Then
                Return
            End If
            If TypeUtil.IsTypeAnonymous(value) Then
                Throw New ArgumentException("匿名型の使用は許可できません.", paramName)
            End If
        End Sub

        Public Sub ExecuteSetting(ByVal sql As String)
            logger.DebugSetting(sql)
            Try
                db.ExecuteNonQuery(sql)
            Catch ex As Exception
                HandleException(ex)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' 例外を処理する
        ''' </summary>
        ''' <param name="ex">例外</param>
        ''' <remarks></remarks>
        Private Sub HandleException(ByVal ex As Exception)

            If IsShowDatabaseName Then
                Holder.log.Error(If(StringUtil.IsEmpty(LinkServerName), db.Database, LinkServerName) & ">" & db.LastSql)
            Else
                Holder.log.Error(db.LastSql)
            End If
            For Each dbParam As System.Data.Common.DbParameter In db.LastParams
                Holder.log.Error(dbParam.ParameterName & ":" & If(dbParam.Value, "<null>").ToString)
            Next
        End Sub

        ''' <summary>
        ''' String型ならTrimEndする
        ''' </summary>
        ''' <param name="value">値</param>
        ''' <returns>String型ならTrimEndされた値。違う場合はその値</returns>
        ''' <remarks></remarks>
        Private Shared Function TrimEndIfString(ByVal value As Object) As Object
            If TypeOf value Is String Then
                Return DirectCast(value, String).TrimEnd
            End If
            Return value
        End Function

        ''' <summary>
        ''' キャメライズしたDataRowの列名に一致するVOのプロパティに値を設定して返す
        ''' </summary>
        ''' <typeparam name="T">戻り値のVOの型（キャメライズしたプロパティを持つこと）</typeparam>
        ''' <remarks></remarks>
        Private Class ConvDataReaderToVo(Of T)

            Private Shared ReadOnly TYPE_OF_STRING As Type = GetType(String)
            Private ReadOnly voType As Type
            Private ReadOnly decamelizeByCamelize As Dictionary(Of String, String)
            Private ReadOnly irregularDbFieldNameByPropertyName As Dictionary(Of String, String)
            Private ReadOnly typeIfNullableByDecamelize As Dictionary(Of String, Type)

            ''' <summary>
            ''' コンストラクタ
            ''' </summary>
            ''' <param name="irregularDbFieldNameByPropertyName">キャメライズしたDataTable列名とは違うプロパティに値を設定する場合のルール</param>
            ''' <remarks></remarks>
            Public Sub New(ByVal irregularDbFieldNameByPropertyName As Dictionary(Of String, String))
                Me.voType = GetType(T)
                Me.irregularDbFieldNameByPropertyName = irregularDbFieldNameByPropertyName

                If Not TypeUtil.IsTypeImmutable(voType) Then
                    AssertCreateInstance()
                End If

                Me.decamelizeByCamelize = New Dictionary(Of String, String)
                Me.typeIfNullableByDecamelize = New Dictionary(Of String, Type)
                For Each voProperty As PropertyInfo In voType.GetProperties()
                    Dim decamelize As String = If(irregularDbFieldNameByPropertyName.ContainsKey(voProperty.Name), _
                                                                    irregularDbFieldNameByPropertyName(voProperty.Name), _
                                                                    StringUtil.Decamelize(voProperty.Name))
                    Dim result As String = If(decamelize.StartsWith("[") AndAlso decamelize.EndsWith("]"), _
                                                                    decamelize.Substring(1, decamelize.Length - 2), _
                                                                    decamelize)
                    decamelizeByCamelize.Add(voProperty.Name, result)
                    typeIfNullableByDecamelize.Add(result, TypeUtil.GetTypeIfNullable(voProperty.PropertyType))
                Next
            End Sub

            ''' <summary>
            ''' 型タイプの新しいインスタンスを作成できる事を保証する
            ''' </summary>
            ''' <remarks></remarks>
            Private Sub AssertCreateInstance()
                Try
                    Dim newInstance As Object = Activator.CreateInstance(voType)
                Catch ex As MissingMethodException
                    Throw New NotSupportedException(voType.Name & " を使用するには引数無しコンストラクタが必要", ex)
                End Try
            End Sub

            ''' <summary>
            ''' 新しいインスタンスを作成する
            ''' </summary>
            ''' <returns>新しいインスタンス</returns>
            ''' <remarks></remarks>
            Private Function CreateInstance() As T

                Return CType(Activator.CreateInstance(voType), T)
            End Function

            ''' <summary>
            ''' キャメライズしたDataRowの列名に一致するVOのプロパティに値を設定して返す
            ''' </summary>
            ''' <param name="adapter">変換元のDataRow値</param>
            ''' <returns>キャメライズルールに従ってDataRowの値を設定したVO</returns>
            ''' <remarks></remarks>
            Public Function Compute(ByVal adapter As IAdapter) As T
                If voType.IsValueType OrElse voType Is TYPE_OF_STRING _
                        OrElse TYPE_OF_PRIMITIVE_VALUE_OBJECT.IsAssignableFrom(voType) Then
                    Dim dbValue As Object = ResolveValue(adapter.GetFirstValue, voType,
                                                         voPropertyName:=Nothing, columnName:=Nothing)
                    If dbValue Is Nothing Then
                        Return Nothing
                    End If
                    Return DirectCast(dbValue, T)
                End If

                Dim result As T = CreateInstance()
                For Each voProperty As PropertyInfo In voType.GetProperties()
                    Dim columnName As String = decamelizeByCamelize(voProperty.Name)
                    If Not adapter.ContainsName(columnName) Then
                        Continue For
                    End If
                    Dim dbValue As Object = adapter.GetValue(columnName)
                    voProperty.SetValue(result, ResolveValue(dbValue, typeIfNullableByDecamelize(columnName),
                                                             voProperty.Name, columnName), Nothing)
                Next
                Return result
            End Function

            Private Function ResolveValue(dbValue As Object, voPropertyType As Type, voPropertyName As String, _
                                          columnName As String) As Object
                If TypeOf dbValue Is System.DBNull AndAlso Not TYPE_OF_PRIMITIVE_VALUE_OBJECT.IsAssignableFrom(voPropertyType) Then
                    Return Nothing
                End If
                If dbValue.GetType Is voPropertyType Then
                    Return dbValue
                ElseIf voPropertyType Is TYPE_OF_STRING Then
                    Return dbValue.ToString
                ElseIf voPropertyType Is GetType(Decimal) OrElse voPropertyType Is GetType(Int32) OrElse voPropertyType Is GetType(Int64) OrElse voPropertyType Is GetType(Int16) OrElse voPropertyType Is GetType(Byte) Then
                    ' ルール1. DBがDoubleでも、精度を省みて、.NET上では、Decimalで扱う
                    ' ルール2. DBがNumeric(4,0)は、voPropertyがDecimal/Doubleだが、VO上はInt32にしている
                    If IsNumeric(dbValue) Then
                        If voPropertyType Is GetType(Decimal) Then
                            Return Convert.ToDecimal(dbValue)
                        Else
                            Dim decValue As Decimal = Convert.ToDecimal(dbValue)
                            If Decimal.Truncate(decValue) <> decValue Then
                                AssertInvalidPropertyTypeAndDbType(columnName, dbValue, voPropertyName, voPropertyType)
                            End If
                            If voPropertyType Is GetType(Int32) Then
                                Return Convert.ToInt32(dbValue)
                            ElseIf voPropertyType Is GetType(Int64) Then
                                Return Convert.ToInt64(dbValue)
                            ElseIf voPropertyType Is GetType(Int16) Then
                                Return Convert.ToInt16(dbValue)
                            ElseIf voPropertyType Is GetType(Byte) Then
                                Return Convert.ToByte(dbValue)
                            End If
                        End If
                    ElseIf StringUtil.IsNotEmpty(dbValue) Then
                        AssertInvalidPropertyTypeAndDbType(columnName, dbValue, voPropertyName, voPropertyType)
                    End If
                ElseIf voPropertyType Is GetType(Boolean) Then
                    Return EzUtil.IsTrue(dbValue)
                ElseIf voPropertyType Is GetType(DateTime) Then
                    Return CDate(dbValue)
                ElseIf TYPE_OF_PRIMITIVE_VALUE_OBJECT.IsAssignableFrom(voPropertyType) Then
                    ' ルール3. DB値がnullならPrimitiveValueObjectの中身をnullにする
                    Dim constructor As ConstructorInfo = ValueObject.DetectConstructor(voPropertyType)
                    Dim parameterInfos As ParameterInfo() = constructor.GetParameters
                    If parameterInfos.Length <> 1 Then
                        Throw New InvalidOperationException("引数のないPrimitiveValueObjectは対応できない")
                    End If
                    Return constructor.Invoke(New Object() {If(TypeOf dbValue Is System.DBNull, Nothing,
                                                               ResolveValue(dbValue,
                                                                            TypeUtil.GetTypeIfNullable(parameterInfos(0).ParameterType),
                                                                            voPropertyName, columnName))})
                ElseIf voPropertyType.IsEnum Then
                    Return [Enum].ToObject(voPropertyType, If(TypeOf dbValue Is Boolean,
                                                              If(DirectCast(dbValue, Boolean), 1, 0),
                                                              dbValue))
                Else
                    Return dbValue
                End If
                AssertInvalidPropertyTypeAndDbType(columnName, dbValue, voPropertyName, voPropertyType)
                Throw New InvalidProgramException
            End Function

            Private Sub AssertInvalidPropertyTypeAndDbType(ByVal columnName As String, ByVal dbValue As Object, ByVal voPropertyName As String, ByVal voPropertyType As Type)

                Throw New ArgumentException(String.Format("列名 {0} の値は {1} だが、VO.{2} は {3} 型", columnName, StringUtil.ToString(dbValue), voPropertyName, voPropertyType.Name))
            End Sub
        End Class


        ''' <summary>
        ''' キャメライズしたDataTableの列名に一致する VOのプロパティに値を設定して、Listにして返す
        ''' </summary>
        ''' <typeparam name="T">戻り値のVOの型（キャメライズしたプロパティを持つこと）</typeparam>
        ''' <param name="aDataTable">変換元のDataTable値</param>
        ''' <returns>キャメライズルールに従ってDataTableの値を設定したVOのList</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvDataTableToList(Of T)(ByVal aDataTable As DataTable) As List(Of T)
            Return ConvDataTableToList(Of T)(aDataTable, Nothing)
        End Function

        ''' <summary>
        ''' DataTableの列名に一致する VOのプロパティ（キャメライズ）に値を設定して、Listにして返す
        ''' </summary>
        ''' <typeparam name="T">戻り値のVOの型（キャメライズしたプロパティを持つこと）</typeparam>
        ''' <param name="aDataTable">変換元のDataTable値</param>
        ''' <param name="irregularDbFieldNameByPropertyName">キャメライズしたDataTable列名とは違うプロパティに値を設定する場合のルール</param>
        ''' <param name="trimsEndOfValueToForce">項目値の末尾を強制的にTrimする場合、true</param>
        ''' <returns>キャメライズルールに従ってDataTableの値を設定したVOのList</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvDataTableToList(Of T)(ByVal aDataTable As DataTable, ByVal irregularDbFieldNameByPropertyName As Dictionary(Of String, String), _
                                                         Optional ByVal trimsEndOfValueToForce As Boolean = False) As List(Of T)
            Dim camelizeIrregulars As Dictionary(Of String, String) = If(irregularDbFieldNameByPropertyName, New Dictionary(Of String, String))
            Dim convertToVo As New ConvDataReaderToVo(Of T)(camelizeIrregulars)
            Dim results As New List(Of T)
            Dim adapter As DataTableAdapter = If(trimsEndOfValueToForce, New DataTableAdapterOnTrim(aDataTable.Columns), New DataTableAdapter(aDataTable.Columns))
            For Each adapter.aDataRow In aDataTable.Rows
                results.Add(convertToVo.Compute(adapter))
            Next
            Return results
        End Function

        ''' <summary>
        ''' DataReaderの列名に一致する VOのプロパティ（キャメライズ）に値を設定して、Listにして返す
        ''' </summary>
        ''' <typeparam name="T">戻り値のVOの型（キャメライズしたプロパティを持つこと）</typeparam>
        ''' <param name="aDataReader">変換元のDataReader値</param>
        ''' <param name="irregularDbFieldNameByPropertyName">キャメライズしたDataTable列名とは違うプロパティに値を設定する場合のルール</param>
        ''' <param name="trimsEndOfValueToForce">項目値の末尾を強制的にTrimする場合、true</param>
        ''' <returns>キャメライズルールに従ってDataTableの値を設定したVOのList</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvDataReaderToList(Of T)(ByVal aDataReader As DbDataReader, ByVal irregularDbFieldNameByPropertyName As Dictionary(Of String, String), _
                                                          Optional ByVal trimsEndOfValueToForce As Boolean = False) As List(Of T)
            Dim results As New List(Of T)
            If Not aDataReader.HasRows Then
                Return results
            End If
            Dim camelizeIrregulars As Dictionary(Of String, String) = If(irregularDbFieldNameByPropertyName, New Dictionary(Of String, String))
            Dim convertToVo As New ConvDataReaderToVo(Of T)(camelizeIrregulars)
            Dim adapter As IAdapter = If(trimsEndOfValueToForce, New DataReaderAdapterOnTrim(aDataReader), New DataReaderAdapter(aDataReader))
            While aDataReader.Read
                results.Add(convertToVo.Compute(adapter))
            End While
            Return results
        End Function

        ''' <summary>
        ''' コンソール出力
        ''' </summary>
        ''' <param name="message">出力メッセージ</param>
        ''' <remarks></remarks>
        Friend Shared Sub logDebug(ByVal message As String)
            EzUtil.logDebug(message)
        End Sub

        ''' <summary>
        ''' 引数の型に一致するテーブルの列名[]を取得する
        ''' </summary>
        ''' <typeparam name="T">取得したいテーブルのVo/Daoの型</typeparam>
        ''' <returns>列名[]</returns>
        ''' <remarks></remarks>
        Public Function FetchTableColumns(Of T)() As String()
            Dim tableName As String = SqlUtil.DeriveTableName(Of T)()
            Return FetchTableColumns(tableName)
        End Function

        ''' <summary>
        ''' テーブルに存在する列名[]を取得する
        ''' </summary>
        ''' <param name="tableName">取得したいテーブル名</param>
        ''' <returns>列名[]</returns>
        ''' <remarks></remarks>
        Public Function FetchTableColumns(tableName As String) As String()
            Dim sql As String = String.Format("SELECT * FROM {0} WHERE 0 = 1", tableName)
            Dim dtData As New DataTable
            Try
                db.Fill(sql, dtData)
            Catch ex As Exception
                HandleException(ex)
                Throw
            End Try
            Return ExtractColumnNames(dtData)
        End Function

        ''' <summary>
        ''' DataTableから列名[]を抽出する
        ''' </summary>
        ''' <param name="aDataTable"></param>
        ''' <returns>列名[]</returns>
        ''' <remarks></remarks>
        Public Shared Function ExtractColumnNames(aDataTable As DataTable) As String()
            Dim columns As DataColumnCollection = aDataTable.Columns
            Return Enumerable.Range(0, columns.Count).Select(Function(index) columns.Item(index).ColumnName).ToArray
        End Function

    End Class
End Namespace
