Imports System.Linq.Expressions
Imports System.Reflection
Imports System.Text
Imports Fhi.Fw.Db.Sql

Namespace Db
    Public Class SqlUtil

#Region "Nested classes..."
        Private Class WrappedSelectionField(Of T) : Implements SelectionField(Of T)
            Private ReadOnly specifier As VoUtil.IPropertySpecifier
            Private ReadOnly fieldVo As T
            Public Sub New(specifier As VoUtil.IPropertySpecifier, fieldVo As T)
                Me.specifier = specifier
                Me.fieldVo = fieldVo
            End Sub
            Public Function [Is](ParamArray fields As Object()) As SelectionField(Of T) Implements SelectionField(Of T).[Is]
                If CollectionUtil.IsNotEmpty(fields) AndAlso fields.Length = 1 AndAlso fields(0) Is DirectCast(fieldVo, Object) Then
                    VoUtil.GetPropertyNames(Of T).ToList.ForEach(Sub(name) specifier.AppendProperties(VoUtil.GetValue(fieldVo, name)))
                Else
                    specifier.AppendProperties(fields)
                End If
                Return Me
            End Function
            Public Function [Is](Of F)(aField As Expression(Of Func(Of F))) As SelectionField(Of T) Implements SelectionField(Of T).[Is]
                specifier.AppendPropertyWithFunc(Of F)(aField)
                Return Me
            End Function
        End Class
#End Region

        ''' <summary>
        ''' Null値以外のプロパティで、Where句を作成して返す
        ''' </summary>
        ''' <param name="clause">検索条件となるVo</param>
        ''' <returns>Null値以外のプロパティをもとに作成したWhere句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeWhereClause(ByVal clause As Object) As String
            Return MakeWhereClause(clause, Nothing)
        End Function

        ''' <summary>
        ''' Where句を作成して返す
        ''' </summary>
        ''' <param name="clause">検索条件となるVo</param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns>プロパティをもとに作成したWhere句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeWhereClause(ByVal clause As Object, ByVal specifiedFieldNames As Dictionary(Of String, String)) As String
            Dim results As New List(Of String)
            If clause Is Nothing Then
                Return String.Empty
            End If
            Dim aType As Type = clause.GetType
            For Each aProperty As PropertyInfo In aType.GetProperties()
                Dim value As Object = aProperty.GetValue(clause, Nothing)
                If value IsNot Nothing Then
                    results.Add(MakePartOfClause(aProperty.Name, specifiedFieldNames))
                End If
            Next
            Return FormatWhereClause(results)
        End Function

        ''' <summary>
        ''' VOのプロパティ名から、"DBの項目名 = @プロパティ名" を作成する
        ''' </summary>
        ''' <param name="propertyName">プロパティ名</param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function MakePartOfClause(ByVal propertyName As String, ByVal specifiedFieldNames As Dictionary(Of String, String)) As String
            If specifiedFieldNames IsNot Nothing AndAlso specifiedFieldNames.ContainsKey(propertyName) Then
                Return specifiedFieldNames(propertyName) & " = @" & propertyName
            End If
            Return StringUtil.Decamelize(propertyName) & " = @" & propertyName
        End Function

        ''' <summary>
        ''' onlyFieldsの項目たちで、Where句を作成して返す
        ''' </summary>
        ''' <param name="clause">検索条件となるVo</param>
        ''' <param name="onlyFields">検索条件に含めるプロパティ情報</param>
        ''' <returns>onlyFieldsをもとに作成したWhere句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeWhereClauseOnly(ByVal clause As Object, ByVal onlyFields As List(Of PropertyInfo)) As String
            Return MakeWhereClauseOnly(clause, onlyFields, Nothing)
        End Function
        ''' <summary>
        ''' onlyFieldsの項目たちで、Where句を作成して返す
        ''' </summary>
        ''' <param name="clause">検索条件となるVo</param>
        ''' <param name="onlyFields">検索条件に含めるプロパティ情報</param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns>onlyFieldsをもとに作成したWhere句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeWhereClauseOnly(ByVal clause As Object, ByVal onlyFields As List(Of PropertyInfo), ByVal specifiedFieldNames As Dictionary(Of String, String)) As String
            Dim results As New List(Of String)
            For Each aProperty As PropertyInfo In onlyFields
                results.Add(MakePartOfClause(aProperty.Name, specifiedFieldNames))
            Next
            Return FormatWhereClause(results)
        End Function

        ''' <summary>
        ''' Where句を作成して返す
        ''' </summary>
        ''' <param name="clauseList">条件式を保持した List</param>
        ''' <returns>Where句</returns>
        ''' <remarks></remarks>
        Private Shared Function FormatWhereClause(ByVal clauseList As List(Of String)) As String
            If clauseList.Count = 0 Then
                Return String.Empty
            End If
            Return " WHERE " & Join(clauseList.ToArray, " AND ")
        End Function

        ''' <summary>
        ''' Where句を作成して返す
        ''' </summary>
        ''' <param name="criteria">検索条件</param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns>プロパティをもとに作成したWhere句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeWhereClauseByCriteria(criteria As CriteriaBinder, specifiedFieldNames As Dictionary(Of String, String)) As String
            Dim results As List(Of String) = criteria.GetParameters.Select(Function(param) MakePartOfClause(param, specifiedFieldNames)).ToList()
            Return " WHERE " & Join(results.ToArray, " ")
        End Function

        ''' <summary>
        ''' 検索条件情報から、SQLの条件（"DBの項目名 = @プロパティ名"など） を作成する
        ''' </summary>
        ''' <param name="param"></param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function MakePartOfClause(param As CriteriaParameter, ByVal specifiedFieldNames As Dictionary(Of String, String)) As String
            Dim columnName As String = Nothing
            If param.PropertyName Is Nothing Then
                'nop
            ElseIf specifiedFieldNames IsNot Nothing AndAlso specifiedFieldNames.ContainsKey(param.PropertyName) Then
                columnName = specifiedFieldNames(param.PropertyName)
            Else
                columnName = StringUtil.Decamelize(param.PropertyName)
            End If
            Select Case param.Condition
                Case CriteriaParameter.SearchCondition.Equal
                    If param.Value Is Nothing Then
                        Return String.Format("{0} IS NULL", columnName)
                    End If
                    Return String.Format("{0} = @{1}", columnName, param.IdentifyName)
                Case CriteriaParameter.SearchCondition.Any
                    Return String.Format("{0} IN (<join property='@{1}' separator=',' />)", columnName, param.IdentifyName)
                Case CriteriaParameter.SearchCondition.[Like]
                    Return String.Format("{0} LIKE @{1}", columnName, param.IdentifyName)
                Case CriteriaParameter.SearchCondition.GreaterThan
                    Return String.Format("@{0} < {1}", param.IdentifyName, columnName)
                Case CriteriaParameter.SearchCondition.LessThan
                    Return String.Format("{0} < @{1}", columnName, param.IdentifyName)
                Case CriteriaParameter.SearchCondition.GreaterEqual
                    Return String.Format("@{0} <= {1}", param.IdentifyName, columnName)
                Case CriteriaParameter.SearchCondition.LessEqual
                    Return String.Format("{0} <= @{1}", columnName, param.IdentifyName)
                Case CriteriaParameter.SearchCondition.And
                    Return "AND"
                Case CriteriaParameter.SearchCondition.AndBracket
                    Return "AND ("
                Case CriteriaParameter.SearchCondition.Or
                    Return "OR"
                Case CriteriaParameter.SearchCondition.OrBracket
                    Return "OR ("
                Case CriteriaParameter.SearchCondition.Not
                    Return "NOT"
                Case CriteriaParameter.SearchCondition.NotBracket
                    Return "NOT ("
                Case CriteriaParameter.SearchCondition.Bracket
                    Return "("
                Case CriteriaParameter.SearchCondition.CloseBracket
                    Return ")"
            End Select
            Throw New NotImplementedException(String.Format("検索条件'{0}'は未実装", param.Condition))
        End Function

        ''' <summary>
        ''' Insert文の、INTO句とVALUES句を返す
        ''' </summary>
        ''' <typeparam name="T">テーブルVOの型</typeparam>
        ''' <param name="parameterObj">テーブルVO</param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns>Insert文の、INTO句とVALUES句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeInsertInto(Of T)(ByVal parameterObj As T, Optional ByVal specifiedFieldNames As Dictionary(Of String, String) = Nothing) As String
            Return PerformMakeInsertInto(Of T)(parameterObj, specifiedFieldNames, isIgnoreValue:=Function(value As Object) False)
        End Function

        ''' <summary>
        ''' Insert文の、INTO句とVALUES句を返す（但し、値がnullの項目は除外する）
        ''' </summary>
        ''' <typeparam name="T">テーブルVOの型</typeparam>
        ''' <param name="parameterObj">テーブルVO</param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns>Insert文の、INTO句とVALUES句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeInsertIntoWithoutNull(Of T)(ByVal parameterObj As T, Optional ByVal specifiedFieldNames As Dictionary(Of String, String) = Nothing) As String
            Return PerformMakeInsertInto(Of T)(parameterObj, specifiedFieldNames, isIgnoreValue:=Function(value As Object) value Is Nothing)
        End Function

        ''' <summary>除外する値かを返す</summary>
        Private Delegate Function IsIgnoreValueCallback(ByVal value As Object) As Boolean

        ''' <summary>
        ''' Insert文の、INTO句とVALUES句を返す（但し、値がnullの項目は除外する）
        ''' </summary>
        ''' <typeparam name="T">テーブルVOの型</typeparam>
        ''' <param name="parameter">テーブルVO</param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <param name="isIgnoreValue">除外する値か？のCallback</param>
        ''' <returns>Insert文の、INTO句とVALUES句</returns>
        ''' <remarks></remarks>
        Private Shared Function PerformMakeInsertInto(Of T)(parameter As T, specifiedFieldNames As Dictionary(Of String, String), _
                                                            isIgnoreValue As IsIgnoreValueCallback) As String
            Dim columns As New List(Of String)
            Dim nameMatrix As New List(Of List(Of String))
            Dim parameters As Object()
            Dim aType As Type = GetType(T)
            If aType.IsArray Then
                aType = aType.GetElementType
                Dim parameterArray As Array = DirectCast(DirectCast(parameter, Object), Array)
                parameters = Enumerable.Range(0, parameterArray.Length).Select(Function(i) parameterArray.GetValue(i)).ToArray()
            Else
                parameters = {parameter}
            End If
            Dim specifiedFieldNames2 As Dictionary(Of String, String) = If(specifiedFieldNames, New Dictionary(Of String, String))
            Dim propertyInfos = aType.GetProperties()
            For Each aProperty As PropertyInfo In propertyInfos
                Dim info As PropertyInfo = aProperty
                Dim firstIgnoreResult As Boolean = isIgnoreValue.Invoke(info.GetValue(parameters(0), Nothing))
                If parameters.Any(Function(param) firstIgnoreResult <> isIgnoreValue.Invoke(info.GetValue(param, Nothing))) Then
                    Throw New ArgumentException(String.Format("プロパティ名 {0} に「列無視の値」と「有効な値」が混在している. SQLを組み立てできない", info.Name))
                End If
                If firstIgnoreResult Then
                    Continue For
                End If
                If specifiedFieldNames2.ContainsKey(aProperty.Name) Then
                    columns.Add(specifiedFieldNames2(aProperty.Name))
                Else
                    columns.Add(StringUtil.Decamelize(aProperty.Name))
                End If
            Next
            For i = 0 To parameters.Length - 1
                Dim parameterObj As Object = parameters(i)
                Dim values As New List(Of String)
                For Each aProperty As PropertyInfo In propertyInfos
                    If isIgnoreValue.Invoke(aProperty.GetValue(parameterObj, Nothing)) Then
                        Continue For
                    End If
                    If 1 < parameters.Length Then
                        values.Add(SqlBindUtil.ConvPropertyNameToInternalBindName(
                            SqlBindUtil.ConvPropertyNameToInternalBindName("Value", i) & "." & aProperty.Name))
                    Else
                        values.Add(SqlBindUtil.ConvPropertyNameToInternalBindName(aProperty.Name))
                    End If
                Next
                nameMatrix.Add(values)
            Next
            If columns.Count = 0 Then
                Throw New ArgumentException("対象項目が無いのでinsertできない")
            End If
            Return "(" & Join(columns.ToArray, ", ") & ") VALUES " & Join(nameMatrix.Select(Function(values) "(" & Join(values.ToArray, ", ") & ")").ToArray, ", ")
        End Function

        ''' <summary>
        ''' Update文のSET句とWHERE句を返す
        ''' </summary>
        ''' <typeparam name="T">テーブルVOの型</typeparam>
        ''' <param name="clauseValue">検索条件値と更新値をもつテーブルVO</param>
        ''' <param name="whereFields">検索条件となるプロパティ</param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns>Update文のSET句とWHERE句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeUpdateSetWhere(Of T)(ByVal clauseValue As T, ByVal whereFields As List(Of PropertyInfo), _
                                                        Optional ByVal specifiedFieldNames As Dictionary(Of String, String) = Nothing) As String
            Return PerformMakeUpdateSetWithoutNullAndWhere(clauseValue, whereFields, specifiedFieldNames, _
                                                           isIgnoreValue:=Function(value As Object) False)
        End Function

        ''' <summary>
        ''' Update文のSET句（但し、値がnullの項目は除外する）とWHERE句を返す
        ''' </summary>
        ''' <typeparam name="T">テーブルVOの型</typeparam>
        ''' <param name="clauseValue">検索条件値と更新値をもつテーブルVO</param>
        ''' <param name="whereFields">検索条件となるプロパティ</param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns>Update文のSET句とWHERE句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeUpdateSetWithoutNullAndWhere(Of T)(ByVal clauseValue As T, ByVal whereFields As List(Of PropertyInfo), _
                                                                      Optional ByVal specifiedFieldNames As Dictionary(Of String, String) = Nothing) As String
            Return PerformMakeUpdateSetWithoutNullAndWhere(clauseValue, whereFields, specifiedFieldNames, _
                                                           isIgnoreValue:=Function(value As Object) value Is Nothing)
        End Function

        ''' <summary>
        ''' Update文のSET句（但し、値がnullの項目は除外する）とWHERE句を返す
        ''' </summary>
        ''' <typeparam name="T">テーブルVOの型</typeparam>
        ''' <param name="clauseValue">検索条件値と更新値をもつテーブルVO</param>
        ''' <param name="whereFields">検索条件となるプロパティ</param>
        ''' <param name="specifiedFieldNames">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <param name="isIgnoreValue">除外する値か？のCallback</param>
        ''' <returns>Update文のSET句とWHERE句</returns>
        ''' <remarks></remarks>
        Private Shared Function PerformMakeUpdateSetWithoutNullAndWhere(Of T)(ByVal clauseValue As T, ByVal whereFields As List(Of PropertyInfo), _
                                                                              ByVal specifiedFieldNames As Dictionary(Of String, String), _
                                                                              ByVal isIgnoreValue As IsIgnoreValueCallback) As String
            Dim sets As New List(Of String)
            Dim aType As Type = GetType(T)
            Dim specifiedFieldNames2 As Dictionary(Of String, String) = If(specifiedFieldNames, New Dictionary(Of String, String))
            For Each aProperty As PropertyInfo In aType.GetProperties()
                If 0 <= whereFields.IndexOf(aProperty) Then
                    Continue For
                End If
                If isIgnoreValue.Invoke(aProperty.GetValue(clauseValue, Nothing)) Then
                    Continue For
                End If
                Dim dbFieldName As String
                If specifiedFieldNames2.ContainsKey(aProperty.Name) Then
                    dbFieldName = specifiedFieldNames2(aProperty.Name)
                Else
                    dbFieldName = StringUtil.Decamelize(aProperty.Name)
                End If
                sets.Add(dbFieldName & " = @" & aProperty.Name)
            Next
            Return Join(sets.ToArray, ", ") & MakeWhereClauseOnly(clauseValue, whereFields, specifiedFieldNames2)
        End Function

        ''' <summary>
        ''' SQLパラメータを直接設定可能なエスケープした値にして返す
        ''' </summary>
        ''' <param name="param">パラメータ値</param>
        ''' <returns>エスケープ後のパラメータ値</returns>
        ''' <remarks></remarks>
        Public Shared Function EscapeParameter(ByVal param As String) As String
            If param Is Nothing Then
                Return Nothing
            End If
            Return param.Replace("'", "''")
        End Function

        ''' <summary>
        ''' SQLパラメータを直接設定可能なエスケープした値にして返す
        ''' </summary>
        ''' <param name="params">パラメータ値[]</param>
        ''' <returns>エスケープした値[]</returns>
        ''' <remarks></remarks>
        Public Shared Function EscapeParameters(ByVal params As ICollection(Of String)) As String()
            Dim result As New List(Of String)
            For Each param As String In params
                result.Add(EscapeParameter(param))
            Next
            Return result.ToArray
        End Function

        ''' <summary>
        ''' Like演算子の@Paramに埋め込むパラメータ用にエスケープ処理をする
        ''' </summary>
        ''' <param name="param"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EscapeForBindParameterOfLike(ByVal param As String) As String
            Dim result As String = param
            If 0 <= param.IndexOf("%"c) Then
                result = result.Replace("%", "[%]")
            End If
            If 0 <= param.IndexOf("_"c) Then
                result = result.Replace("_", "[_]")
            End If
            If 0 <= param.IndexOf("["c) Then
                result = result.Replace("[", "[[]")
            End If
            Return result
        End Function

        ''' <summary>
        ''' 部分一致のワイルドカード処理をする
        ''' </summary>
        ''' <param name="param">パラメータ値</param>
        ''' <returns>部分一致のワイルドカード処理したパラメータ</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvToPartialMatchWildcard(ByVal param As String) As String
            Return PerformConvToWildcard(param, Function(param2) "%" & param2 & "%")
        End Function

        ''' <summary>
        ''' 前方一致のワイルドカード処理をする
        ''' </summary>
        ''' <param name="param">パラメータ値</param>
        ''' <returns>前方一致のワイルドカード処理したパラメータ</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvToForwardMatchWildcard(ByVal param As String) As String
            Return PerformConvToWildcard(param, Function(param2) param2 & "%")
        End Function

        ''' <summary>
        ''' 後方一致のワイルドカード処理をする
        ''' </summary>
        ''' <param name="param">パラメータ値</param>
        ''' <returns>後方一致のワイルドカード処理したパラメータ</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvToBackwardMatchWildcard(ByVal param As String) As String
            Return PerformConvToWildcard(param, Function(param2) "%" & param2)
        End Function

        Private Delegate Function ConvToWildcardCallback(Of T, TResult)(ByVal arg As t) As TResult

        ''' <summary>
        ''' ワイルドカード処理をする ※ * をワイルドカードとして扱う. '*' を検索するには '\*' を指定する.
        ''' </summary>
        ''' <param name="param">パラメータ値</param>
        ''' <param name="convToWildcardCallback"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function PerformConvToWildcard(ByVal param As String, ByVal convToWildcardCallback As ConvToWildcardCallback(Of String, String)) As String
            EzUtil.AssertParameterIsNotEmpty(param, "param")

            Const ASTERISK As Char = "*"c
            Const PERCENT As Char = "%"c
            Const YEN As Char = "\"c

            Dim sb As New StringBuilder
            Dim escapedParamChars As Char() = EscapeForBindParameterOfLike(param).ToCharArray
            Dim i As Integer = 0
            While i < escapedParamChars.Length
                If YEN = escapedParamChars(i) AndAlso i < escapedParamChars.Length - 1 AndAlso ASTERISK = escapedParamChars(i + 1) Then
                    sb.Append(ASTERISK)
                    i += 2
                    Continue While
                End If
                If ASTERISK = escapedParamChars(i) Then
                    sb.Append(PERCENT)
                Else
                    sb.Append(escapedParamChars(i))
                End If
                i += 1
            End While

            Dim result As String = sb.ToString
            Dim percentCount As Integer = 0
            For Each c As Char In result
                If c = "%"c Then
                    percentCount += 1
                End If
            Next
            Dim escapedPercentCount As Integer = 0
            For j As Integer = 0 To result.Length - 3
                If "[%]".Equals(result.Substring(j, 3)) Then
                    escapedPercentCount += 1
                End If
            Next
            If percentCount <> escapedPercentCount Then
                Return result
            End If
            Return convToWildcardCallback(result)
        End Function

        ''' <summary>
        ''' Vo/Daoからテーブル名を導出する
        ''' </summary>
        ''' <typeparam name="T">Vo/Daoの型</typeparam>
        ''' <returns>テーブル名</returns>
        ''' <remarks></remarks>
        Public Shared Function DeriveTableName(Of T)() As String
            Dim typeName As String = GetType(T).Name
            Dim suffixLength As Integer
            If typeName.EndsWith("Vo") Then
                suffixLength = 2
            ElseIf typeName.EndsWith("Dao") Then
                suffixLength = 3
            ElseIf typeName.EndsWith("DaoImpl") Then
                suffixLength = 7
            Else
                Throw New ArgumentException(String.Format("{0} からテーブル名を導出できない. 未対応", typeName))
            End If
            Return StringUtil.DecamelizeIgnoreNumber(typeName.Substring(0, typeName.Length - suffixLength))
        End Function

        ''' <summary>
        ''' CREATE TABLE文を返す
        ''' </summary>
        ''' <typeparam name="T">テーブルVOの型</typeparam>
        ''' <returns>Update文のSET句とWHERE句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeDropTable(Of T)() As String
            Return String.Format("DROP TABLE {0}", DeriveTableName(Of T))
        End Function

        ''' <summary>
        ''' CREATE TABLE文を返す
        ''' </summary>
        ''' <typeparam name="T">テーブルVOの型</typeparam>
        ''' <returns>Update文のSET句とWHERE句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeCreateTableForSQLite(Of T)() As String
            Return MakeCreateTableForSQLite(Of T)(Nothing)
        End Function

        ''' <summary>
        ''' CREATE TABLE文を返す
        ''' </summary>
        ''' <typeparam name="T">テーブルVOの型</typeparam>
        ''' <param name="pkFields">検索条件となるプロパティ</param>
        ''' <returns>Update文のSET句とWHERE句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeCreateTableForSQLite(Of T)(ByVal pkFields As List(Of PropertyInfo)) As String
            Return MakeCreateTableForSQLite(Of T)(pkFields, DeriveTableName(Of T))
        End Function

        ''' <summary>
        ''' CREATE TABLE文を返す
        ''' </summary>
        ''' <typeparam name="T">テーブルVOの型</typeparam>
        ''' <param name="pkFields">検索条件となるプロパティ</param>
        ''' <returns>Update文のSET句とWHERE句</returns>
        ''' <remarks></remarks>
        Public Shared Function MakeCreateTableForSQLite(Of T)(ByVal pkFields As List(Of PropertyInfo), ByVal tableName As String) As String
            Dim pkDefines As New List(Of String)
            If pkFields IsNot Nothing Then
                For Each info As PropertyInfo In pkFields
                    pkDefines.Add(StringUtil.Decamelize(info.Name))
                Next
            End If
            Dim fieldDefines As New List(Of String)
            Dim aType As Type = GetType(T)
            For Each aProperty As PropertyInfo In aType.GetProperties()
                Dim voPropertyType As Type = TypeUtil.GetTypeIfNullable(aProperty.PropertyType)
                If voPropertyType Is GetType(String) Then
                    fieldDefines.Add(StringUtil.Decamelize(aProperty.Name) & " TEXT")
                ElseIf voPropertyType Is GetType(Integer) OrElse voPropertyType Is GetType(Long) Then
                    fieldDefines.Add(StringUtil.Decamelize(aProperty.Name) & " INTEGER")
                ElseIf voPropertyType Is GetType(Decimal) Then
                    fieldDefines.Add(StringUtil.Decamelize(aProperty.Name) & " REAL")
                ElseIf voPropertyType Is GetType(DateTime) Then
                    fieldDefines.Add(StringUtil.Decamelize(aProperty.Name) & " TEXT")
                ElseIf voPropertyType.IsEnum Then
                    fieldDefines.Add(StringUtil.Decamelize(aProperty.Name) & " INTEGER")
                Else
                    Throw New ArgumentException(String.Format("VO({0})のプロパティ {1} の {2} 型は未対応", aType.Name, aProperty.Name, voPropertyType.Name))
                End If
            Next
            Dim result As String = String.Format("CREATE TABLE {0} ({1})", tableName, Join(fieldDefines.ToArray, ", "))
            If 0 < pkDefines.Count Then
                result &= String.Format(" PRIMARY KEY ({0})", Join(pkDefines.ToArray, ", "))
            End If
            Return result
        End Function

        ''' <summary>
        ''' 指定プロパティ名をもとにSelect句を作成して返す
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <param name="specifiedNameByNormal">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MakeSelectClause(Of T)(selectionCallback As Func(Of SelectionField(Of T), T, Object),
                                                      specifiedNameByNormal As Dictionary(Of String, String)) As String
            Dim sb As New StringBuilder("SELECT ")
            Dim results As String() = ConvertPropertyNameToDbFieldName(selectionCallback, specifiedNameByNormal)
            sb.Append(If(CollectionUtil.IsEmpty(results), "*", Join(results, ","))).Append(" ")
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 指定プロパティ名をDB項目名にして返す
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <param name="specifiedNameByNormal">VOのプロパティ名を個別のDB項目名に対応させる場合に指定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertPropertyNameToDbFieldName(Of T)(selectionCallback As Func(Of SelectionField(Of T), T, Object),
                                                                      specifiedNameByNormal As Dictionary(Of String, String)) As String()
            If selectionCallback Is Nothing Then
                Return Nothing
            End If
            Return GetSpecifyPropertyNames(selectionCallback).Select(
                Function(name) If(specifiedNameByNormal Is Nothing OrElse Not specifiedNameByNormal.ContainsKey(name),
                                  StringUtil.Decamelize(name), specifiedNameByNormal(name))).ToArray()
        End Function

        ''' <summary>
        ''' 指定プロパティ名を取得する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="selectionCallback">取得項目を指定するcallback</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSpecifyPropertyNames(Of T)(selectionCallback As Func(Of SelectionField(Of T), T, Object)) As String()
            Return VoUtil.GetSpecifyPropertyNames(Of T)(Function(define, vo)
                                                            Dim selection As New WrappedSelectionField(Of T)(define, vo)
                                                            selectionCallback.Invoke(selection, vo)
                                                            Return define
                                                        End Function)
        End Function

    End Class
End Namespace