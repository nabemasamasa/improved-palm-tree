Imports System.Xml
Imports System.Text.RegularExpressions
Imports Fhi.Fw.Domain

Namespace Db.Sql
    ''' <summary>
    ''' XML化したSQLを解析するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SqlXmlAnalyzer

        Private ReadOnly xmlSql As String
        Private ReadOnly bindParamValue As Object
        Private ReadOnly evaluator As ISqlExpressionEvaluator
        Private _AnalyzedSql As String
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="xmlSql">XML化したSQL</param>
        ''' <param name="bindParamValue">埋め込みパラメータ値</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal xmlSql As String, ByVal bindParamValue As Object, ByVal evaluator As ISqlExpressionEvaluator)
            Me.xmlSql = xmlSql
            Me.bindParamValue = bindParamValue
            Me.evaluator = evaluator
        End Sub

        ''' <summary>解析したSQL結果</summary>
        Public ReadOnly Property AnalyzedSql() As String
            Get
                Return _AnalyzedSql
            End Get
        End Property

        ''' <summary>
        ''' 解析する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Analyze()
            Dim doc As New XmlDocument
            Try
                doc.LoadXml(xmlSql)
            Catch ex As XmlException
                Throw New InvalidOperationException("タグが閉じられていない等の原因によるエラー:" & ex.Message, ex)
            End Try
            Dim results As New List(Of String)
            PerformAnalyze(doc.ChildNodes(0), results)
            _AnalyzedSql = Join(results.ToArray)
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="node"></param>
        ''' <param name="results"></param>
        ''' <remarks></remarks>
        Private Sub PerformAnalyze(ByVal node As XmlNode, ByVal results As List(Of String))
            For Each childNode As XmlNode In node
                Select Case childNode.Name.ToLower
                    Case "where", "and", "or"
                        Dim whereResults As New List(Of String)
                        PerformAnalyze(childNode, whereResults)
                        If 0 < whereResults.Count Then
                            If "where".Equals(childNode.Name.ToLower) Then
                                results.Add("WHERE " & JoinWhereParts(whereResults))
                            Else
                                results.Add(String.Format("{0} ({1})", childNode.Name.ToUpper, JoinWhereParts(whereResults)))
                            End If
                        End If
                    Case "set"
                        Dim setResults As New List(Of String)
                        PerformAnalyze(childNode, setResults)
                        If 0 < setResults.Count Then
                            results.Add("SET " & JoinSetParts(setResults))
                        End If
                    Case "if"
                        If EvaluateIf(childNode) Then
                            PerformAnalyze(childNode, results)
                        End If
                    Case "join"
                        results.Add(PerformJoin(childNode))
                    Case "#cdata-section"
                        results.Add(childNode.Value.Trim)
                    Case "#text"
                        results.Add(SqlBindUtil.ConvInternalBindNameIfNecessary(childNode.Value, bindNameRegex))
                    Case Else
                        Throw New NotSupportedException(childNode.Name & " は未対応です.")
                End Select
            Next
        End Sub

        Private bindNameRegex As Regex = SqlBindUtil.NewUserBindNameRegex()

        ''' <summary>
        ''' ifタグの条件を評価して返す
        ''' </summary>
        ''' <param name="node">ifのXmlNode</param>
        ''' <returns>評価した結果</returns>
        ''' <remarks></remarks>
        Private Function EvaluateIf(ByVal node As XmlNode) As Boolean
            For Each attrNode As XmlNode In node.Attributes
                If attrNode.Name.ToLower = "test" Then
                    Return evaluator.Evaluate(attrNode.Value, bindParamValue)
                Else
                    Throw New NotSupportedException("<if>の属性 " & attrNode.Name & " は未対応です.")
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Set句のカンマ区切りを考慮して結合して返す
        ''' </summary>
        ''' <param name="setParts">パート毎に分かれたSet句</param>
        ''' <returns>結合したSet句</returns>
        ''' <remarks></remarks>
        Private Function JoinSetParts(ByVal setParts As List(Of String)) As String
            Dim count As Integer = setParts.Count
            If 0 < count Then
                If 0 < setParts(0).Length AndAlso setParts(0).StartsWith(",") Then
                    setParts(0) = setParts(0).Substring(1).Trim
                End If
                If 0 < setParts(count - 1).Length AndAlso setParts(count - 1).EndsWith(",") Then
                    setParts(count - 1) = setParts(count - 1).Substring(0, setParts(count - 1).Length - 1).Trim
                End If
            End If
            Return Join(setParts.ToArray)
        End Function

        ''' <summary>
        ''' 論理演算子から始まっているかを返す
        ''' </summary>
        ''' <param name="str">判定文字列</param>
        ''' <param name="logicalStr">論理演算子</param>
        ''' <returns>始まっている場合、true</returns>
        ''' <remarks></remarks>
        Private Function StartsWithLogical(ByVal str As String, ByVal logicalStr As String) As Boolean
            Return logicalStr.Length < str.Length AndAlso str.ToLower.StartsWith(logicalStr.ToLower) AndAlso IsLogicalNextChar(str.Chars(logicalStr.Length))
        End Function

        ''' <summary>
        ''' 論理演算子の次に来るべき文字かを返す
        ''' </summary>
        ''' <param name="c">判定文字</param>
        ''' <returns>次の文字なら、true</returns>
        ''' <remarks></remarks>
        Private Function IsLogicalNextChar(ByVal c As Char) As Boolean
            Return 0 <= (" (" & vbCrLf & vbTab).IndexOf(c)
        End Function

        ''' <summary>
        ''' 論理演算子で終わっているかを返す
        ''' </summary>
        ''' <param name="str">判定文字列</param>
        ''' <param name="logicalStr">論理演算子</param>
        ''' <returns>終わっている場合、true</returns>
        ''' <remarks></remarks>
        Private Function EndsWithLogical(ByVal str As String, ByVal logicalStr As String) As Boolean
            Return logicalStr.Length < str.Length AndAlso str.ToLower.EndsWith(logicalStr.ToLower) AndAlso IsLoicalPrevChar(str.Chars(str.Length - logicalStr.Length - 1))
        End Function

        ''' <summary>
        ''' 論理演算子の前に来るべき文字かを返す
        ''' </summary>
        ''' <param name="c">判定文字</param>
        ''' <returns>前の文字なら、true</returns>
        ''' <remarks></remarks>
        Private Function IsLoicalPrevChar(ByVal c As Char) As Boolean
            Const LOGICAL_PREV_CHARS As String = " )" & vbCrLf & vbTab
            Return 0 <= LOGICAL_PREV_CHARS.IndexOf(c)
        End Function

        ''' <summary>
        ''' Where句のAnd,Or区切りを考慮して、結合して返す
        ''' </summary>
        ''' <param name="whereParts">パート毎に分かれたWhere句の</param>
        ''' <returns>結合したWhere句</returns>
        ''' <remarks></remarks>
        Private Function JoinWhereParts(ByVal whereParts As List(Of String)) As String
            Dim count As Integer = whereParts.Count
            If 0 < count Then
                Dim wherePartStart As String = whereParts(0).TrimStart
                If StartsWithLogical(wherePartStart, "AND") Then
                    whereParts(0) = wherePartStart.Substring(3).Trim
                ElseIf StartsWithLogical(wherePartStart, "OR") Then
                    whereParts(0) = wherePartStart.Substring(2).Trim
                End If
                Dim wherePartEnd As String = whereParts(count - 1).TrimEnd
                If EndsWithLogical(wherePartEnd, "AND") Then
                    whereParts(count - 1) = wherePartEnd.Substring(0, wherePartEnd.Length - 3).Trim
                ElseIf EndsWithLogical(wherePartEnd, "OR") Then
                    whereParts(count - 1) = wherePartEnd.Substring(0, wherePartEnd.Length - 2).Trim
                End If
            End If
            Return Join(whereParts.ToArray)
        End Function

        Private Function PerformJoin(ByVal node As XmlNode) As String
            Dim bindName As String = Nothing
            Dim separator As String = Nothing
            For Each attrNode As XmlNode In node.Attributes
                If "property".Equals(attrNode.Name.ToLower) Then
                    bindName = attrNode.Value
                ElseIf "separator".Equals(attrNode.Name.ToLower) Then
                    separator = attrNode.Value
                Else
                    Throw New NotSupportedException("<join>の属性 " & attrNode.Name & " は未対応です.")
                End If
            Next
            If bindName Is Nothing Then
                Throw New InvalidOperationException("<join>の属性 property が必要.")
            End If
            If separator Is Nothing Then
                separator = String.Empty
            End If

            Dim propertyName As String = SqlBindUtil.ConvBindNameToPropertyName(bindName)
            Dim bindNameUserFormat As String = bindName & "()"

            If 0 < node.ChildNodes.Count AndAlso node.InnerXml.IndexOf(bindNameUserFormat) < 0 Then
                Throw New InvalidOperationException("<join> : '" & bindName & "' の一要素を表す '" & bindNameUserFormat & "' を使用すべき. " & node.InnerXml)
            End If

            Dim value As Object
            If VoUtil.PROPERTY_NAME_SELF_VALUE.Equals(propertyName) Then
                value = bindParamValue
            ElseIf TypeOf bindParamValue Is CriteriaBinder Then
                value = DirectCast(bindParamValue, CriteriaBinder).GetValueByIdentifyName(propertyName)
            Else
                value = VoUtil.GetValue(bindParamValue, propertyName)
            End If
            If value Is Nothing Then
                Return Nothing
            End If
            Dim values As New List(Of Object)
            If TypeOf value Is ICollection Then
                For Each obj As Object In DirectCast(value, ICollection)
                    values.Add(obj)
                Next
            ElseIf TypeOf value Is ICollectionObject Then
                Dim collectionObject As ICollectionObject = DirectCast(value, ICollectionObject)
                For i = 0 To collectionObject.Count - 1
                    values.Add(collectionObject(i))
                Next
            Else
                Throw New InvalidOperationException("<join>は、配列またはcollectionに使用すべきだが、'" & bindName & "' は、" & value.GetType.Name & "型")
            End If

            Dim result As New List(Of String)
            If node.ChildNodes.Count = 0 Then
                For index As Integer = 0 To values.Count - 1
                    result.Add(SqlBindUtil.ConvPropertyNameToInternalBindName(propertyName, index))
                Next
            Else
                For index As Integer = 0 To values.Count - 1
                    Dim doc As New XmlDocument
                    doc.LoadXml(node.OuterXml.Replace(bindNameUserFormat, bindName & String.Format("({0})", index)))
                    Dim childNodeResults As New List(Of String)
                    PerformAnalyze(doc.ChildNodes(0), childNodeResults)
                    result.Add(Join(childNodeResults.ToArray, ""))
                Next
            End If
            Return Join(result.ToArray, separator)
        End Function
    End Class
End Namespace