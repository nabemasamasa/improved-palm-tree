Namespace Db.Sql

    ''' <summary>
    ''' if/where/setのXMLタグと、埋め込みパラメータ@Parameter を含むSQLを、解析し、適切なSQLを作成するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SqlAnalyzer
        Private ReadOnly sql As String
        Private ReadOnly bindParamValue As Object
        Private ReadOnly dbTypeOfString As DbType

        Private bindAnalyzer As SqlBindAnalyzer

        Private _analyzedSql As String
        ''' <summary>解析後のSQL</summary>
        Public Property AnalyzedSql() As String
            Get
                Return _analyzedSql
            End Get
            Private Set(ByVal value As String)
                _analyzedSql = value
            End Set
        End Property

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="sql">SQL</param>
        ''' <param name="bindParamValue">埋め込みパラメータを持つobject</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal sql As String, ByVal bindParamValue As Object)
            Me.New(sql, bindParamValue, DbType.String)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="sql">SQL</param>
        ''' <param name="bindParamValue">埋め込みパラメータを持つobject</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal sql As String, ByVal bindParamValue As Object, ByVal dbTypeOfString As DbType)
            Me.sql = sql
            Me.bindParamValue = bindParamValue
            Me.dbTypeOfString = dbTypeOfString
        End Sub

        ''' <summary>
        ''' SQLを解析する
        ''' </summary>
        ''' <remarks>解析の結果、SQL文の構成・埋め込みパラメータの抽出</remarks>
        Public Sub Analyze()
            Dim encloser As New XmlCDataEncloser("if", "where", "set", "join", "and", "or")
            Dim sqlWithTag As String = "<sql>" & XmlInequality.ConvInequality(encloser.Enclose(sql.Replace("&", "&amp;"))) & "</sql>"
            Dim tagAnalyzer As New SqlXmlAnalyzer(sqlWithTag, bindParamValue, New SqlExpressionEvaluator)
            tagAnalyzer.Analyze()
            Dim cleaner As New SqlWhitespaceCleaner
            AnalyzedSql = cleaner.Clean(tagAnalyzer.AnalyzedSql)

            bindAnalyzer = New SqlBindAnalyzer(AnalyzedSql, bindParamValue, dbTypeOfString)
            bindAnalyzer.Analyze()
        End Sub

        ''' <summary>
        ''' パラメータを SqlAccess に追加設定する
        ''' </summary>
        ''' <param name="db">パラメータを追加する SqlAccess</param>
        ''' <remarks></remarks>
        Public Sub AddParametersTo(ByVal db As DbAccess)
            bindAnalyzer.AddAllTo(db)
        End Sub

        ''' <summary>
        ''' バインドパラメータを直接書き込んだSQL文を作成する
        ''' </summary>
        ''' <returns>埋め込み後のSQL文</returns>
        ''' <remarks></remarks>
        Public Function MakeParametersBoundSql() As String
            Return bindAnalyzer.MakeParametersBoundSql
        End Function
    End Class
End Namespace