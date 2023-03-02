Namespace Db.Sys.Vo.Helper
    Public Class InformationSchemaColumnsVoHelper

        ''' <summary>テーブルスキーマ</summary>
        Public Class TableSchema
            ''' <summary>初期値のテーブルスキーマ名</summary>
            Public Const [DEFAULT] As String = "dbo"
        End Class

        ''' <summary>データ型</summary>
        Public Class DataType
            Public Const [CHAR] As String = "char"
            Public Const NCHAR As String = "nchar"
            Public Const VARCHAR As String = "varchar"
            Public Const NVARCHAR As String = "nvarchar"
            Public Const [INT] As String = "int"
            Public Const [TINYINT] As String = "tinyint"
            Public Const [BIGINT] As String = "bigint"
            Public Const NUMERIC As String = "numeric"
            Public Const [DECIMAL] As String = "decimal"
            Public Const [FLOAT] As String = "float"
            Public Const [DATETIME] As String = "datetime"
            Public Const REAL As String = "real"
        End Class

        ''' <summary>Null許可</summary>
        Public Class IsNullableValue
            Public Const [YES] As String = "YES"
            Public Const [NO] As String = "NO"
        End Class

        ''' <summary>
        ''' Nullが許可されているか？を返す
        ''' </summary>
        ''' <param name="vo">列定義vo</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Shared Function IsNullable(ByVal vo As InformationSchemaColumnsVo) As Boolean
            Return InformationSchemaColumnsVoHelper.IsNullableValue.YES.Equals(vo.IsNullable)
        End Function

        ''' <summary>
        ''' 初期値はNullか？を返す
        ''' </summary>
        ''' <param name="vo">列定義vo</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Shared Function IsDevaultValueNull(ByVal vo As InformationSchemaColumnsVo) As Boolean
            Return vo.ColumnDefault Is Nothing
        End Function

        Private vo As InformationSchemaColumnsVo
        Public Sub New(ByVal vo As InformationSchemaColumnsVo)
            Me.vo = vo
        End Sub
    End Class
End Namespace