Namespace Db.Sys.Vo
    ''' <summary>
    ''' INFORMATION_SCHEMA.COLUMNSビュー
    ''' </summary>
    ''' <remarks></remarks>
    Public Class InformationSchemaColumnsVo

        ' テーブルカタログ名
        Private _TableCatalog As String
        ' テーブルスキーマ名
        Private _TableSchema As String
        ' テーブル名
        Private _TableName As String
        ' 列名
        Private _ColumnName As String
        ' 並び順
        Private _OrdinalPosition As Int32?
        ' デフォルト値
        Private _ColumnDefault As String
        ' Null許可
        Private _IsNullable As String
        ' データ型
        Private _DataType As String
        ' 最大文字数
        Private _CharacterMaximumLength As Int32?
        ' 文字の長さ
        Private _CharacterOctetLength As Int32?
        ' 
        Private _NumericPrecision As Int16?
        ' 
        Private _NumericPrecisionRadix As Int16?
        ' 
        Private _NumericScale As Int32?
        ' 
        Private _DatetimePrecision As Int16?
        ' 
        Private _CharacterSetCatalog As String
        ' 
        Private _CharacterSetSchema As String
        ' 文字コード名
        Private _CharacterSetName As String
        ' 
        Private _CollationCatalog As String
        ' 
        Private _CollationSchema As String
        ' 
        Private _CollationName As String
        ' 
        Private _DomainCatalog As String
        ' 
        Private _DomainSchema As String
        ' 
        Private _DomainName As String

        ''' <summary>テーブルカタログ名</summary>
        Public Property TableCatalog() As String
            Get
                Return _TableCatalog
            End Get
            Set(ByVal value As String)
                _TableCatalog = value
            End Set
        End Property

        ''' <summary>テーブルスキーマ名</summary>
        Public Property TableSchema() As String
            Get
                Return _TableSchema
            End Get
            Set(ByVal value As String)
                _TableSchema = value
            End Set
        End Property

        ''' <summary>テーブル名</summary>
        Public Property TableName() As String
            Get
                Return _TableName
            End Get
            Set(ByVal value As String)
                _TableName = value
            End Set
        End Property

        ''' <summary>列名</summary>
        Public Property ColumnName() As String
            Get
                Return _ColumnName
            End Get
            Set(ByVal value As String)
                _ColumnName = value
            End Set
        End Property

        ''' <summary>並び順</summary>
        Public Property OrdinalPosition() As Int32?
            Get
                Return _OrdinalPosition
            End Get
            Set(ByVal value As Int32?)
                _OrdinalPosition = value
            End Set
        End Property

        ''' <summary>デフォルト値</summary>
        Public Property ColumnDefault() As String
            Get
                Return _ColumnDefault
            End Get
            Set(ByVal value As String)
                _ColumnDefault = value
            End Set
        End Property

        ''' <summary>Null許可</summary>
        Public Property IsNullable() As String
            Get
                Return _IsNullable
            End Get
            Set(ByVal value As String)
                _IsNullable = value
            End Set
        End Property

        ''' <summary>データ型</summary>
        Public Property DataType() As String
            Get
                Return _DataType
            End Get
            Set(ByVal value As String)
                _DataType = value
            End Set
        End Property

        ''' <summary>最大文字数</summary>
        Public Property CharacterMaximumLength() As Int32?
            Get
                Return _CharacterMaximumLength
            End Get
            Set(ByVal value As Int32?)
                _CharacterMaximumLength = value
            End Set
        End Property

        ''' <summary>文字の長さ</summary>
        Public Property CharacterOctetLength() As Int32?
            Get
                Return _CharacterOctetLength
            End Get
            Set(ByVal value As Int32?)
                _CharacterOctetLength = value
            End Set
        End Property

        ''' <summary>Private _NumericPrecision As Int16?</summary>
        Public Property NumericPrecision() As Int16?
            Get
                Return _NumericPrecision
            End Get
            Set(ByVal value As Int16?)
                _NumericPrecision = value
            End Set
        End Property

        ''' <summary>Private _NumericPrecisionRadix As Int16?</summary>
        Public Property NumericPrecisionRadix() As Int16?
            Get
                Return _NumericPrecisionRadix
            End Get
            Set(ByVal value As Int16?)
                _NumericPrecisionRadix = value
            End Set
        End Property

        ''' <summary>Private _NumericScale As Int32?</summary>
        Public Property NumericScale() As Int32?
            Get
                Return _NumericScale
            End Get
            Set(ByVal value As Int32?)
                _NumericScale = value
            End Set
        End Property

        ''' <summary>Private _DatetimePrecision As Int16?</summary>
        Public Property DatetimePrecision() As Int16?
            Get
                Return _DatetimePrecision
            End Get
            Set(ByVal value As Int16?)
                _DatetimePrecision = value
            End Set
        End Property

        ''' <summary>Private _CharacterSetCatalog As String</summary>
        Public Property CharacterSetCatalog() As String
            Get
                Return _CharacterSetCatalog
            End Get
            Set(ByVal value As String)
                _CharacterSetCatalog = value
            End Set
        End Property

        ''' <summary>Private _CharacterSetSchema As String</summary>
        Public Property CharacterSetSchema() As String
            Get
                Return _CharacterSetSchema
            End Get
            Set(ByVal value As String)
                _CharacterSetSchema = value
            End Set
        End Property

        ''' <summary>文字コード名</summary>
        Public Property CharacterSetName() As String
            Get
                Return _CharacterSetName
            End Get
            Set(ByVal value As String)
                _CharacterSetName = value
            End Set
        End Property

        ''' <summary>Private _CollationCatalog As String</summary>
        Public Property CollationCatalog() As String
            Get
                Return _CollationCatalog
            End Get
            Set(ByVal value As String)
                _CollationCatalog = value
            End Set
        End Property

        ''' <summary>Private _CollationSchema As String</summary>
        Public Property CollationSchema() As String
            Get
                Return _CollationSchema
            End Get
            Set(ByVal value As String)
                _CollationSchema = value
            End Set
        End Property

        ''' <summary>Private _CollationName As String</summary>
        Public Property CollationName() As String
            Get
                Return _CollationName
            End Get
            Set(ByVal value As String)
                _CollationName = value
            End Set
        End Property

        ''' <summary>Private _DomainCatalog As String</summary>
        Public Property DomainCatalog() As String
            Get
                Return _DomainCatalog
            End Get
            Set(ByVal value As String)
                _DomainCatalog = value
            End Set
        End Property

        ''' <summary>Private _DomainSchema As String</summary>
        Public Property DomainSchema() As String
            Get
                Return _DomainSchema
            End Get
            Set(ByVal value As String)
                _DomainSchema = value
            End Set
        End Property

        ''' <summary>Private _DomainName As String</summary>
        Public Property DomainName() As String
            Get
                Return _DomainName
            End Get
            Set(ByVal value As String)
                _DomainName = value
            End Set
        End Property
    End Class
End Namespace