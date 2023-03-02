Namespace Db
    ''' <summary>
    ''' SQLの検索条件を格納するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CriteriaParameter
        ''' <summary>検索条件</summary>
        Public Enum SearchCondition
            ''' <summary>等しいか(=)</summary>
            Equal
            ''' <summary>いずれかに等しい(IN)</summary>
            Any
            ''' <summary>ワイルドカード(LIKE)</summary>
            [Like]
            ''' <summary>検索値より大きいか(value &lt; DB_VALUE)</summary>
            GreaterThan
            ''' <summary>検索値より小さいか(DB_VALUE &lt; value)</summary>
            LessThan
            ''' <summary>検索値以上か(value &lt;= DB_VALUE)</summary>
            GreaterEqual
            ''' <summary>検索値以下か(DB_VALUE &lt;= value)</summary>
            LessEqual
            ''' <summary>AND</summary>
            [And]
            ''' <summary>AND (</summary>
            AndBracket
            ''' <summary>OR</summary>
            [Or]
            ''' <summary>OR (</summary>
            OrBracket
            ''' <summary>NOT</summary>
            [Not]
            ''' <summary>NOT (</summary>
            NotBracket
            ''' <summary>(</summary>
            Bracket
            ''' <summary>)</summary>
            CloseBracket
        End Enum

        Private ReadOnly _propertyName As String
        Private ReadOnly _identifyName As String
        Private ReadOnly _condition As SearchCondition
        Private ReadOnly _value As Object

        Public Sub New(propertyName As String, identifyName As String, condition As SearchCondition, value As Object)
            _propertyName = propertyName
            _identifyName = identifyName
            _condition = condition
            _value = value
        End Sub

        ''' <summary>プロパティ名</summary>
        Public ReadOnly Property PropertyName() As String
            Get
                Return _propertyName
            End Get
        End Property

        ''' <summary>識別名</summary>
        Public ReadOnly Property IdentifyName() As String
            Get
                Return _identifyName
            End Get
        End Property

        ''' <summary>条件</summary>
        Public ReadOnly Property Condition() As SearchCondition
            Get
                Return _condition
            End Get
        End Property

        ''' <summary>値</summary>
        Public ReadOnly Property Value() As Object
            Get
                Return _value
            End Get
        End Property

    End Class
End Namespace