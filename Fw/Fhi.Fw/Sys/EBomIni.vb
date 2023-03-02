Imports Microsoft.VisualBasic.FileIO
Imports System.Collections.Specialized
Imports System.Environment
Imports System.IO
Imports System.Text

Namespace Sys

    ''' <summary>
    ''' EBOM用INIファイル(CSV形式)にアクセスする機能を提供します.
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history>
    '''     2009/06/17    T.Hirasawa    新規作成
    ''' </history>
    Public Class EBomIni

        ''' <summary>EBOM用環境変数</summary>
        Public Const ENV_EBOM As String = "SDISINI"

        ''' <summary>INIファイルコメント文字</summary>
        Public Const INI_COMMENT As Char = "#"c

#Region " メンバー変数 "
        ''' <summary>読み込みを行うINIファイル名</summary>
        Private m_fileName As String = String.Empty

        ''' <summary>INIファイルの内容</summary>
        Private m_item As New NameValueCollection

        ''' <summary>区切り文字</summary>
        Private m_delimiter As String = ","

        ''' <summary>設定ファイルがあるルートフォルダ</summary>
        Private m_ebomEnv As String = ""
#End Region

#Region " コンストラクタ "
        ''' <summary>
        ''' デフォルトコンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            m_fileName = String.Empty
        End Sub

        ''' <summary>
        ''' ファイル名を指定してオブジェクトを生成します.
        ''' </summary>
        ''' <param name="fileName">読み込みを行うINIファイル名</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal fileName As String)
            m_fileName = fileName
            Me.Read()
        End Sub

        ''' <summary>
        ''' ファイル名,区切り文字を指定してオブジェクトを生成します.
        ''' </summary>
        ''' <param name="fileName">読み込みを行うINIファイル名</param>
        ''' <param name="delimiter">区切り文字</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal fileName As String, ByVal delimiter As String)
            m_fileName = fileName
            m_delimiter = delimiter
            Me.Read()
        End Sub
#End Region

#Region " INIファイルの内容を取得 "
        ''' <summary>
        ''' INIファイルの内容を読み込みます.
        ''' 既に読み込まれている場合,最新の情報で上書きします.
        ''' </summary>
        ''' <remarks></remarks>
        ''' <history>
        '''     2009/06/17    T.Hirasawa    新規作成
        ''' </history>
        Public Sub Read()
            Dim filePath As String = String.Empty        'INIファイル(フルパス)

            m_item.Clear()

            If m_ebomEnv.Equals(String.Empty) Then
                'SDISINI値取得
                m_ebomEnv = GetEnvironmentVariable(ENV_EBOM)
            End If

            'INIファイルのパスを生成
            filePath = Path.Combine(m_ebomEnv, m_fileName)

            'INIファイル読み込み
            Using parser = New TextFieldParser(filePath, Encoding.GetEncoding("SHIFT_JIS"))
                parser.TextFieldType = FieldType.Delimited     '区切り形式
                parser.SetDelimiters(m_delimiter)              '区切り文字
                parser.HasFieldsEnclosedInQuotes = False       'ダブルクォーテーション中のカンマは無視
                parser.TrimWhiteSpace = True                   '空白を取り除く

                '行を読み込み
                While Not parser.EndOfData()
                    Dim rowData() As String = parser.ReadFields()       '行データ

                    '列数が1以上であり,コメント行ではない
                    If rowData.Length > 1 AndAlso _
                       Not rowData(0).Chars(0).Equals(INI_COMMENT) Then

                        Dim bufValue As String = String.Empty

                        '列数が2以上の値を繋げる
                        For i As Integer = 1 To rowData.Length - 1
                            '3以上の場合, デリミタを付加する.
                            If Not i = 1 Then
                                bufValue += m_delimiter
                            End If

                            bufValue += rowData(i)
                        Next

                        'キーと値を保持
                        m_item.Add(rowData(0), bufValue)
                    End If
                End While
            End Using

        End Sub
#End Region

#Region " プロパティー "
        ''' <summary>
        ''' 読み込みを行うINIファイル名
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' <history>
        '''     2009/06/17    T.Hirasawa    新規作成
        ''' </history>
        Public Property FileName() As String
            Get
                Return m_fileName
            End Get
            Set(ByVal value As String)
                m_fileName = value
            End Set
        End Property

        ''' <summary>
        ''' キー値に対応する値を取得します.
        ''' </summary>
        ''' <param name="key">キー値</param>
        ''' <value></value>
        ''' <returns>
        ''' キー値に対応する値.
        ''' 存在しない場合, 空文字(String.Empty)を返します.
        ''' </returns>
        ''' <remarks></remarks>
        ''' <history>
        '''     2009/06/17    T.Hirasawa    新規作成
        ''' </history>
        Default Public Property Item(ByVal key As String) As String
            Get
                Dim ret As String = m_item(key)
                Return IIf(ret Is Nothing, String.Empty, ret).ToString()
            End Get
            Set(ByVal value As String)
                m_item(key) = value
            End Set
        End Property

        ''' <summary>
        ''' 区切り文字
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Delimiter() As String
            Get
                Return m_delimiter
            End Get
            Set(ByVal value As String)
                m_delimiter = value
            End Set
        End Property

        ''' <summary>
        ''' INIファイルのあるフォルダ
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property EBomEnv() As String
            Get
                Return m_ebomEnv
            End Get
            Set(ByVal value As String)
                m_ebomEnv = value
            End Set
        End Property
#End Region

    End Class


End Namespace
