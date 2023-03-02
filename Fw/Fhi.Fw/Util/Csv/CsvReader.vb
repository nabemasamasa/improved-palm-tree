Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic.FileIO

Namespace Util.Csv
    Public Class CsvReader

        Public Class DataTableEditor
            ''' <summary>
            ''' データテーブルの列を生成します.
            ''' </summary>
            ''' <param name="dt">データテーブル</param>
            ''' <param name="columnName">列名称</param>
            ''' <param name="type">データ型</param>
            ''' <param name="defaultValue">デフォルト値[省略可]</param>
            ''' <remarks></remarks>
            Public Shared Sub ColumnsAdd(ByVal dt As DataTable, _
                                         ByVal columnName As String, _
                                         ByVal type As System.Type, _
                                         Optional ByVal defaultValue As Object = Nothing)
                dt.Columns.Add(columnName, type)

                If Not defaultValue Is Nothing Then
                    dt.Columns(columnName).DefaultValue = defaultValue
                End If
            End Sub

            ''' <summary>
            ''' データテーブルからCSV用2次元文字列を生成します.
            ''' </summary>
            ''' <param name="dt"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Shared Function GetMatrixStr(ByVal dt As DataTable) As String
                '削除対象データログ用文字列姿勢
                Dim bufData As New StringBuilder()
                Dim first As Boolean = True

                For Each newRow As DataRow In dt.Rows
                    For Each col As DataColumn In dt.Columns
                        If Not first Then bufData.Append(",")
                        bufData.Append(newRow(col.Caption).ToString())
                        first = False
                    Next

                    bufData.AppendLine("")
                Next

                Return bufData.ToString()
            End Function
        End Class

        ''' <summary>INIファイルコメント文字</summary>
        Public Const INI_COMMENT As Char = "#"c

        ''' <summary>CSVデータの内容</summary>
        Protected m_dtCsv As DataTable

        Public Sub New()
            m_dtCsv = New DataTable()
        End Sub

        Private _trimWhitespace As Boolean = True

        Public Property TrimWhitespace() As Boolean
            Get
                Return _trimWhitespace
            End Get
            Set(ByVal value As Boolean)
                _trimWhitespace = value
            End Set
        End Property

        ''' <summary>Csvデータテーブル</summary>
        Public Function GetTable() As DataTable
            Return m_dtCsv
        End Function

        ''' <summary>行コレクション</summary>
        Public ReadOnly Property Rows() As DataRowCollection
            Get
                Return m_dtCsv.Rows
            End Get
        End Property

        ''' <summary>列コレクション</summary>
        Public ReadOnly Property Columns() As DataColumnCollection
            Get
                Return m_dtCsv.Columns
            End Get
        End Property

        ''' <summary>
        ''' CSVファイルの内容を読み込み, データテーブルを生成します.
        ''' </summary>
        ''' <param name="file">読込みを行うCSVファイル(フルパス)</param>
        ''' <param name="isHeader">ヘッダー行があるか？</param>
        ''' <remarks></remarks>
        Public Sub Read(ByVal file As String, ByVal isHeader As Boolean)
            m_dtCsv.Clear()

            Dim str As String
            Using fs As New FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using reader As New StreamReader(fs, Encoding.GetEncoding("SHIFT_JIS"))
                    str = reader.ReadToEnd()
                End Using
            End Using

            Using parser = New TextFieldParser(New StringReader(CorrectIllegalDblQuote(str)))
                parser.CommentTokens = New String() {INI_COMMENT}
                parser.TextFieldType = FieldType.Delimited
                parser.SetDelimiters(",")
                parser.HasFieldsEnclosedInQuotes = True ' ダブルクォーテーション中のカンマは無視
                parser.TrimWhiteSpace = TrimWhitespace

                If parser.EndOfData() Then
                    Return
                End If

                Dim firstRowDatas As String() = parser.ReadFields()
                Me.CreateCsvColumn(firstRowDatas, isHeader)

                If Not isHeader Then
                    m_dtCsv.Rows.Add(firstRowDatas)
                End If

                While Not parser.EndOfData()
                    Dim rowDatas As String() = parser.ReadFields()
                    m_dtCsv.Rows.Add(rowDatas)
                End While
            End Using
        End Sub

        ''' <summary>
        ''' CSVのダブルクォートで囲まれた要素内で、ダブルクォートを二重にエスケープしていない文字列を補正する
        ''' </summary>
        ''' <param name="str">CSV内容</param>
        ''' <returns>補正したCSV内容</returns>
        ''' <remarks>
        ''' CSVファイルの仕様は、
        ''' 1)各値はダブルクォートで囲まれることがある 
        ''' 2)ダブルクォート内ではコンマや改行はエスケープされる 
        ''' 3)ダブルクォート内でダブルクォートをエスケープするにはダブルクォートを二重に書く 
        ''' であるが、3について、一重で書いた不正なCSVを二重に補正する
        ''' </remarks>
        Public Shared Function CorrectIllegalDblQuote(ByVal str As String) As String
            Const DQ As Char = """"c
            Dim chars As Char() = str.ToCharArray
            Dim dqStart As Boolean
            Dim sb As New StringBuilder
            For i As Integer = 0 To chars.Length - 1
                If dqStart Then
                    If chars(i) = DQ Then
                        If IsCharDqEndNext(chars, i + 1) Then
                            dqStart = False
                        ElseIf chars(i + 1) = DQ Then ' ダブルクォートの中にダブルクォートを含める場合は二重にする、という仕様のとおり
                            If IsCharDqEndNext(chars, i + 2) Then
                                ' ... ,"あ"", ...
                                sb.Append(DQ) ' 一重しかないので、1文字追加
                            Else
                                ' ... ,"あ""...
                                i += 1
                                sb.Append(DQ) ' 1文字先のダブルクォートもAppendする意
                            End If
                        Else ' ダブルクォートの中に、ダブルクォートが一重しかない、間違った仕様
                            sb.Append(DQ) ' 一重しかないので、1文字追加
                        End If
                    End If
                Else
                    If chars(i) = DQ Then
                        dqStart = True
                    End If
                End If
                sb.Append(chars(i))
            Next
            Return sb.ToString
        End Function

        Private Shared Function IsCharDqEndNext(ByVal chars As Char(), ByVal index As Integer) As Boolean
            Return chars.Length = index OrElse chars(index) = ","c OrElse chars(index) = ControlChars.Cr OrElse chars(index) = ControlChars.Lf
        End Function

        ''' <summary>
        ''' CSVデータテーブルの列を生成する
        ''' </summary>
        ''' <param name="rowDatas">行データ</param>
        ''' <param name="isHeader">ヘッダー行があるか？</param>
        ''' <remarks></remarks>
        Protected Sub CreateCsvColumn(ByVal rowDatas As String(), ByVal isHeader As Boolean)
            Dim tryCnt As Integer = 0 '列名称が被った場合の再試行回数

            If isHeader Then
                For Each rowData As String In rowDatas
                    Do
                        Dim colName As String

                        If tryCnt = 0 Then
                            '初回はそのまま
                            colName = rowData
                        Else
                            '2回目以降は'_[回数]'とする.
                            colName = rowData & "_" & tryCnt.ToString()
                        End If

                        If m_dtCsv.Columns.Contains(colName) Then
                            tryCnt += 1

                            '30回まで試す(plの員数とかそのくらいあるかも？)
                            If tryCnt > 5000 Then
                                Throw New InvalidProgramException
                            End If
                            Continue Do
                        End If
                        DataTableEditor.ColumnsAdd(m_dtCsv, colName, GetType(String), String.Empty)
                        Exit Do
                    Loop
                Next
            Else
                For i As Integer = 0 To rowDatas.Length - 1
                    DataTableEditor.ColumnsAdd(m_dtCsv, String.Format("F{0}", i + 1), GetType(String), String.Empty)
                Next
            End If
        End Sub
    End Class
End Namespace