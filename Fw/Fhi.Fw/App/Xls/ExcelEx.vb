'遅延バインディングを使用しているため OFF
Option Strict Off

Imports System.Runtime.InteropServices.Marshal
Imports System.Runtime.InteropServices

Namespace App.Xls

    ''' <summary>
    ''' Excel操作を担うクラス
    ''' </summary>
    ''' <remarks>
    ''' 【注意】
    ''' 1) 特定のバージョンのExcelに依存しないよう、遅延バインディングを採用している.
    '''   その為、当クラス使用時にExcelオブジェクトへの参照設定は不要だが、
    '''   メソッド候補表示なし・定数候補表示なしであり、パフォーマンスも低下している.
    ''' 2) メソッド内で生成したCOMオブジェクトはメソッド内でメモリ解放(Free)しないと、
    '''   Excelプロセスが残り続けるので要注意である.
    ''' 3) ただし、m_xlBook.Sheets(i) で取得する Worksheet オブジェクトは解放してはいけない.
    '''   一度解放すると、m_xlBook.Sheets(i) で取得できなくなる
    '''   （解放は、Dispose(Boolean) で行っているので解放漏れにはならない）
    ''' </remarks>
    Public Class ExcelEx
        Implements IDisposable

#Region "Private Enums..."
        ''' <summary>バーコードのスタイル</summary>
        Private Enum XlBarCodeStyle
            ''' <summary>商品のマーキングに使うPOSシンボル（米国とカナダ）</summary>
            UPC_A = 0
            ''' <summary>UPCの短縮版</summary>
            UPC_E = 1
            ''' <summary>国際的な規格のPOSシンボル</summary>
            JAN_13 = 2
            ''' <summary>JANの短縮版</summary>
            JAN_8 = 3
            ''' <summary>小売り店向けのシンボル</summary>
            CASECODE = 4
            ''' <summary>英数字を表す。</summary>
            NW_7 = 5
            ''' <summary>英数字を表す。工業用</summary>
            CODE_39 = 6
            ''' <summary>ASCII 128文字を表す</summary>
            CODE_128 = 7
            ''' <summary>郵便物（米国）</summary>
            US_POSTNET = 8
            ''' <summary>郵便物（特殊）（米国）</summary>
            US_POSTAL_FIM = 9
            ''' <summary>郵便物（日本）</summary>
            CUSTOMER_BARCODE = 10
        End Enum

        ''' <summary>バーコードの表示方向</summary>
        Private Enum XlBarCodeDirection
            ''' <summary>標準方向</summary>
            NORMAL = 0
            ''' <summary>標準方向より時計回りに90°回転</summary>
            ROTATION_90 = 3
            ''' <summary>標準方向より時計回りに180°回転</summary>
            ROTATION_180 = 2
            ''' <summary>標準方向より時計回りに270°回転</summary>
            ROTATION_270 = 1
        End Enum

        Private Enum XlFixedFormatType
            ''' <summary>"PDF" — Portable Document Format file (.pdf). </summary>
            xlTypePDF
            ''' <summary>"XPS" — XPS Document (.xps). </summary>
            xlTypeXPS
        End Enum

        Private Enum XlFixedFormatQuality
            xlQualityStandard
            xlQualityMinimum
        End Enum

        Private Enum XlWindowView
            xlNormalView = 1
            xlPageBreakPreview = 2
        End Enum

        Private Enum XlDirection
            xlDown = -4121
            xlToLeft = -4159
            xlToRight = -4161
            xlUp = -4162
        End Enum

#End Region
#Region "Nested classes..."
        Public Interface IBehavior
            Function CreateObject() As Object
        End Interface

        Private Class DefaultBehaviorImpl : Implements IBehavior
            Public Function CreateObject2() As Object Implements IBehavior.CreateObject
                Return CreateObject("Excel.Application")
            End Function
        End Class

        ''' <summary>セルの書式 - 配置</summary>
        Public Class Alignment
            ''' <summary>文字の配置 - 水平位置</summary>
            Public HorizontalAlignment As XlHAlign
            ''' <summary>文字の配置 - 垂直位置</summary>
            Public VerticalAlignment As XlVAlign
            ''' <summary>文字の配置 - インデント</summary>
            Public IndentLevel As Integer?
            ''' <summary>文字の制御 - 折り返して全体を表示する場合、true</summary>
            Public WrapText As Boolean?
            ''' <summary>文字の制御 - 縮小して全体を表示する場合、true</summary>
            Public ShrinkToFit As Boolean?
            ''' <summary>文字の制御 - セルを結合する場合、true</summary>
            Public MergeCells As Boolean?
            ''' <summary>文字列の方向</summary>
            Public Orientation As XlOrientation
            ''' <summary>均等割り付け時に、文字列を自動的にインデントする場合、true</summary>
            Public AddIndent As Boolean?
        End Class
#End Region

        Public Shared DefaultBehavior As IBehavior = New DefaultBehaviorImpl

        ''' <summary>エクセルアプリケーション</summary>
        Protected m_xlApp As Object

        ''' <summary>エクセルブック</summary>
        Protected m_xlBook As Object

        ''' <summary>ブックを開いているかを表すフラグ</summary>
        Protected m_isBookOpen As Boolean = False

        ''' <summary>操作対象のシート</summary>
        Protected m_activeSheet As Object

        ''' <summary>開いているエクセルファイル(フルパス)</summary>
        Protected m_fileName As String = String.Empty

        ' アクティブなセル範囲定義
        Private definition As CellRangeDefinition

        ''' <summary>開いているエクセルファイル(フルパス)</summary>
        Public ReadOnly Property FileName() As String
            Get
                Return m_fileName
            End Get
        End Property

        ''' <summary>
        ''' デフォルトコンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me.New(Nothing)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="fileName">ファイル名(フルパス)</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal fileName As String)
            'クラスの初期化
            Me.Initialize()

            If String.IsNullOrEmpty(fileName) OrElse (Not System.IO.File.Exists(fileName)) Then
                PerformCreateBook(fileName)
                Return
            End If

            PerformOpenBook(fileName)
        End Sub

        ''' <summary>
        ''' クラスの初期化
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Initialize()
            'アプリケーションオブジェクト生成
            '※この時点でプロセスが作成されます
            m_xlApp = DefaultBehavior.CreateObject

            '非表示
            m_xlApp.Visible = False

            '警告メッセージを表示しない
            m_xlApp.DisplayAlerts = False
        End Sub

        ''' <summary>
        ''' ブックを作成する
        ''' </summary>
        ''' <param name="fileName">作成するファイル名</param>
        ''' <remarks></remarks>
        Private Sub PerformCreateBook(ByVal fileName As String)
            Dim xlBooks As Object = Nothing
            Dim defaultSaveFormat As XlFileFormat = m_xlApp.DefaultSaveFormat

            Try
                'ファイル名(フルパス)を保持
                m_fileName = fileName

                If StringUtil.IsNotEmpty(fileName) Then
                    m_xlApp.DefaultSaveFormat = GetFileFormat(fileName)
                End If

                xlBooks = m_xlApp.Workbooks
                m_xlBook = xlBooks.Add()
                m_isBookOpen = True

                'アクティブシートを設定
                Me.SetActiveSheet(1)

                '再計算OFF.ブックを開いてから指定しないとエラー.
                m_xlApp.Calculation = XlCalculation.xlCalculationManual

            Finally
                'ユーザー設定を上書きしてしまう為、ブック作成が終わったら元に戻す
                m_xlApp.DefaultSaveFormat = defaultSaveFormat
                Me.Free(xlBooks)
            End Try
        End Sub

        ''' <summary>
        ''' ブックを開く
        ''' </summary>
        ''' <param name="fileName">開くファイル名</param>
        ''' <remarks></remarks>
        Private Sub PerformOpenBook(ByVal fileName As String)
            Dim xlBooks As Object = Nothing

            Try
                'ファイル名(フルパス)を保持
                m_fileName = fileName

                xlBooks = m_xlApp.Workbooks
                m_xlBook = xlBooks.Open(m_fileName, True)
                m_isBookOpen = True

                'アクティブシートを設定
                Me.SetActiveSheet(1)

                '再計算OFF.ブックを開いてから指定しないとエラー.
                m_xlApp.Calculation = XlCalculation.xlCalculationManual

            Finally
                Me.Free(xlBooks)
            End Try
        End Sub

        ''' <summary>
        ''' 上書き保存します.
        ''' </summary>
        ''' <param name="isReadOnly">読み取り専用で保存するか? (初期値:False)</param>
        ''' <returns>成功した場合 True, 失敗した場合 False を返します.</returns>
        ''' <remarks></remarks>
        Public Function Save(Optional isReadOnly As Boolean = False) As Boolean
            Return SaveAs(m_fileName, isReadOnly)
        End Function

        ''' <summary>
        ''' 名前を付けてブックを保存します.
        ''' </summary>
        ''' <param name="fileName">保存するファイル(フルパス)</param>
        ''' <param name="isReadOnly">読み取り専用で保存するか? (初期値:False)</param>
        ''' <returns>成功した場合 True, 失敗した場合 False を返します.</returns>
        ''' <remarks></remarks>
        Public Function SaveAs(ByVal fileName As String, Optional isReadOnly As Boolean = False) As Boolean
            '再計算ON.ブックを保存前に指定しないと戻らない?
            m_xlApp.Calculation = XlCalculation.xlCalculationAutomatic
            m_xlBook.SaveAs(fileName, GetFileFormat(fileName), ReadOnlyRecommended:=isReadOnly)

            Return True
        End Function

        ''' <summary>
        ''' ファイル名にあったファイル形式定数を返す
        ''' </summary>
        ''' <param name="fileName">ファイル名</param>
        ''' <returns>ファイル形式定数</returns>
        ''' <remarks></remarks>
        Private Function GetFileFormat(ByVal fileName As String) As XlFileFormat

            Dim ext As String = StringUtil.ToUpper(System.IO.Path.GetExtension(fileName))
            If ".XLS".Equals(ext) Then
                Return XlFileFormat.xlWorkbookNormal
            ElseIf ".XLSX".Equals(ext) Then
                Return XlFileFormat.xlWorkbookDefault
            Else
                Return XlFileFormat.xlWorkbookNormal
            End If
        End Function

        ''' <summary>
        ''' ブックを閉じます.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub CloseBook()
            If m_isBookOpen Then
                m_xlApp.Calculation = XlCalculation.xlCalculationAutomatic
                m_xlBook.Close()
            End If
        End Sub

        ''' <summary>
        ''' シートインデックスをもとに, アクティブシートの変更を行います.
        ''' </summary>
        ''' <param name="index">シートインデックス(1から始まります)</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetActiveSheet(ByVal index As Integer)
            Dim xlSheets As Object = Nothing

            Try
                xlSheets = m_xlBook.Sheets

                'シートの選択
                m_activeSheet = xlSheets.Item(index)
                m_xlBook.Activate()
                m_activeSheet.Select()

            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' シート名をもとにアクティブシートの変更を行います.
        ''' </summary>
        ''' <param name="sheetName">シート名</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetActiveSheet(ByVal sheetName As String)
            Dim xlSheets As Object = Nothing

            Try
                xlSheets = m_xlBook.Sheets

                'シートの選択
                m_activeSheet = xlSheets.Item(sheetName)
                m_xlBook.Activate()
                m_activeSheet.Select()

            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' シートの追加を行います.
        ''' </summary>
        ''' <param name="sheetName">シート名[省略可]</param>
        ''' <remarks></remarks>
        Public Sub AddSheet(Optional ByVal sheetName As String = "")
            Dim xlSheets As Object = Nothing
            Dim xlLastSheet As Object = Nothing

            Try
                xlSheets = m_xlBook.Sheets
                xlLastSheet = xlSheets.Item(xlSheets.Count)

                'シートの追加
                If xlLastSheet Is Nothing Then
                    m_activeSheet = xlSheets.Add()
                Else
                    m_activeSheet = xlSheets.Add(After:=xlLastSheet)    '末尾に追加
                End If

                'シート名設定
                If Not sheetName.Equals(String.Empty) Then m_activeSheet.Name = sheetName

            Finally
                'Dispose 時に開放する(ここで開放すると開放したシートにアクセスできない)
                'Me.Free(xlLastSheet)

                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' アクティブシートのシート名設定
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <remarks></remarks>
        Public Sub SetSheetName(ByVal sheetName As String)
            m_activeSheet.Name = sheetName
        End Sub

        ''' <summary>
        ''' 指定されたシートのシート名を設定します.
        ''' </summary>
        ''' <param name="sheet">変更するシートのインデックス及び名称</param>
        ''' <param name="sheetName">シート名称</param>
        ''' <remarks></remarks>
        Public Sub SetSheetName(ByVal sheet As Object, ByVal sheetName As String)
            Dim xlSheets As Object = Nothing
            Dim xlSheet As Object = Nothing

            Try
                xlSheets = m_xlBook.Sheets
                xlSheet = xlSheets.Item(sheet)
                xlSheet.Name = sheetName

            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' シートの表示・非表示設定
        ''' </summary>
        ''' <param name="index">対象のシート名</param>
        ''' <param name="visible">表示・非表示フラグ</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetVisibleToSheet(ByVal index As Integer, ByVal visible As Boolean)
            Dim xlSheets As Object = Nothing
            Dim xlSheet As Object = Nothing

            Try
                xlSheets = m_xlBook.Sheets
                xlSheet = xlSheets.Item(index)
                xlSheet.Visible = visible

            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' シートの表示・非表示設定
        ''' </summary>
        ''' <param name="name">対象のシート名</param>
        ''' <param name="visible">表示・非表示フラグ</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetVisibleToSheet(ByVal name As String, ByVal visible As Boolean)
            Dim xlSheets As Object = Nothing
            Dim xlSheet As Object = Nothing

            Try
                xlSheets = m_xlBook.Sheets
                xlSheet = xlSheets.Item(name)
                xlSheet.Visible = visible

            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' 列の挿入を行います.
        ''' </summary>
        ''' <param name="col">挿入する列位置</param>
        ''' <param name="count">挿入する列数</param>
        ''' <remarks></remarks>
        Public Sub InsertColumn(ByVal col As Integer, Optional ByVal count As Integer = 1)
            Dim xlColumn As Object = Nothing

            Try
                xlColumn = Me.GetColumn(col, col + count - 1)
                xlColumn.Insert()

            Finally
                Me.Free(xlColumn)
            End Try
        End Sub

        ''' <summary>
        ''' 複数列コピーして列挿入する
        ''' </summary>
        ''' <param name="copyColumn">コピー列位置</param>
        ''' <param name="copyColumnCount">コピー列数</param>
        ''' <param name="insertColumn">挿入列位置</param>
        ''' <param name="repeatCount">繰り返し数</param>
        ''' <remarks></remarks>
        Public Sub CopyColumnsAndInsertRepeat(copyColumn As Integer, copyColumnCount As Integer, insertColumn As Integer, _
                                              Optional repeatCount As Integer = 1)
            ' 空列を挿入
            Me.InsertColumn(insertColumn, copyColumnCount * repeatCount)
            Dim srcCopyColumn As Integer = If(insertColumn <= copyColumn,
                                           copyColumn + copyColumnCount * repeatCount, copyColumn)
            Dim xlColumn As Object = Me.GetColumn(srcCopyColumn, srcCopyColumn + copyColumnCount - 1)
            Try
                Dim xlDestColumn As Object = Me.GetColumn(insertColumn, insertColumn + copyColumnCount * repeatCount - 1)
                Try
                    'クリップボードを経由しないコピー貼り付け
                    xlColumn.Copy(xlDestColumn)
                Finally
                    Me.Free(xlDestColumn)
                End Try
            Finally
                Me.Free(xlColumn)
            End Try
        End Sub

        ''' <summary>
        ''' 1列コピーして、それを繰り返し列挿入する
        ''' </summary>
        ''' <param name="copyColumn">コピー元行</param>
        ''' <param name="insertColumn">挿入開始位置</param>
        ''' <param name="repeatCount">繰り返し数</param>
        ''' <remarks></remarks>
        Public Sub CopyColumnAndInsertRepeat(ByVal copyColumn As Integer, ByVal insertColumn As Integer, ByVal repeatCount As Integer)
            CopyColumnsAndInsertRepeat(copyColumn, 1, insertColumn, repeatCount)
        End Sub

        ''' <summary>
        ''' 列の削除を行います.
        ''' </summary>
        ''' <param name="col">挿入する列位置</param>
        ''' <param name="count">挿入する列数</param>
        ''' <remarks></remarks>
        Public Sub DeleteCol(ByVal col As Integer, Optional ByVal count As Integer = 1)
            Dim xlColumn As Object = Nothing

            Try
                xlColumn = Me.GetColumn(col, col + count - 1)
                xlColumn.Delete()

            Finally
                Me.Free(xlColumn)
            End Try
        End Sub

        ''' <summary>
        ''' 行の挿入を行います.
        ''' </summary>
        ''' <param name="row">挿入する行位置</param>
        ''' <param name="count">挿入する行数</param>
        ''' <remarks></remarks>
        Public Sub InsertRow(ByVal row As Integer, Optional ByVal count As Integer = 1)
            Dim xlRow As Object = Nothing

            Try
                xlRow = Me.GetRow(row, row + count - 1)
                xlRow.Insert()

            Finally
                Me.Free(xlRow)
            End Try
        End Sub

        ''' <summary>
        ''' 複数行コピーして行挿入する
        ''' </summary>
        ''' <param name="copyRow">コピー行位置</param>
        ''' <param name="copyRowCount">コピー行数</param>
        ''' <param name="insertRow">挿入行位置</param>
        ''' <param name="repeatCount">繰り返し数</param>
        ''' <remarks></remarks>
        Public Sub CopyRowsAndInsertRepeat(copyRow As Integer, copyRowCount As Integer, insertRow As Integer, _
                                           Optional repeatCount As Integer = 1)
            ' 空行を挿入
            Me.InsertRow(insertRow, copyRowCount * repeatCount)
            Dim srcCopyRow As Integer = If(insertRow <= copyRow,
                                           copyRow + copyRowCount * repeatCount, copyRow)
            Dim xlBaseRow As Object = GetRow(srcCopyRow, srcCopyRow + copyRowCount - 1)
            Try
                Dim xlDestRow As Object = GetRow(insertRow, insertRow + copyRowCount * repeatCount - 1)
                Try
                    'クリップボードを経由しないコピー貼り付け
                    xlBaseRow.Copy(xlDestRow)
                Finally
                    Me.Free(xlDestRow)
                End Try
            Finally
                Me.Free(xlBaseRow)
            End Try
        End Sub

        ''' <summary>
        ''' 1行コピーして、それを繰り返し行挿入する
        ''' </summary>
        ''' <param name="copyRow">コピー元行</param>
        ''' <param name="insertRow">挿入開始位置</param>
        ''' <param name="repeatCount">繰り返し数</param>
        ''' <remarks></remarks>
        Public Sub CopyRowAndInsertRepeat(ByVal copyRow As Integer, ByVal insertRow As Integer, ByVal repeatCount As Integer)
            CopyRowsAndInsertRepeat(copyRow, 1, insertRow, repeatCount)
        End Sub

        ''' <summary>
        ''' 行の削除
        ''' </summary>
        ''' <param name="row">削除する行位置</param>
        ''' <param name="count">削除する行</param>
        ''' <remarks></remarks>
        Public Sub DeleteRow(ByVal row As Integer, Optional ByVal count As Integer = 1)
            Dim xlRow As Object = Nothing

            Try
                xlRow = Me.GetRow(row, row + count - 1)
                xlRow.Delete()

            Catch ex As Exception
                Me.Free(xlRow)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' 新しいセル範囲定義を作成する
        ''' </summary>
        ''' <typeparam name="T">セル範囲に適用する型</typeparam>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="define">定義関数</param>
        ''' <remarks></remarks>
        Public Sub SetUpDefinition(Of T As New)(ByVal startRow As Integer, ByVal startCol As Integer, ByVal define As XlVoRuleBuilder(Of T).Configure)
            Dim definition As New CellRangeDefinition
            With definition
                .ObjectType = GetType(T)
                .StartRow = startRow
                .StartCol = startCol
                .MutableRules = (New XlVoRuleBuilder(Of T)(define)).Rules
            End With
            Me.definition = definition
        End Sub

        Private Sub AssertExistsSameDefinition(Of T)()
            Dim aType As Type = GetType(T)
            If Me.definition Is Nothing Then
                Throw New InvalidOperationException(String.Format("型 {0} のセル範囲定義がされていません. SetUpDefinition()で定義を作成してください.", aType.Name))
            End If
            If Me.definition.ObjectType IsNot aType Then
                Throw New InvalidOperationException(String.Format("型 {0} のセル範囲定義と異なる型 {1} が指定されました. SetUpDefinition()で再定義してください.", Me.definition.ObjectType.Name, aType.Name))
            End If
        End Sub

        ''' <summary>
        ''' 指定されたセル範囲に値を設定します.
        ''' </summary>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="value">設定する値</param>
        ''' <remarks></remarks>
        Public Sub SetValueRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer, ByVal value As Object)

            Dim xlCell As Object = Nothing

            Try
                xlCell = Me.GetCellRC(startRow, startCol, endRow, endCol)

                'データ型に応じて適切なセルタイプを設定する
                '必要な場合, 追記してください.
                If TypeOf value Is String Then
                    Me.SetNumberFormatLocalRC(startRow, startCol, endRow, endCol, "@")
                End If

                xlCell.Value = value

            Finally
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' 指定されたセルに値を設定します.
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <param name="value">設定する値</param>
        ''' <remarks></remarks>
        Public Sub SetValueRC(ByVal axis As XlAxis, ByVal value As Object)
            SetValueRC(axis.Row, axis.Column, value)
        End Sub
        ''' <summary>
        ''' 指定されたセルに値を設定します.
        ''' </summary>
        ''' <param name="col">列インデックス</param>
        ''' <param name="row">行インデックス</param>
        ''' <param name="value">設定する値</param>
        ''' <remarks></remarks>
        Public Sub SetValueRC(ByVal row As Integer, ByVal col As Integer, ByVal value As Object)
            Me.SetValueRC(row, col, -1, -1, value)
        End Sub

        ''' <summary>
        ''' 指定されたキーのセルに値を設定します.
        ''' </summary>
        ''' <param name="key">キー</param>
        ''' <param name="value">設定する値</param>
        ''' <remarks></remarks>
        Public Sub SetValue(ByVal key As String, ByVal value As Object)
            Dim axis As XlAxis = Me.Find(key)
            If axis IsNot Nothing Then
                SetValueRC(axis.Row, axis.Column, value)
            End If
        End Sub

        ''' <summary>
        ''' デリゲートを元に値を設定する
        ''' </summary>
        ''' <typeparam name="T">設定したい値の型</typeparam>
        ''' <param name="startRow">開始行</param>
        ''' <param name="startCol">開始列</param>
        ''' <param name="vos">Vo[]</param>
        ''' <param name="configure">デリゲート</param>
        ''' <remarks></remarks>
        Public Sub SetValuesRC(Of T As New)(startRow As Integer, startCol As Integer, vos As IEnumerable(Of T), configure As XlVoRuleBuilder(Of T).Configure)
            Dim broker As New XlVoBroker(Of T)(configure)
            Me.SetValuesRC(startRow, startCol, broker.ConvertToValues(vos))
        End Sub

        ''' <summary>
        ''' 指定されたセルを起点に値を設定します.
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <param name="values">設定する値[][]</param>
        ''' <remarks></remarks>
        Public Sub SetValuesRC(ByVal axis As XlAxis, ByVal values As Object)
            SetValuesRC(axis.Row, axis.Column, values)
        End Sub
        ''' <summary>
        ''' 指定されたセルを起点に値を設定します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="values">設定する値[][]</param>
        ''' <remarks></remarks>
        Public Sub SetValuesRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal values As Object)
            Dim data As Object(,) = CollectionUtil.ConvToTwoDimensionalArray(values)
            Dim endRow As Integer = data.GetLength(0) + startRow - 1
            Dim endColumn As Integer = data.GetLength(1) + startCol - 1
            Dim xlCell As Object = Me.GetCellRC(startRow, startCol, endRow, endColumn)
            Try
                Me.SetNumberFormatLocalRC(startRow, startCol, endRow, endColumn, "@")
                xlCell.Value = data
            Finally
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' アクティブなセル範囲定義を起点に値を設定する
        ''' </summary>
        ''' <typeparam name="T">値の型</typeparam>
        ''' <param name="vos">設定したい値[]</param>
        ''' <remarks></remarks>
        Public Sub SetValuesRC(Of T As New)(ByVal vos As IEnumerable(Of T))
            AssertExistsSameDefinition(Of T)()

            Dim broker As New XlVoBroker(Of T)(Me.definition.MutableRules)
            Me.SetValuesRC(Me.definition.StartRow, Me.definition.StartCol, broker.ConvertToValues(vos))
        End Sub

        ''' <summary>
        ''' 指定されたセルに設定されている値を取得します.
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <returns>セルに設定されている値</returns>
        ''' <remarks>
        ''' Object 型で返されることに注意してください.
        ''' 文字列として欲しい場合は, GetStr を使用してください.
        ''' </remarks>
        Public Function GetValueRC(ByVal axis As XlAxis) As Object
            Return GetValueRC(axis.Row, axis.Column)
        End Function
        ''' <summary>
        ''' 指定されたセルに設定されている値を取得します.
        ''' </summary>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <returns>セルに設定されている値</returns>
        ''' <remarks>
        ''' Object 型で返されることに注意してください.
        ''' 文字列として欲しい場合は, GetStr を使用してください.
        ''' </remarks>
        Public Function GetValueRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                                   Optional ByVal endRow As Integer = -1, Optional ByVal endCol As Integer = -1) As Object

            Dim xlCell As Object = Nothing

            Try
                xlCell = Me.GetCellRC(startRow, startCol, endRow, endCol)
                Return xlCell.Value

            Finally
                Me.Free(xlCell)
            End Try
        End Function

        ''' <summary>
        ''' セルの値を取得する
        ''' </summary>
        ''' <param name="startColumn">開始列インデックス</param>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="endColumn">終了列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <returns>値[][]</returns>
        ''' <remarks></remarks>
        Public Function GetValuesRC(ByVal startRow As Integer, ByVal startColumn As Integer, _
                                    ByVal endRow As Integer, ByVal endColumn As Integer) As Object()()

            Dim xlCell As Object = Me.GetCellRC(startRow, startColumn, endRow, endColumn)
            Try
                Return CollectionUtil.ConvToTwoJaggedArray(xlCell.Value)

            Finally
                Me.Free(xlCell)
            End Try
        End Function

        ''' <summary>
        ''' デリゲートを元に値を取得する
        ''' </summary>
        ''' <typeparam name="T">データの型</typeparam>
        ''' <param name="startRow">開始行</param>
        ''' <param name="startColumn">開始列</param>
        ''' <param name="count">データ件数</param>
        ''' <param name="configure">デリケート</param>
        ''' <returns>データ[]</returns>
        ''' <remarks></remarks>
        Public Function GetValuesRC(Of T As New)(startRow As Integer, startColumn As Integer, count As Integer, configure As XlVoRuleBuilder(Of T).Configure) As T()
            Dim rule As New XlVoRuleBuilder(Of T)(configure)
            Dim broker As New XlVoBroker(Of T)(configure)

            Dim endRow As Integer = startRow + count - 1
            Dim endCol As Integer = startColumn + rule.Rules.Count - 1
            Dim data As Object()() = Me.GetValuesRC(startRow, startColumn, endRow, endCol)

            Return Enumerable.Range(0, count).Select(Function(index) broker.CreateVo(data(index))).ToArray
        End Function

        ''' <summary>
        ''' アクティブなセル範囲定義をもとに値を取得する（データ件数を指定）
        ''' </summary>
        ''' <typeparam name="T">データの型</typeparam>
        ''' <param name="count">データ件数</param>
        ''' <returns>データ[]</returns>
        ''' <remarks></remarks>
        Public Function GetValuesRC(Of T As New)(ByVal count As Integer) As T()
            AssertExistsSameDefinition(Of T)()

            Dim endRow As Integer = Me.definition.StartRow + count - 1
            Dim endCol As Integer = Me.definition.StartCol + Me.definition.MutableRules.Count - 1
            Dim data As Object()() = Me.GetValuesRC(Me.definition.StartRow, Me.definition.StartCol, endRow, endCol)

            Dim broker As New XlVoBroker(Of T)(Me.definition.MutableRules)
            Return Enumerable.Range(0, count).Select(Function(index) broker.CreateVo(data(index))).ToArray
        End Function

        ''' <summary>
        ''' 指定されたキーのセル+オフセットに設定されている値を取得します.
        ''' </summary>
        ''' <param name="key">キー値</param>
        ''' <returns>セルに設定されている値</returns>
        ''' <param name="offsetCol">オフセット列</param>
        ''' <param name="offsetRow">オフセット行</param>
        ''' <remarks></remarks>
        Public Function GetValue(ByVal key As String, _
                                 Optional ByVal offsetCol As Integer = 1, _
                                 Optional ByVal offsetRow As Integer = 0) As Object

            Dim keyCol As Integer = 0, keyRow As Integer = 0    'キーの行列番号
            Dim xlCell As Object = Nothing

            Try
                'キーの存在する行列を捜す
                If Not Me.Find(key, keyCol, keyRow) Then
                    Return String.Empty
                End If

                xlCell = Me.GetCellRC(keyRow + offsetRow, keyCol + offsetCol)

                If xlCell Is Nothing Then
                    Return String.Empty
                Else
                    Return xlCell.Value
                End If

            Finally
                Me.Free(xlCell)
            End Try
        End Function

        ''' <summary>
        ''' 指定されたセルに設定されている値を文字列で取得します.
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <returns>セルに設定されている値</returns>
        ''' <remarks></remarks>
        Public Function GetValueAsStringRC(ByVal axis As XlAxis) As String
            Return GetValueAsStringRC(axis.Row, axis.Column)
        End Function
        ''' <summary>
        ''' 指定されたセルに設定されている値を文字列で取得します.
        ''' </summary>
        ''' <param name="row">行インデックス</param>
        ''' <param name="col">列インデックス</param>
        ''' <returns>セルに設定されている値</returns>
        ''' <remarks></remarks>
        Public Function GetValueAsStringRC(ByVal row As Integer, ByVal col As Integer) As String
            Dim xlCell As Object = Nothing

            Try
                xlCell = Me.GetCellRC(row, col)

                If xlCell.Value Is Nothing Then
                    Return String.Empty
                Else
                    Return xlCell.Value.ToString()
                End If

            Finally
                Me.Free(xlCell)
            End Try
        End Function

        ''' <summary>
        ''' 指定されたキーのセル+オフセットに設定されている値を文字列形式で取得します.
        ''' </summary>
        ''' <param name="key">キー値</param>
        ''' <returns>セルに設定されている値</returns>
        ''' <param name="offsetRow">オフセット行</param>
        ''' <param name="offsetCol">オフセット列</param>
        ''' <remarks></remarks>
        Public Function GetValueAsStringRC(ByVal key As String, _
                                           Optional ByVal offsetRow As Integer = 0, _
                                           Optional ByVal offsetCol As Integer = 1) As String

            Dim value As Object = Me.GetValue(key, offsetCol, offsetRow)

            If value Is Nothing Then
                Return String.Empty
            Else
                Return value.ToString()
            End If
        End Function

        ''' <summary>
        ''' オートシェイプテキスト設定
        ''' </summary>
        ''' <param name="shapeName"></param>
        ''' <param name="value"></param>
        ''' <remarks></remarks>
        Public Sub SetShapeValue(ByVal shapeName As String, ByVal value As Object)
            Dim xlsShapes As Object = Nothing
            Dim xlsShape As Object = Nothing
            Dim xlsTextFrame As Object = Nothing
            Dim xlsCharacters As Object = Nothing

            Try
                xlsShapes = m_activeSheet.Shapes
                xlsShape = xlsShapes.Item(shapeName)
                xlsTextFrame = xlsShape.TextFrame
                xlsCharacters = xlsTextFrame.Characters
                xlsCharacters.Text = value

            Catch ex As Exception
                Me.Free(xlsCharacters)
                Me.Free(xlsTextFrame)
                Me.Free(xlsShape)
                Me.Free(xlsShapes)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' オートシェイプの表示・非表示設定
        ''' </summary>
        ''' <param name="shapeName"></param>
        ''' <param name="visible"></param>
        ''' <remarks></remarks>
        Public Sub SetShapeVisible(ByVal shapeName As String, ByVal visible As Boolean)
            Dim xlsShapes As Object = Nothing
            Dim xlsShape As Object = Nothing

            Try
                xlsShapes = m_activeSheet.Shapes
                xlsShape = xlsShapes.Item(shapeName)
                xlsShape.Visible = visible

            Finally
                Me.Free(xlsShape)
                Me.Free(xlsShapes)
            End Try
        End Sub

        ''' <summary>
        ''' 指定されたキー値を元にセルを検索して返します.
        ''' </summary>
        ''' <param name="key">探索するキー値</param>
        ''' <param name="lookAt">検索方法</param>
        ''' <param name="searchOrder">検索方向</param>
        ''' <param name="searchDirection">検索順序</param>
        ''' <param name="matchCase">True - 大文字と小文字を区別する, False - 区別しない.</param>
        ''' <param name="matchByte">True - 半角と全角を区別する, False - 区別しない.</param>
        ''' <param name="lookIn">検索対象</param>
        ''' <returns>最初に見つかった座標</returns>
        ''' <remarks></remarks>
        Public Function Find(ByVal key As String, _
                                Optional ByVal lookAt As XlLookAt = XlLookAt.xlWhole, _
                                Optional ByVal searchOrder As XlSearchOrder = XlSearchOrder.xlByRows, _
                                Optional ByVal searchDirection As XlSearchDirection = XlSearchDirection.xlNext, _
                                Optional ByVal matchCase As Boolean = False, _
                                Optional ByVal matchByte As Boolean = False, _
                                Optional ByVal lookIn As XlFindLookIn = XlFindLookIn.xlValues) As XlAxis

            Dim xlCells As Object = m_activeSheet.Cells
            Try
                Dim xlCell As Object = xlCells.Find(key, , lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte)
                Try
                    If xlCell Is Nothing Then
                        Return Nothing
                    End If
                    Return New XlAxis(xlCell.Row, xlCell.Column)
                Finally
                    Me.Free(xlCell)
                End Try
            Finally
                Me.Free(xlCells)
            End Try
        End Function

        ''' <summary>
        ''' 指定されたキー値を元にセルをして行列番号を返します.
        ''' </summary>
        ''' <param name="key">探索するキー値</param>
        ''' <param name="col">列番号[戻り値]</param>
        ''' <param name="row">行番号[戻り値]</param>
        ''' <param name="lookAt">検索方法</param>
        ''' <param name="searchOrder">検索方向</param>
        ''' <param name="searchDirection">検索順序</param>
        ''' <param name="matchCase">True - 大文字と小文字を区別する, False - 区別しない.</param>
        ''' <param name="matchByte">True - 半角と全角を区別する, False - 区別しない.</param>
        ''' <param name="lookIn">検索対象</param>
        ''' <returns>見つかった場合 True, 見つからなかった場合 False を返します.</returns>
        ''' <remarks></remarks>
        Public Function Find(ByVal key As String, _
                             ByRef col As Integer, ByRef row As Integer, _
                             Optional ByVal lookAt As XlLookAt = XlLookAt.xlWhole, _
                             Optional ByVal searchOrder As XlSearchOrder = XlSearchOrder.xlByRows, _
                             Optional ByVal searchDirection As XlSearchDirection = XlSearchDirection.xlNext, _
                             Optional ByVal matchCase As Boolean = False, _
                             Optional ByVal matchByte As Boolean = False, _
                             Optional ByVal lookIn As XlFindLookIn = XlFindLookIn.xlValues) As Boolean

            Dim axis As XlAxis = Me.Find(key, lookAt, searchOrder, searchDirection, matchCase, matchByte, lookIn)
            If axis Is Nothing Then
                col = -1
                row = -1

                Return False
            Else
                col = axis.Column
                row = axis.Row

                Return True
            End If
        End Function

        ''' <summary>
        ''' Findの条件で「次」を検索する
        ''' </summary>
        ''' <param name="after">「次」を検索する基点</param>
        ''' <returns>次に見つかった座標</returns>
        ''' <remarks></remarks>
        Private Function FindNext(after As XlAxis) As XlAxis
            Dim xlCells As Object = m_activeSheet.Cells
            Try
                Dim afterCell As Object = If(after Is Nothing, Nothing, GetCellRC(after))
                Try
                    Dim xlCell As Object = If(after Is Nothing, xlCells.FindNext(), xlCells.FindNext(afterCell))
                    Try
                        If xlCell Is Nothing Then
                            Return Nothing
                        End If
                        Return New XlAxis(xlCell.Row, xlCell.Column)
                    Finally
                        Me.Free(xlCell)
                    End Try
                Finally
                    Free(afterCell)
                End Try
            Finally
                Me.Free(xlCells)
            End Try
        End Function

        ''' <summary>
        ''' すべて検索する
        ''' </summary>
        ''' <param name="key">条件</param>
        ''' <param name="lookAt">検索方法</param>
        ''' <param name="searchOrder">検索方向</param>
        ''' <param name="searchDirection">検索順序</param>
        ''' <param name="matchCase">True - 大文字と小文字を区別する, False - 区別しない.</param>
        ''' <param name="matchByte">True - 半角と全角を区別する, False - 区別しない.</param>
        ''' <param name="lookIn">検索対象</param>
        ''' <returns>見つかった座標[]</returns>
        ''' <remarks></remarks>
        Public Function FindAll(ByVal key As String, _
                                Optional ByVal lookAt As XlLookAt = XlLookAt.xlWhole, _
                                Optional ByVal searchOrder As XlSearchOrder = XlSearchOrder.xlByRows, _
                                Optional ByVal searchDirection As XlSearchDirection = XlSearchDirection.xlNext, _
                                Optional ByVal matchCase As Boolean = False, _
                                Optional ByVal matchByte As Boolean = False, _
                                Optional ByVal lookIn As XlFindLookIn = XlFindLookIn.xlValues) As XlAxis()
            Dim results As New List(Of XlAxis)
            Dim axis As XlAxis = Me.Find(key, lookAt, searchOrder, searchDirection, matchCase, matchByte, lookIn)
            If axis Is Nothing Then
                Return results.ToArray
            End If
            Do
                results.Add(axis)
                axis = Me.FindNext(axis)
            Loop Until results.Contains(axis)
            Return results.ToArray
        End Function

        ''' <summary>
        ''' 列幅を取得する
        ''' </summary>
        ''' <param name="columnIndex">開始列インデックス</param>
        ''' <remarks></remarks>
        Public Function GetColumnWidth(ByVal columnIndex As Integer) As Double
            Dim xlColumn As Object = Me.GetColumn(columnIndex)
            Try
                Return xlColumn.ColumnWidth
            Finally
                Me.Free(xlColumn)
            End Try
        End Function

        ''' <summary>
        ''' 指定した列範囲の幅を設定します.
        ''' </summary>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="colWidth">幅(ピクセルではなく,エクセルの値)</param>
        ''' <remarks></remarks>
        Public Sub SetColumnWidth(ByVal startCol As Integer, ByVal endCol As Integer, ByVal colWidth As Double)
            Dim xlColumn As Object = Nothing

            Try
                xlColumn = Me.GetColumn(startCol, endCol)
                xlColumn.ColumnWidth = colWidth

            Finally
                Me.Free(xlColumn)
            End Try
        End Sub

        ''' <summary>
        ''' 指定した列の幅を設定します.
        ''' </summary>
        ''' <param name="col">列インデックス</param>
        ''' <param name="colWidth">幅(ピクセルではなく,エクセルの値)</param>
        ''' <remarks></remarks>
        Public Sub SetColumnWidth(ByVal col As Integer, ByVal colWidth As Double)
            Me.SetColumnWidth(col, -1, colWidth)
        End Sub

        ''' <summary>
        ''' 指定した列の幅を自動的に設定します.
        ''' </summary>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endCol">終了列インデックス[省略可]</param>
        ''' <remarks></remarks>
        Public Sub AutoFitCol(ByVal startCol As Integer, Optional ByVal endCol As Integer = -1)
            Dim xlColumn As Object = Nothing

            Try
                xlColumn = Me.GetColumn(startCol, endCol)
                xlColumn.AutoFit()

            Finally
                Me.Free(xlColumn)
            End Try
        End Sub

        ''' <summary>
        ''' 行高さを取得する
        ''' </summary>
        ''' <param name="rowIndex">行インデックス</param>
        ''' <remarks></remarks>
        Public Function GetRowHeight(ByVal rowIndex As Integer) As Double
            Dim xlRow As Object = Me.GetRow(rowIndex)
            Try
                Return xlRow.RowHeight
            Finally
                Me.Free(xlRow)
            End Try
        End Function

        ''' <summary>
        ''' 指定した行範囲の高さを設定します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="rowHeight">高さ(ピクセルではなく,エクセルの値)</param>
        ''' <remarks></remarks>
        Public Sub SetRowHeight(ByVal startRow As Integer, ByVal endRow As Integer, ByVal rowHeight As Double)
            Dim xlRow As Object = Nothing

            Try
                xlRow = Me.GetRow(startRow, endRow)
                xlRow.RowHeight = rowHeight

            Finally
                Me.Free(xlRow)
            End Try
        End Sub

        ''' <summary>
        ''' 指定した行の高さを設定します.
        ''' </summary>
        ''' <param name="row">行インデックス</param>
        ''' <param name="rowHeight">高さ(ピクセルではなく,エクセルの値)</param>
        ''' <remarks></remarks>
        Public Sub SetRowHeight(ByVal row As Integer, ByVal rowHeight As Double)
            Me.SetRowHeight(row, -1, rowHeight)
        End Sub

        ''' <summary>
        ''' 指定した行の高さを自動的に設定します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="endRow">終了行インデックス[省略可]</param>
        ''' <remarks></remarks>
        Public Sub AutoFitRow(ByVal startRow As Integer, Optional ByVal endRow As Integer = -1)
            Dim xlRow As Object = Nothing

            Try
                xlRow = Me.GetRow(startRow, endRow)
                xlRow.AutoFit()

            Finally
                Me.Free(xlRow)
            End Try
        End Sub

        ''' <summary>
        ''' 配置書式を取得する
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAlignmentRC(ByVal axis As XlAxis) As Alignment
            Return GetAlignmentRC(axis.Row, axis.Column)
        End Function
        ''' <summary>
        ''' 配置書式を取得する
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス 省略可</param>
        ''' <param name="endCol">終了列インデックス 省略可</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAlignmentRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                                       Optional ByVal endRow As Integer = -1, Optional ByVal endCol As Integer = -1) As Alignment
            Dim xlRange As Object = Me.GetCellRC(startRow, startCol, endRow, endCol)
            Try
                Return New Alignment With {.HorizontalAlignment = xlRange.HorizontalAlignment, _
                                           .VerticalAlignment = xlRange.VerticalAlignment, _
                                           .IndentLevel = xlRange.IndentLevel, _
                                           .WrapText = xlRange.WrapText, _
                                           .ShrinkToFit = xlRange.ShrinkToFit, _
                                           .MergeCells = xlRange.MergeCells, _
                                           .Orientation = xlRange.Orientation, _
                                           .AddIndent = xlRange.AddIndent}
            Finally
                Free(xlRange)
            End Try
        End Function

        ''' <summary>
        ''' 配置書式を設定する
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <param name="align">配置書式</param>
        ''' <remarks></remarks>
        Public Sub SetAlignmentRC(ByVal axis As XlAxis, ByVal align As Alignment)
            SetAlignmentRC(axis.Row, axis.Column, align)
        End Sub
        ''' <summary>
        ''' 配置書式を設定する
        ''' </summary>
        ''' <param name="row">行index</param>
        ''' <param name="col">列index</param>
        ''' <param name="align">配置書式</param>
        ''' <remarks></remarks>
        Public Sub SetAlignmentRC(ByVal row As Integer, ByVal col As Integer, ByVal align As Alignment)
            SetAlignmentRC(row, col, -1, -1, align)
        End Sub

        ''' <summary>
        ''' 配置書式を設定する
        ''' </summary>
        ''' <param name="startRow">開始行index</param>
        ''' <param name="startCol">開始列index</param>
        ''' <param name="endRow">終了行index</param>
        ''' <param name="endCol">終了列index</param>
        ''' <param name="align">配置書式</param>
        ''' <remarks></remarks>
        Public Sub SetAlignmentRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                                  ByVal endRow As Integer, ByVal endCol As Integer, ByVal align As Alignment)
            Dim xlRange As Object = Me.GetCellRC(startRow, startCol, endRow, endCol)
            Try
                PerformSetAlignmentTo(xlRange, align)
            Finally
                Free(xlRange)
            End Try
        End Sub

        Private Sub PerformSetAlignmentTo(ByVal xlRange As Object, ByVal align As Alignment)
            If align.HorizontalAlignment <> 0 Then
                xlRange.HorizontalAlignment = align.HorizontalAlignment
            End If
            If align.VerticalAlignment <> 0 Then
                xlRange.VerticalAlignment = align.VerticalAlignment
            End If
            If align.IndentLevel.HasValue Then
                xlRange.IndentLevel = align.IndentLevel.Value
            End If
            If align.WrapText.HasValue Then
                xlRange.WrapText = align.WrapText.Value
            End If
            If align.ShrinkToFit.HasValue Then
                xlRange.ShrinkToFit = align.ShrinkToFit.Value
            End If
            If align.MergeCells.HasValue Then
                xlRange.MergeCells = align.MergeCells.Value
            End If
            If align.Orientation <> 0 Then
                xlRange.Orientation = align.Orientation
            End If
            If align.AddIndent.HasValue Then
                xlRange.AddIndent = align.AddIndent.Value
            End If
        End Sub

        ''' <summary>
        ''' 配置書式を設定する
        ''' </summary>
        ''' <param name="col">列index</param>
        ''' <param name="align">配置書式</param>
        ''' <remarks></remarks>
        Public Sub SetAlignmentColumn(ByVal col As Integer, ByVal align As Alignment)
            SetAlignmentColumn(col, -1, align)
        End Sub

        ''' <summary>
        ''' 配置書式を設定する
        ''' </summary>
        ''' <param name="startCol">開始列index</param>
        ''' <param name="endCol">終了列index</param>
        ''' <param name="align">配置書式</param>
        ''' <remarks></remarks>
        Public Sub SetAlignmentColumn(ByVal startCol As Integer, ByVal endCol As Integer, ByVal align As Alignment)
            Dim xlRange As Object = Me.GetColumn(startCol, endCol)
            Try
                PerformSetAlignmentTo(xlRange, align)
            Finally
                Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 指定されたセル範囲を結合します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="merge">True - 結合する, False - 結合を解除</param>
        ''' <remarks></remarks>
        Public Sub MergeCellsRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                              ByVal endRow As Integer, ByVal endCol As Integer, ByVal merge As Boolean)
            Me.SetAlignmentRC(startRow, startCol, endRow, endCol, New Alignment With {.MergeCells = merge})
        End Sub

        ''' <summary>
        ''' 指定されたセル範囲を列単位で結合します.
        ''' </summary>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="merge">True - 結合する, False - 結合を解除</param>
        ''' <remarks></remarks>
        Public Sub MergeCellsColUnit(ByVal startCol As Integer, ByVal startRow As Integer, _
                                     ByVal endCol As Integer, ByVal endRow As Integer, ByVal merge As Boolean)

            Dim xlRange As Object = Nothing
            Dim rangeStr As String = String.Empty

            Const BAT_COLS As Integer = 20     '一度に処理する列数

            For i As Integer = startCol To endCol
                rangeStr &= String.Format("{0}{1}:{0}{2}", _
                                          ConvertToLetter(i), startRow, endRow)

                If Not i Mod BAT_COLS = 0 And Not i = endCol Then
                    rangeStr &= ","
                End If

                'この方法で指定する場合, 30個が限界？なため
                '20個区切りで結合を行う.
                If i Mod BAT_COLS = 0 Or i = endCol Then
                    Try
                        xlRange = Me.GetCell(rangeStr)
                        xlRange.MergeCells = merge
                        rangeStr = String.Empty

                    Finally
                        Me.Free(xlRange)
                    End Try
                End If
            Next
        End Sub

        ''' <summary>
        ''' シート全体のフォントを設定します.
        ''' </summary>
        ''' <param name="fontFamily">フォント名</param>
        ''' <param name="size">フォントサイズ</param>
        ''' <param name="color">RGB関数の値, または, 16進数RGB文字列(#FFFFFF)[省略可]</param>
        ''' <param name="bold">太字[省略可]</param>
        ''' <param name="italic">斜体[省略可]</param>
        ''' <param name="strike">取り消し線[省略可]</param>
        ''' <param name="subscript">下付き文字[省略可]</param>
        ''' <param name="superscript">上付き文字[省略可]</param>
        ''' <param name="shadow">影付き[省略可]</param>
        ''' <remarks></remarks>
        Public Sub SetFont(ByVal fontFamily As String, ByVal size As Integer, _
                           Optional ByVal color As Object = Nothing, _
                           Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, _
                           Optional ByVal strike As Boolean = False, _
                           Optional ByVal subscript As Boolean = False, Optional ByVal superscript As Boolean = False, _
                           Optional ByVal shadow As Boolean = False)

            Dim xlRange As Object = Nothing
            Dim xlFont As Object = Nothing

            Try
                xlRange = m_activeSheet.Cells
                xlFont = xlRange.Font
                xlFont.Name = fontFamily
                xlFont.Size = size
                If Not color Is Nothing Then xlFont.Color = color
                xlFont.Bold = bold
                xlFont.Italic = italic
                xlFont.Strikethrough = strike
                xlFont.Subscript = subscript
                xlFont.Superscript = superscript
                xlFont.Shadow = shadow

            Finally
                Me.Free(xlRange)
                Me.Free(xlFont)
            End Try
        End Sub

        ''' <summary>
        ''' 指定されたセル範囲のフォントを設定します.
        ''' </summary>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="fontFamily">フォント名</param>
        ''' <param name="size">フォントサイズ</param>
        ''' <param name="color">RGB関数の値, または, 16進数RGB文字列(#FFFFFF)[省略可]</param>
        ''' <param name="bold">太字[省略可]</param>
        ''' <param name="italic">斜体[省略可]</param>
        ''' <param name="strike">取り消し線[省略可]</param>
        ''' <param name="subscript">下付き文字[省略可]</param>
        ''' <param name="superscript">上付き文字[省略可]</param>
        ''' <param name="shadow">影付き[省略可]</param>
        ''' <remarks></remarks>
        Public Sub SetFontRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                             ByVal endRow As Integer, ByVal endCol As Integer, _
                             ByVal fontFamily As String, ByVal size As Integer, _
                             Optional ByVal color As Object = Nothing, Optional _
                             ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, _
                             Optional ByVal strike As Boolean = False, _
                             Optional ByVal subscript As Boolean = False, Optional ByVal superscript As Boolean = False, _
                             Optional ByVal shadow As Boolean = False)

            Dim xlRange As Object = Nothing
            Dim xlFont As Object = Nothing

            Try
                xlRange = Me.GetCellRC(startRow, startCol, endRow, endCol)
                xlFont = xlRange.Font
                xlFont.Name = fontFamily
                xlFont.Size = size
                If Not color Is Nothing Then xlFont.Color = color
                xlFont.Bold = bold
                xlFont.Italic = italic
                xlFont.Strikethrough = strike
                xlFont.Subscript = subscript
                xlFont.Superscript = superscript
                xlFont.Shadow = shadow

            Finally
                Me.Free(xlFont)
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 指定されたセルのフォントを設定します.
        ''' </summary>
        ''' <param name="col">列インデックス</param>
        ''' <param name="row">行インデックス</param>
        ''' <param name="fontFamily">フォント名</param>
        ''' <param name="size">フォントサイズ</param>
        ''' <param name="color">RGB関数の値, または, 16進数RGB文字列(#FFFFFF)[省略可]</param>
        ''' <param name="bold">太字[省略可]</param>
        ''' <param name="italic">斜体[省略可]</param>
        ''' <param name="strike">取り消し線[省略可]</param>
        ''' <param name="subscript">下付き文字[省略可]</param>
        ''' <param name="superscript">上付き文字[省略可]</param>
        ''' <param name="shadow">影付き[省略可]</param>
        ''' <remarks></remarks>
        Public Sub SetFontRC(ByVal row As Integer, ByVal col As Integer, _
                             ByVal fontFamily As String, ByVal size As Integer, _
                             Optional ByVal color As Object = Nothing, _
                             Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, _
                             Optional ByVal strike As Boolean = False, _
                             Optional ByVal subscript As Boolean = False, Optional ByVal superscript As Boolean = False, _
                             Optional ByVal shadow As Boolean = False)

            Me.SetFontRC(row, col, -1, -1, fontFamily, size, color, bold, italic, strike, subscript, superscript, shadow)
        End Sub

        ''' <summary>
        ''' 指定されたセル範囲の背景色を設定します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="color">RGB関数の値, または, 16進数RGB文字列(#FFFFFF)</param>
        ''' <remarks></remarks>
        Public Sub SetBackColorRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer, ByVal color As Object)
            Dim xlRange As Object = Nothing
            Dim xlInterior As Object = Nothing

            Try
                xlRange = Me.GetCellRC(startRow, startCol, endRow, endCol)
                xlInterior = xlRange.Interior
                xlInterior.Color = color

            Finally
                Me.Free(xlInterior)
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 指定されたセルの背景色を設定します.
        ''' </summary>
        ''' <param name="row">開始行インデックス</param>
        ''' <param name="col">開始列インデックス</param>
        ''' <param name="color">RGB関数の値, または, 16進数RGB文字列(#FFFFFF)</param>
        ''' <remarks></remarks>
        Public Sub SetBackColorRC(ByVal row As Integer, ByVal col As Integer, ByVal color As Object)
            Me.SetBackColorRC(row, col, -1, -1, color)
        End Sub

        ''' <summary>
        ''' 指定されたセル範囲の背景色を設定します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="colorIndex">カラーインデックス</param>
        ''' <remarks></remarks>
        Public Sub SetBackColorIndexRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer, ByVal colorIndex As Integer)
            Dim xlRange As Object = Nothing
            Dim xlInterior As Object = Nothing

            Try
                xlRange = Me.GetCellRC(startRow, startCol, endRow, endCol)
                xlInterior = xlRange.Interior
                xlInterior.ColorIndex = colorIndex

            Finally
                Me.Free(xlInterior)
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 指定されたセルの背景色を設定します.
        ''' </summary>
        ''' <param name="row">開始行インデックス</param>
        ''' <param name="col">開始列インデックス</param>
        ''' <param name="colorIndex">カラーインデックス</param>
        ''' <remarks></remarks>
        Public Sub SetBackColorIndexRC(ByVal row As Integer, ByVal col As Integer, ByVal colorIndex As Integer)
            Me.SetBackColorIndexRC(row, col, -1, -1, colorIndex)
        End Sub

        ''' <summary>
        ''' シート全体の表示形式を設定します.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SetFont(ByVal formatStr As String)
            Dim xlRange As Object = Nothing

            Try
                xlRange = m_activeSheet.Cells
                xlRange.NumberFormatLocal = formatStr

            Finally
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 指定された列範囲の表示形式を設定します.
        ''' </summary>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="formatStr">書式指定文字列(詳しくはVBAのヘルプ参照)</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetNumberFormatLocalCol(ByVal startCol As Integer, ByVal endCol As Integer, ByVal formatStr As String)
            Dim xlColumn As Object = Nothing

            Try
                xlColumn = Me.GetColumn(startCol, endCol)
                xlColumn.NumberFormatLocal = formatStr

            Finally
                Me.Free(xlColumn)
            End Try
        End Sub

        ''' <summary>
        ''' 指定された列の表示形式を設定します.
        ''' </summary>
        ''' <param name="col">列インデックス</param>
        ''' <param name="formatStr">書式指定文字列(詳しくはVBAのヘルプ参照)</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetNumberFormatLocalCol(ByVal col As Integer, ByVal formatStr As String)
            Me.SetNumberFormatLocalCol(col, -1, formatStr)
        End Sub

        ''' <summary>
        ''' 指定された行範囲の表示形式を設定します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="formatStr">書式指定文字列(詳しくはVBAのヘルプ参照)</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetNumberFormatLocalRow(ByVal startRow As Integer, ByVal endRow As Integer, ByVal formatStr As String)
            Dim xlRow As Object = Nothing

            Try
                xlRow = Me.GetRow(startRow, endRow)
                xlRow.NumberFormatLocal = formatStr

            Finally
                Me.Free(xlRow)
            End Try
        End Sub

        ''' <summary>
        ''' 指定された行の表示形式を設定します.
        ''' </summary>
        ''' <param name="row">行インデックス</param>
        ''' <param name="formatStr">書式指定文字列(詳しくはVBAのヘルプ参照)</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetNumberFormatLocalRow(ByVal row As Integer, ByVal formatStr As String)
            Me.SetNumberFormatLocalRow(row, -1, formatStr)
        End Sub

        ''' <summary>
        ''' 指定されたセル範囲の表示形式を設定します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="formatStr">書式指定文字列(詳しくはVBAのヘルプ参照)</param>
        ''' <remarks></remarks>
        Public Sub SetNumberFormatLocalRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer, ByVal formatStr As String)
            Dim xlCell As Object = Nothing

            Try
                xlCell = Me.GetCellRC(startRow, startCol, endRow, endCol)
                xlCell.NumberFormatLocal = formatStr

            Finally
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' 指定されたセル範囲の表示形式を設定します.
        ''' </summary>
        ''' <param name="col">列インデックス</param>
        ''' <param name="row">行インデックス</param>
        ''' <param name="formatStr">書式指定文字列(詳しくはVBAのヘルプ参照)</param>
        ''' <remarks></remarks>
        Public Sub SetNumberFormatLocalRC(ByVal row As Integer, ByVal col As Integer, ByVal formatStr As String)
            Me.SetNumberFormatLocalRC(row, col, -1, -1, formatStr)
        End Sub

        ''' <summary>
        ''' コメントの設定
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <param name="text">コメント</param>
        ''' <remarks></remarks>
        Public Sub SetCommentRC(ByVal axis As XlAxis, ByVal text As String)
            SetCommentRC(axis.Row, axis.Column, text)
        End Sub
        ''' <summary>
        ''' コメントの設定
        ''' </summary>
        ''' <param name="startRow">開始行</param>
        ''' <param name="startCol">開始列</param>
        ''' <param name="text">コメント</param>
        ''' <remarks></remarks>
        Public Sub SetCommentRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal text As String)
            SetCommentRC(startRow, startCol, -1, -1, text)
        End Sub
        ''' <summary>
        ''' コメントの設定
        ''' </summary>
        ''' <param name="startRow">開始行</param>
        ''' <param name="startCol">開始列</param>
        ''' <param name="endRow">終了行</param>
        ''' <param name="endCol">終了列</param>
        ''' <param name="text">コメント</param>
        ''' <remarks></remarks>
        Public Sub SetCommentRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                                ByVal endRow As Integer, ByVal endCol As Integer, ByVal text As String)

            Dim xlCell As Object = Nothing
            Dim xlComment As Object = Nothing
            Dim xlShape As Object = Nothing

            Try
                xlCell = Me.GetCellRC(startRow, startCol, endRow, endCol)
                xlCell.AddComment(text)
                xlComment = xlCell.Comment

                '結合されたセルの最左上セル以外ではコメントを取得できないため
                If Not xlComment Is Nothing Then
                    xlShape = xlComment.Shape

                    ''TODO: 文字列の長さ,改行に応じて幅を設定できるようにする.
                    xlShape.ScaleHeight(0.48, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft)
                End If

            Finally
                Me.Free(xlShape)
                Me.Free(xlComment)
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' コメントを取得する
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <returns>取得したコメント</returns>
        ''' <remarks></remarks>
        Public Function GetCommentRC(axis As XlAxis) As String
            Return GetCommentRC(axis.Row, axis.Column)
        End Function
        ''' <summary>
        ''' コメントを取得する
        ''' </summary>
        ''' <param name="row">行index</param>
        ''' <param name="column">列index</param>
        ''' <returns>取得したコメント</returns>
        ''' <remarks></remarks>
        Public Function GetCommentRC(ByVal row As Integer, ByVal column As Integer) As String
            Dim xlCell As Object = Me.GetCellRC(row, column)
            Try
                Dim xlComment As Object = xlCell.Comment
                Try
                    If xlComment Is Nothing Then
                        Return Nothing
                    End If
                    Return xlComment.Text
                Finally
                    Me.Free(xlComment)
                End Try
            Finally
                Me.Free(xlCell)
            End Try
        End Function

        ''' <summary>
        ''' ウィンドウの固定
        ''' </summary>
        ''' <param name="col"></param>
        ''' <param name="row"></param>
        ''' <param name="status"></param>
        ''' <remarks></remarks>
        Public Sub FreezePanesRC(ByVal row As Integer, ByVal col As Integer, ByVal status As Boolean)
            Dim xlCell As Object = Nothing
            Dim xlWin As Object = Nothing

            Try
                xlCell = Me.GetCellRC(row, col)
                xlCell.Select()
                xlWin = m_xlApp.ActiveWindow

                xlWin.FreezePanes = True

            Finally
                Me.Free(xlWin)
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' オートフィルタ
        ''' </summary>
        ''' <param name="startRow">開始行</param>
        ''' <param name="startCol">開始列</param>
        ''' <param name="endRow">終了行</param>
        ''' <param name="endCol">終了列</param>
        ''' <remarks></remarks>
        Public Sub AutoFilterRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                                Optional ByVal endRow As Integer = -1, Optional ByVal endCol As Integer = -1)

            Dim xlRange As Object = Nothing

            Try
                xlRange = Me.GetCellRC(startRow, startCol, endRow, endCol)
                xlRange.AutoFilter()

            Finally
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' シートのオートフィルター設定を解除する
        ''' </summary>
        ''' <param name="sheetIndex">シートindex</param>
        ''' <remarks></remarks>
        Public Sub ClearAutoFilterMode(ByVal sheetIndex As Integer)
            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                Dim xlSheet As Object = xlSheets.Item(sheetIndex)
                xlSheet.AutoFilterMode = False
            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' シートがオートフィルター中かを返す
        ''' </summary>
        ''' <param name="sheetIndex">シートindex</param>
        ''' <remarks></remarks>
        Public Function IsAutoFilterMode(ByVal sheetIndex As Integer) As Boolean
            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                Dim xlSheet As Object = xlSheets.Item(sheetIndex)
                Return xlSheet.AutoFilterMode
            Finally
                Me.Free(xlSheets)
            End Try
        End Function

        ''' <summary>
        ''' コピー
        ''' </summary>
        ''' <param name="startCol">開始列</param>
        ''' <param name="startRow">開始行</param>
        ''' <param name="endCol">終了列</param>
        ''' <param name="endRow">終了行</param>
        ''' <remarks></remarks>
        Public Sub CopyRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                          Optional ByVal endRow As Integer = -1, Optional ByVal endCol As Integer = -1)

            Dim xlRange As Object = Nothing

            Try
                xlRange = Me.GetCellRC(startRow, startCol, endRow, endCol)
                xlRange.Copy()

            Finally
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 図のリンク貼付
        ''' </summary>
        ''' <param name="row"></param>
        ''' <param name="col"></param>
        ''' <param name="shapeName"></param>
        ''' <remarks></remarks>
        Public Sub LinkPasteRC(ByVal row As Integer, ByVal col As Integer, ByVal shapeName As String)
            Dim xlCell As Object = Nothing
            Dim xlPictures As Object = Nothing
            Dim xlPicture As Object = Nothing

            Try
                xlCell = Me.GetCellRC(row, col)
                xlCell.Select()

                xlPictures = m_activeSheet.Pictures
                xlPicture = xlPictures.Paste(Link:=True)

                xlPicture.Name = shapeName

            Finally
                Me.Free(xlCell)
                Me.Free(xlPicture)
                Me.Free(xlPictures)
            End Try
        End Sub

        ''' <summary>
        ''' 図の位置及び大きさを取得
        ''' </summary>
        ''' <param name="shapeName"></param>
        ''' <param name="retSize"></param>
        ''' <param name="retLocation"></param>
        ''' <remarks></remarks>
        Public Sub GetShapeBounds(ByVal shapeName As String, ByRef retSize As System.Drawing.Size, _
                                                             ByRef retLocation As System.Drawing.Point)

            Dim xlShapes As Object = Nothing
            Dim xlShape As Object = Nothing

            Try
                xlShapes = m_activeSheet.Shapes
                xlShape = xlShapes.Item(shapeName)

                retSize.Width = xlShape.Width
                retSize.Height = xlShape.Height
                retLocation.X = xlShape.Left
                retLocation.Y = xlShape.Top

            Finally
                Me.Free(xlShape)
                Me.Free(xlShapes)
            End Try
        End Sub

        ''' <summary>
        ''' シェイプのサイズや位置を設定する
        ''' </summary>
        ''' <param name="shapeName">シェイプ名</param>
        ''' <param name="size">サイズ</param>
        ''' <param name="location">位置</param>
        ''' <remarks></remarks>
        Public Sub SetShapeBounds(ByVal shapeName As String, ByVal size As System.Drawing.Size, _
                                                             ByVal location As System.Drawing.Point)
            Dim xlShapes As Object = Nothing
            Dim xlShape As Object = Nothing

            Try
                xlShapes = m_activeSheet.Shapes
                xlShape = xlShapes.Item(shapeName)

                If size.Width > -1 Then xlShape.Width = size.Width
                If size.Height > -1 Then xlShape.Height = size.Height

                If location.X > -1 Then xlShape.Left = location.X
                If location.Y > -1 Then xlShape.Top = location.Y

            Finally
                Me.Free(xlShape)
                Me.Free(xlShapes)
            End Try
        End Sub

        ''' <summary>
        ''' 指定された列範囲を取得します.
        ''' </summary>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endCol">終了列インデックス[省略可]</param>
        ''' <returns>列オブジェクト. エラーの場合 Nothing を返します.</returns>
        ''' <remarks>列範囲オブジェクト. 失敗した場合 Nothing を返します.</remarks>
        Protected Function GetColumn(ByVal startCol As Integer, Optional ByVal endCol As Integer = -1) As Object
            Dim xlColumns As Object = Nothing
            Dim xlColFrom As Object = Nothing
            Dim xlColTo As Object = Nothing

            Try
                xlColumns = m_activeSheet.Columns
                xlColFrom = xlColumns.Item(startCol)

                If endCol = -1 Then
                    Return m_activeSheet.Range(xlColFrom, xlColFrom)
                Else
                    xlColTo = xlColumns.Item(endCol)
                    Return m_activeSheet.Range(xlColFrom, xlColTo)
                End If

            Finally
                Me.Free(xlColTo)
                Me.Free(xlColFrom)
                Me.Free(xlColumns)
            End Try
        End Function

        ''' <summary>
        ''' 指定された行範囲を取得します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="endRow">終了行インデックス[省略可]</param>
        ''' <returns>行オブジェクト. エラーの場合 Nothing を返します.</returns>
        ''' <remarks></remarks>
        Protected Function GetRow(ByVal startRow As Integer, Optional ByVal endRow As Integer = -1) As Object
            Dim xlRows As Object = Nothing
            Dim xlRowFrom As Object = Nothing
            Dim xlRowTo As Object = Nothing

            Try
                xlRows = m_activeSheet.Rows
                xlRowFrom = xlRows.Item(startRow)

                If endRow = -1 Then
                    Return m_activeSheet.Range(xlRowFrom, xlRowFrom)
                Else
                    xlRowTo = xlRows.Item(endRow)
                    Return m_activeSheet.Range(xlRowFrom, xlRowTo)
                End If

            Finally
                Me.Free(xlRowTo)
                Me.Free(xlRowFrom)
                Me.Free(xlRows)
            End Try
        End Function

        ''' <summary>
        ''' 指定されたセル範囲を取得します.
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <returns>セル範囲オブジェクト. エラーの場合は Nothing が返ります.</returns>
        ''' <remarks></remarks>
        Protected Function GetCellRC(ByVal axis As XlAxis) As Object
            Return GetCellRC(axis.Row, axis.Column)
        End Function
        ''' <summary>
        ''' 指定されたセル範囲を取得します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス[省略可]</param>
        ''' <param name="startCol">開始列インデックス[省略可]</param>
        ''' <param name="endRow">終了行インデックス[省略可]</param>
        ''' <param name="endCol">終了列インデックス[省略可]</param>
        ''' <returns>セル範囲オブジェクト. エラーの場合は Nothing が返ります.</returns>
        ''' <remarks></remarks>
        Protected Function GetCellRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                                     Optional ByVal endRow As Integer = -1, Optional ByVal endCol As Integer = -1) As Object
            Dim xlCells As Object = Nothing
            Dim xlCellFrom As Object = Nothing
            Dim xlCellTo As Object = Nothing

            Try
                xlCells = m_activeSheet.Cells
                xlCellFrom = xlCells(startRow, startCol)

                If endCol = -1 OrElse endRow = -1 Then
                    Return m_activeSheet.Range(xlCellFrom, xlCellFrom)
                Else
                    xlCellTo = xlCells(endRow, endCol)
                    Return m_activeSheet.Range(xlCellFrom, xlCellTo)
                End If

            Finally
                Me.Free(xlCellTo)
                Me.Free(xlCellFrom)
                Me.Free(xlCells)
            End Try
        End Function

        ''' <summary>
        ''' 指定されたセル範囲を取得します.
        ''' </summary>
        ''' <param name="rangeStr">範囲文字列</param>
        ''' <returns>セル範囲オブジェクト. エラーの場合は Nothing が返ります.</returns>
        ''' <remarks></remarks>
        Protected Function GetCell(ByVal rangeStr As String) As Object
            Dim xlCells As Object = Nothing

            Try
                xlCells = m_activeSheet.Cells
                Return m_activeSheet.Range(rangeStr)

            Finally
                Me.Free(xlCells)
            End Try
        End Function

        ''' <summary>
        ''' 印刷処理(現在のシート)
        ''' </summary>
        ''' <param name="printerName">印刷するプリンタ名</param>
        ''' <param name="isPreview">プレビュー画面を表示する場合、true</param>
        ''' <param name="FromPage"></param>
        ''' <param name="ToPage"></param>
        ''' <param name="copyCount">部数</param>
        ''' <param name="isCollate"></param>
        ''' <remarks></remarks>
        Public Sub Print(Optional ByVal printerName As String = "", Optional ByVal isPreview As Boolean = False, Optional ByVal FromPage As Integer = -1, Optional ByVal ToPage As Integer = -1, Optional ByVal copyCount As Integer = 1, Optional ByVal isCollate As Boolean = True)
            PerformPrint(m_activeSheet, printerName, isPreview, copyCount, isCollate)
        End Sub

        ''' <summary>
        ''' 印刷処理(ワークブック)
        ''' </summary>
        ''' <param name="printerName">印刷するプリンタ名</param>
        ''' <param name="isPreview">プレビュー画面を表示する場合、true</param>
        ''' <param name="FromPage"></param>
        ''' <param name="ToPage"></param>
        ''' <param name="copyCount">部数</param>
        ''' <param name="isCollate"></param>
        ''' <remarks></remarks>
        Public Sub PrintWorkbook(Optional ByVal printerName As String = "", Optional ByVal isPreview As Boolean = False, Optional ByVal FromPage As Integer = -1, Optional ByVal ToPage As Integer = -1, Optional ByVal copyCount As Integer = 1, Optional ByVal isCollate As Boolean = True)
            PerformPrint(m_xlBook, printerName, isPreview, copyCount, isCollate)
        End Sub

        Private Sub PerformPrint(ByVal targetOfPrint As Object, ByVal printerName As String, ByVal isPreview As Boolean, ByVal copyCount As Integer, ByVal isCollate As Boolean)
            Try
                If printerName.Equals(String.Empty) Then
                    targetOfPrint.PrintOut(Preview:=isPreview, Copies:=copyCount, Collate:=isCollate) ', Background:=isBackground)
                Else
                    targetOfPrint.PrintOut(ActivePrinter:=printerName, Preview:=isPreview, Copies:=copyCount, Collate:=isCollate) ', Background:=isBackground)
                End If

            Catch ex As Exception
                '印刷中ダイアログでキャンセルされた場合にエラーとならないよう,
                'エラーをトラップし, 何もしないことで対応.
            End Try
        End Sub

        ''' <summary>
        ''' 列番号から列名への変換
        ''' </summary>
        ''' <param name="col">列番号</param>
        ''' <returns>列名</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToLetter(ByVal col As Integer) As String
            If col <= 0 Then
                Return String.Empty
            End If

            'EXCELの列番号(1始まり)をインデックス番号(0始まり)に変換してから呼び出し
            Return EzUtil.ConvIndexToAlphabet(index:=col - 1)
        End Function

        ''' <summary>
        ''' グラフマーカー処理
        ''' </summary>
        ''' <param name="graphIndex"></param>
        ''' <param name="jikuIndex"></param>
        ''' <remarks></remarks>
        Public Sub ShowGraphMarker(ByVal graphIndex As Integer, ByVal jikuIndex As Integer)
            Dim xlChartObjects As Object = Nothing
            Dim xlChartObject As Object = Nothing
            Dim xlChart As Object = Nothing
            Dim xlSeriesCollection As Object = Nothing
            Dim xlSeries As Object = Nothing

            Try
                xlChartObjects = m_activeSheet.ChartObjects
                xlChartObject = xlChartObjects.Item(1)
                xlChart = xlChartObject.Chart
                xlSeriesCollection = xlChart.SeriesCollection           '軸コレクション
                xlSeries = xlSeriesCollection.Item(2)                   '軸

                'マーカー設定
                xlSeries.MarkerStyle = -4105                            'スタイル自動
                xlSeries.MarkerForegroundColorIndex = 1                 '黒
                xlSeries.MarkerBackgroundColorIndex = 1                 '黒
                xlSeries.MarkerSize = 10                                '大きさ
                xlSeries.Shadow = True                                  '影

            Catch ex As Exception
                Me.Free(xlSeries)
                Me.Free(xlSeriesCollection)
                Me.Free(xlChart)
                Me.Free(xlChartObject)
                Me.Free(xlChartObjects)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' COMオブジェクトの解放
        ''' </summary>
        ''' <param name="obj">解放するオブジェクト</param>
        ''' <remarks></remarks>
        Protected Sub Free(ByRef obj As Object)
            If Not obj Is Nothing Then
                FinalReleaseComObject(obj)
                obj = Nothing
            End If
        End Sub

        ''' <summary>
        ''' ブックが開かれているか返します.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property IsBookOpen() As Boolean
            Get
                Return m_isBookOpen
            End Get
        End Property

        ''' <summary>
        ''' データのある最終行を取得します.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property EndRow() As Integer
            Get
                Dim xlCells As Object = Nothing
                Dim xlSpecialCells As Object = Nothing

                Try
                    xlCells = m_activeSheet.Cells
                    xlSpecialCells = xlCells.SpecialCells(XlCellType.xlCellTypeLastCell)
                    Return xlSpecialCells.Row

                Finally
                    Me.Free(xlSpecialCells)
                    Me.Free(xlCells)
                End Try
            End Get
        End Property

        ''' <summary>
        ''' データのある最終列を取得します.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property EndCol() As Integer
            Get
                Dim xlCells As Object = Nothing
                Dim xlSpecialCells As Object = Nothing

                Try
                    xlCells = m_activeSheet.Cells
                    xlSpecialCells = xlCells.SpecialCells(XlCellType.xlCellTypeLastCell)
                    Return xlSpecialCells.Column

                Finally
                    Me.Free(xlSpecialCells)
                    Me.Free(xlCells)
                End Try
            End Get
        End Property

        ''' <summary>
        ''' データのある最終座標を取得する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAxisOfLastCell() As XlAxis
            Dim xlCells As Object = m_activeSheet.Cells
            Try
                Dim xlSpecialCells As Object = xlCells.SpecialCells(XlCellType.xlCellTypeLastCell)
                Try
                    Return New XlAxis(xlSpecialCells.Row, xlSpecialCells.Column)
                Finally
                    Me.Free(xlSpecialCells)
                End Try
            Finally
                Me.Free(xlCells)
            End Try
        End Function

        ''' <summary>
        ''' EXCELのバージョンを取得します.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Version() As String
            Get
                Return m_xlApp.Version
            End Get
        End Property

        Private disposedValue As Boolean = False        ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    '' TODO: 明示的に呼び出されたときにマネージ リソースを解放します
                End If

                '' TODO: 共有のアンマネージ リソースを解放します



                Dim xlSheets As Object = Nothing
                Dim xlSheet As Object = Nothing

                xlSheets = m_xlBook.Sheets

                'すべてのシートを開放する.
                For i As Integer = 1 To xlSheets.Count
                    xlSheet = xlSheets(i)
                    Me.Free(xlSheet)
                Next

                Me.Free(xlSheets)



                If Not m_xlBook Is Nothing Then
                    'ブックが開かれたままの場合, 保存せず閉じる
                    If m_isBookOpen Then
                        Try
                            m_xlBook.Close()
                        Catch ignore As Exception
                            'nop
                        End Try
                    End If
                End If

                'アクティブシートだけではなく, すべてのシートを上で開放する.
                ''Me.Free(m_activeSheet)

                Me.Free(m_xlBook)

                If Not m_xlApp Is Nothing Then
                    m_xlApp.Quit()                                                  'EXCEL終了
                    m_xlApp.DisplayAlerts = True                                    '警告表示を元に戻す
                End If

                Me.Free(m_xlApp)
            End If

            Me.disposedValue = True
        End Sub

        ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub


        ''' <summary>
        ''' グラフデータラベル設定
        ''' </summary>
        ''' <param name="graphName">グラフ名称(エクセルで設定した値)</param>
        ''' <param name="seriesName">データ系列名称(エクセルで設定した値)</param>
        ''' <param name="value">データラベルに表示する値</param>
        ''' <remarks>
        ''' 注意 既に用意されているデータラベルの文字列しか設定できません.
        ''' </remarks>
        Public Sub SetDataLabelText(ByVal graphName As String, ByVal seriesName As String, ByVal value As String)

            Dim chart As Object = Nothing
            Dim series As Object = Nothing
            'Dim points As Object = Nothing
            'Dim point As Object = Nothing
            Dim trendLines As Object = Nothing
            Dim trendLine As Object = Nothing
            Dim dataLabel As Object = Nothing
            Dim characters As Object = Nothing

            Try
                chart = Me.GetChart(graphName)
                series = Me.GetSeries(chart, seriesName)
                'points = series.points
                'Dim lastPoint As Integer = points.Count
                'point = points.Item(lastPoint)                  '最後のポイントのデータラベルとするため
                'dataLabel = Point.dataLabel
                trendLines = series.TrendLines
                trendLine = trendLines.Item(1)
                dataLabel = trendLine.DataLabel
                characters = dataLabel.characters

                If value.Equals(String.Empty) Then
                    'ブランクの場合,データラベルを削除
                    dataLabel.Delete()
                Else
                    characters.Text = value
                End If

                '設定できない場合にエラーとしたくない場合はコメントを解除してください.
                'Catch ex As Exception

            Finally
                Me.Free(characters)
                Me.Free(dataLabel)
                'Me.Free(point)
                'Me.Free(points)
                Me.Free(trendLine)
                Me.Free(trendLines)
                Me.Free(series)
                Me.Free(chart)
            End Try
        End Sub

        ''' <summary>
        ''' チャート(グラフ)オブジェクトの取得
        ''' </summary>
        ''' <param name="graphName">グラフ名称(エクセルで設定した値)</param>
        ''' <returns>チャート(グラフ)オブジェクト</returns>
        ''' <remarks></remarks>
        Private Function GetChart(ByVal graphName As String) As Object
            Dim chartObjects As Object = Nothing
            Dim chartObject As Object = Nothing

            Try
                chartObjects = m_activeSheet.chartObjects
                chartObject = chartObjects.Item(graphName)
                Return chartObject.chart

            Finally
                Me.Free(chartObject)
                Me.Free(chartObjects)
            End Try
        End Function

        ''' <summary>
        ''' データ系列オブジェクトの取得
        ''' </summary>
        ''' <param name="chart"></param>
        ''' <param name="seriesName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetSeries(ByVal chart As Object, ByVal seriesName As String) As Object
            Dim seriesCollection As Object = Nothing

            Try
                seriesCollection = chart.seriesCollection
                Return seriesCollection.Item(seriesName)

            Finally
                Me.Free(seriesCollection)
            End Try
        End Function

        ''' <summary>
        ''' 指定セルに計算式を設定する
        ''' </summary>
        ''' <param name="row">行インデックス</param>
        ''' <param name="col">列インデックス</param>
        ''' <param name="formula">計算式 ex."=A1*B2/100"</param>
        ''' <remarks></remarks>
        Public Sub SetFormulaRC(ByVal row As Integer, ByVal col As Integer, ByVal formula As String)
            SetFormulaRC(row, col, formula:=formula, isR1C1:=False)
        End Sub

        ''' <summary>
        ''' 指定セルに計算式を設定する
        ''' </summary>
        ''' <param name="row">行インデックス</param>
        ''' <param name="col">列インデックス</param>
        ''' <param name="formula">計算式 ex."=SUM(R[-3]C:R[-1]C)"</param>
        ''' <remarks></remarks>
        Public Sub SetFormulaR1C1RC(ByVal row As Integer, ByVal col As Integer, ByVal formula As String)
            SetFormulaRC(row, col, formula:=formula, isR1C1:=True)
        End Sub

        ''' <summary>
        ''' 指定セルに計算式を設定する
        ''' </summary>
        ''' <param name="row">行インデックス</param>
        ''' <param name="col">列インデックス</param>
        ''' <param name="isR1C1">R1C1形式の計算式かどうか</param>
        ''' <param name="formula">計算式</param>
        ''' <remarks></remarks>
        Private Sub SetFormulaRC(ByVal row As Integer, ByVal col As Integer, ByVal formula As String, ByVal isR1C1 As Boolean)
            Dim xlCell As Object = Nothing
            Try
                xlCell = Me.GetCellRC(row, col)
                If isR1C1 Then
                    xlCell.FormulaR1C1 = formula
                Else
                    xlCell.Formula = formula
                End If
            Finally
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' 合併セールの列数を貰います
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <returns>合併セールの列数</returns>
        ''' <remarks></remarks>
        Public Function GetMergedCellsColumnCountRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                                                    Optional ByVal endRow As Integer = -1, Optional ByVal endCol As Integer = -1) As Integer
            Dim xlRange As Object = Nothing
            Dim xlCell As Object = Nothing
            Dim xlMarea As Object = Nothing
            Dim xlColumns As Object = Nothing

            Try
                xlCell = GetCellRC(startRow, startCol, endRow, endCol)
                xlRange = m_activeSheet.Range(xlCell, xlCell)
                xlMarea = xlRange.MergeArea()
                xlColumns = xlMarea.Columns
                Return xlColumns.Count
            Finally
                Me.Free(xlColumns)
                Me.Free(xlMarea)
                Me.Free(xlRange)
                Me.Free(xlCell)
            End Try
        End Function

        ''' <summary>
        ''' シートのヘッダー設定を行います.
        ''' </summary>
        ''' <param name="name">対象のシート名</param>
        ''' <param name="setString">設定する内容</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetSheetHeader(name As String, setString As String)
            SetSheetHeader(GetSheetIndexByName(name), setString)
        End Sub

        ''' <summary>
        ''' シートのヘッダー設定を行います.
        ''' </summary>
        ''' <param name="index">シートインデックス(1から始まります)</param>
        ''' <param name="setString">設定する内容</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetSheetHeader(ByVal index As Integer, ByVal setString As String)
            Dim xlSheets As Object = Nothing
            Dim xlSheet As Object = Nothing

            Try
                xlSheets = m_xlBook.Sheets
                xlSheet = xlSheets.Item(index)
                xlSheet.PageSetup.RightHeader = setString
            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' シートのフッター設定を行います.
        ''' </summary>
        ''' <param name="name">対象のシート名</param>
        ''' <param name="setString">設定する内容</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetSheetFooter(name As String, setString As String)
            SetSheetFooter(GetSheetIndexByName(name), setString)
        End Sub

        ''' <summary>
        ''' シートのフッター設定を行います.
        ''' </summary>
        ''' <param name="index">シートインデックス(1から始まります)</param>
        ''' <param name="setString">設定する内容</param>
        ''' <remarks></remarks>
        Public Overloads Sub SetSheetFooter(ByVal index As Integer, ByVal setString As String)
            Dim xlSheets As Object = Nothing
            Dim xlSheet As Object = Nothing

            Try
                xlSheets = m_xlBook.Sheets
                xlSheet = xlSheets.Item(index)
                xlSheet.PageSetup.RightFooter = setString
            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' シートを削除
        ''' </summary>
        ''' <param name="sheetIndex">目的シートのIndex</param>
        ''' <remarks></remarks>
        Public Sub DeleteSheet(ByVal sheetIndex As Integer)
            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                Dim xlSheet As Object = xlSheets.Item(sheetIndex)
                Try
                    m_xlBook.Activate()
                    xlSheet.Delete()

                    If xlSheet Is m_activeSheet Then
                        m_activeSheet = m_xlBook.ActiveSheet
                    End If

                Finally
                    ' 削除したから個別に解放する
                    Me.Free(xlSheet)
                End Try
            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' アクティブシートの削除を行います.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub DeleteActiveSheet()
            DeleteSheet(GetActiveSheetIndex)
        End Sub

        ''' <summary>
        ''' シートを末尾にコピー挿入する
        ''' </summary>
        ''' <param name="srcSheetIndex">コピーするシートの位置index</param>
        ''' <param name="copiedSheetName">コピー後シート名</param>
        ''' <remarks></remarks>
        Public Sub CopySheetToLast(ByVal srcSheetIndex As Integer, Optional ByVal copiedSheetName As String = Nothing)
            PerformCopySheet(srcSheetIndex, SheetCount, True, copiedSheetName)
        End Sub

        ''' <summary>
        ''' シートを右隣にコピー挿入する
        ''' </summary>
        ''' <param name="srcSheetIndex">コピーするシートの位置index</param>
        ''' <param name="copiedSheetName">コピー後シート名</param>
        ''' <remarks></remarks>
        Public Sub CopySheet(ByVal srcSheetIndex As Integer, Optional ByVal copiedSheetName As String = Nothing)
            PerformCopySheet(srcSheetIndex, srcSheetIndex, True, copiedSheetName)
        End Sub

        ''' <summary>
        ''' シートをコピー挿入する
        ''' </summary>
        ''' <param name="srcSheetIndex">コピーするシートの位置index</param>
        ''' <param name="destSheetIndex">コピー挿入先のシートの位置index</param>
        ''' <param name="copiedSheetName">コピー後シート名</param>
        ''' <remarks></remarks>
        Public Sub CopySheet(ByVal srcSheetIndex As Integer, ByVal destSheetIndex As Integer, Optional ByVal copiedSheetName As String = Nothing)
            PerformCopySheet(srcSheetIndex, destSheetIndex, False, copiedSheetName)
        End Sub

        ''' <summary>
        ''' シートをコピー挿入する
        ''' </summary>
        ''' <param name="srcSheetIndex">コピーするシートの位置index</param>
        ''' <param name="destSheetIndex">コピー先のシートの位置index</param>
        ''' <param name="copiesAtAfter">コピー先シートの後ろに挿入する場合、true</param>
        ''' <param name="copiedSheetName">コピー後シート名</param>
        ''' <remarks></remarks>
        Private Sub PerformCopySheet(ByVal srcSheetIndex As Integer, ByVal destSheetIndex As Integer, ByVal copiesAtAfter As Boolean, ByVal copiedSheetName As String)

            AssertSheetIndex(srcSheetIndex)
            AssertSheetIndex(destSheetIndex)

            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                Dim xlSrcSheet As Object = xlSheets(srcSheetIndex)
                Dim xlDestSheet As Object = xlSheets(destSheetIndex)
                If copiesAtAfter Then
                    xlSrcSheet.Copy(After:=xlDestSheet)
                Else
                    xlSrcSheet.Copy(xlDestSheet)
                End If
                SetActiveSheet(destSheetIndex + If(copiesAtAfter, 1, 0))

                If Not String.IsNullOrEmpty(copiedSheetName) Then
                    m_activeSheet.Name = copiedSheetName
                End If
            Finally
                Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' シートを右隣にコピー挿入する
        ''' </summary>
        ''' <param name="srcSheetName">コピー元シート名</param>
        ''' <param name="copiedSheetName">コピー後シート名</param>
        ''' <remarks>コピーシートは作られた時点でアクティブシートになる</remarks>
        Public Sub CopySheet(ByVal srcSheetName As String, Optional ByVal copiedSheetName As String = Nothing)

            Dim srcSheetIndex As Integer = GetSheetIndexByName(srcSheetName)
            PerformCopySheet(srcSheetIndex, srcSheetIndex, True, copiedSheetName)
        End Sub

        ''' <summary>
        ''' 別のワークブックのシートをコピーする
        ''' </summary>
        ''' <param name="copyFileName">コピー元ファイル名称</param>
        ''' <param name="copySheetIndex">コピー元シート位置</param>
        ''' <param name="destSheetIndex">コピー先シート位置</param>
        ''' <returns>成否</returns>
        ''' <remarks></remarks>
        Public Function CopySheetOfAnotherBookSupportedExcel2013(ByVal copyFileName As String, ByVal copySheetIndex As Integer, ByVal destSheetIndex As Integer, Optional ByVal isAfterSheet As Boolean = True) As Boolean
            Dim xlCopyBooks As Object = Nothing
            Dim xlCopyBook As Object = Nothing

            Try
                If Not System.IO.File.Exists(copyFileName) Then
                    Throw New ArgumentException("コピー対象として指定されたブックが存在しません")
                End If

                xlCopyBooks = m_xlApp.Workbooks
                xlCopyBook = xlCopyBooks.Open(copyFileName, True)
                If isAfterSheet Then
                    xlCopyBook.Sheets(copySheetIndex).Copy(after:=m_xlBook.Sheets(destSheetIndex))
                Else
                    xlCopyBook.Sheets(copySheetIndex).Copy(before:=m_xlBook.Sheets(destSheetIndex))
                End If

                Return True
            Finally
                If copyFileName.Equals(m_xlApp.ActiveWorkBook.FullName) Then
                    xlCopyBook.Close(False)
                End If
                Me.Free(xlCopyBook)
                Me.Free(xlCopyBooks)
            End Try
        End Function

        ''' <summary>
        ''' 別のワークブックのシートを全てコピーする
        ''' </summary>
        ''' <param name="copyFileName">コピー元ファイル名称</param>
        ''' <param name="destSheetIndex">コピー先シート位置</param>
        ''' <returns>成否</returns>
        ''' <remarks></remarks>
        Public Function CopySheetOfAnotherBookForAllSupportedExcel2013(ByVal copyFileName As String, ByVal destSheetIndex As Integer, Optional ByVal isAfterSheet As Boolean = True) As Boolean
            If Not System.IO.File.Exists(copyFileName) Then
                Throw New ArgumentException("コピー対象として指定されたブックが存在しません")
            End If
            If Not (System.IO.Path.GetExtension(copyFileName).ToUpper).Equals(System.IO.Path.GetExtension(m_fileName).ToUpper) Then
                Throw New ArgumentException("コピー先のブックとコピー元のブックで拡張子が違います")
            End If

            Dim xlCopyBooks As Object = Nothing
            Dim xlCopyBook As Object = Nothing
            Dim xlCopyBookWorksheets As Object = Nothing
            Try

                xlCopyBooks = m_xlApp.Workbooks
                xlCopyBook = xlCopyBooks.Open(copyFileName, True)
                xlCopyBookWorksheets = xlCopyBook.Worksheets
                If isAfterSheet Then
                    xlCopyBookWorksheets.Copy(after:=m_xlBook.Sheets(destSheetIndex))
                Else
                    xlCopyBookWorksheets.Copy(before:=m_xlBook.Sheets(destSheetIndex))
                End If

                Return True
            Finally
                Me.Free(xlCopyBookWorksheets)
                If copyFileName.Equals(m_xlApp.ActiveWorkBook.FullName) Then
                    xlCopyBook.Close(False)
                End If
                Me.Free(xlCopyBook)
                Me.Free(xlCopyBooks)
            End Try
        End Function

        ''' <summary>
        ''' シート名のシート位置indexを返す
        ''' </summary>
        ''' <param name="sheetName">シート名</param>
        ''' <returns>シート位置index</returns>
        ''' <remarks></remarks>
        Public Function GetSheetIndexByName(ByVal sheetName As String) As Integer

            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                For i As Integer = 1 To xlSheets.Count
                    If sheetName.Equals(xlSheets(i).Name) Then
                        Return i
                    End If
                Next
                Return -1
            Finally
                Me.Free(xlSheets)
            End Try
        End Function

        ''' <summary>
        ''' シート名一覧からシートインデックス一覧を返す
        ''' </summary>
        ''' <param name="sheetNames">シート名[]</param>
        ''' <returns>シート位置index</returns>
        ''' <remarks></remarks>
        Public Function GetSheetIndexesByNames(ByVal sheetNames As String()) As Integer()
            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                Dim result As New List(Of Integer)
                For sheetIndex As Integer = 1 To xlSheets.Count
                    Dim xlSheet As Object = xlSheets.Item(sheetIndex)
                    Try
                        If sheetNames.Contains(xlSheet.Name) AndAlso Not result.Contains(sheetIndex) Then
                            result.Add(sheetIndex)
                        End If
                    Finally
                        Me.Free(xlSheet)
                    End Try
                Next
                Return result.ToArray
            Finally
                Me.Free(xlSheets)
            End Try
        End Function

        ''' <summary>
        ''' シートインデックスからシート名を返す
        ''' </summary>
        ''' <param name="sheetIndex">シートIndex</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSheetNameBySheetIndex(ByVal sheetIndex As Integer) As String
            Dim xlSheets As Object = Nothing
            Dim xlSheet As Object = Nothing
            Try
                xlSheets = m_xlBook.Sheets
                xlSheet = xlSheets.Item(sheetIndex)

                Return xlSheet.Name
            Finally
                Me.Free(xlSheets)
            End Try
        End Function

        ''' <summary>
        ''' シートインデックス一覧からシート名一覧を返す
        ''' </summary>
        ''' <param name="sheetIndexes">シートIndex[]</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSheetNamesBySheetIndexes(ByVal sheetIndexes As Integer()) As String()
            Dim xlSheets As Object = m_xlBook.Sheets

            Try
                Dim result As New List(Of String)
                For Each sheetIndex As Integer In sheetIndexes
                    Dim xlSheet As Object = xlSheets.Item(sheetIndex)
                    Try
                        If result.Contains(xlSheet.name) Then
                            Continue For
                        End If
                        result.Add(xlSheet.Name)
                    Finally
                        Me.Free(xlSheet)
                    End Try
                Next

                Return result.ToArray
            Finally
                Me.Free(xlSheets)
            End Try
        End Function

        ''' <summary>
        ''' ExcelのWorkBookをクリア
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ClearWorkBook()

            Dim xlSheets As Object = Nothing
            Dim sheetCount As Integer

            Try
                xlSheets = m_xlBook.Sheets
                sheetCount = xlSheets.Count()
                AddSheet()

                While xlSheets.Count() > 1
                    DeleteSheet(1)
                End While

                AddSheet()
                AddSheet()
                SetActiveSheet(1)
                SetSheetName("Sheet1")
                SetActiveSheet(2)
                SetSheetName("Sheet2")
                SetActiveSheet(3)
                SetSheetName("Sheet3")
                SetActiveSheet(1)

            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' シート保護を解除
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub UnProtectAtActiveSheet()
            UnProtect(GetActiveSheetIndex)
        End Sub

        ''' <summary>
        ''' シート保護を解除
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub UnProtect(ByVal sheetIndex As Integer)
            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                xlSheets(sheetIndex).UnProtect()
            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' 列の非表示を設定する
        ''' </summary>
        ''' <param name="startColumn">開始列index</param>
        ''' <param name="hidden">非表示値</param>
        ''' <remarks></remarks>
        Public Sub SetColumnHidden(ByVal startColumn As Integer, ByVal hidden As Boolean)
            SetColumnHidden(startColumn, -1, hidden)
        End Sub

        ''' <summary>
        ''' 列の非表示を設定する
        ''' </summary>
        ''' <param name="startColumn">開始列index</param>
        ''' <param name="endColumn">終了列index</param>
        ''' <param name="hidden">非表示値</param>
        ''' <remarks></remarks>
        Public Sub SetColumnHidden(ByVal startColumn As Integer, ByVal endColumn As Integer, ByVal hidden As Boolean)
            Dim xColumn As Object = Me.GetColumn(startColumn, endColumn)
            Try
                Dim xEntire As Object = xColumn.EntireColumn
                Try
                    xEntire.Hidden = hidden
                Finally
                    Me.Free(xEntire)
                End Try
            Finally
                Me.Free(xColumn)
            End Try
        End Sub

        ''' <summary>
        ''' 行の非表示を設定する
        ''' </summary>
        ''' <param name="stRow">開始行index</param>
        ''' <param name="hidden">非表示値</param>
        ''' <remarks></remarks>
        Public Sub SetRowHidden(ByVal stRow As Integer, ByVal hidden As Boolean)
            SetRowHidden(stRow, -1, hidden)
        End Sub

        ''' <summary>
        ''' 行の非表示を設定する
        ''' </summary>
        ''' <param name="stRow">開始行index</param>
        ''' <param name="endRow">終了行index</param>
        ''' <param name="hidden">非表示値</param>
        ''' <remarks></remarks>
        Public Sub SetRowHidden(ByVal stRow As Integer, ByVal endRow As Integer, ByVal hidden As Boolean)
            Dim xCells As Object = Nothing
            Dim XEntire As Object = Nothing

            Try
                If endRow = -1 Then
                    xCells = Me.GetCellRC(stRow, 1, stRow, 1)
                Else
                    xCells = Me.GetCellRC(stRow, 1, endRow, 1)
                End If

                XEntire = xCells.EntireRow

                If hidden = True Then
                    XEntire.Hidden = True
                Else
                    XEntire.Hidden = False
                End If

            Finally
                Me.Free(XEntire)
                Me.Free(xCells)
            End Try
        End Sub

        ''' <summary>
        ''' 単一アクティブセルを選択
        ''' </summary>
        ''' <param name="aSheetIndex">対象シートNo</param>
        ''' <param name="aCellAdr">単一セルアドレス(A1形式)</param>
        ''' <remarks></remarks>
        Public Sub SetActiveCell(ByVal aSheetIndex As Integer, ByVal aCellAdr As String)
            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                Dim xlRange As Object = xlSheets(aSheetIndex).Range(aCellAdr)
                Try
                    xlRange.Select()

                Finally
                    Me.Free(xlRange)
                End Try
            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' 印刷範囲を設定する
        ''' </summary>
        ''' <param name="startCol">開始列(1start)</param>
        ''' <param name="startRow">開始行(1start)</param>
        ''' <param name="endCol">終了列(1start)</param>
        ''' <param name="endRow">終了行(1start)</param>
        ''' <remarks></remarks>
        Public Sub SetPrintRange(ByVal startCol As Integer, ByVal startRow As Integer, _
                                 ByVal endCol As Integer, ByVal endRow As Integer)

            AssertCellRange(startCol, startRow, endCol, endRow)

            Dim xlPageSetup As Object = m_activeSheet.PageSetup
            Try
                Dim rangeString As String = MakeRangeFormat(startCol, startRow, endCol, endRow)

                xlPageSetup.printArea = rangeString

            Finally
                Me.Free(xlPageSetup)
            End Try
        End Sub

        ''' <summary>
        ''' 印刷用拡大率を変更する
        ''' </summary>
        ''' <param name="ZoomValue">拡大率(10%～400%まで)</param>
        ''' <remarks></remarks>
        Public Sub SetPrintZoom(ByVal ZoomValue As Integer)
            Dim xlPageSetup As Object = m_activeSheet.PageSetup
            Try
                xlPageSetup.Zoom = ZoomValue

            Finally
                Me.Free(xlPageSetup)
            End Try
        End Sub


        ''' <summary>
        ''' 印刷の向きを設定する
        ''' </summary>
        ''' <param name="pageOrientation">印刷の向き</param>
        ''' <remarks></remarks>
        Public Sub SetPrintOrientation(ByVal pageOrientation As XlPageOrientation)
            Dim xlPageSetup As Object = m_activeSheet.PageSetup
            Try
                xlPageSetup.Orientation = pageOrientation

            Finally
                Me.Free(xlPageSetup)
            End Try
        End Sub

        ''' <summary>
        ''' 印刷用紙の設定をする(まだA3とA4のみ)
        ''' </summary>
        ''' <param name="paperSize">用紙サイズ</param>
        ''' <remarks></remarks>
        Public Sub SetPrintPaper(ByVal paperSize As XlPaperSize)
            Dim xlPageSetup As Object = m_activeSheet.PageSetup
            Try
                xlPageSetup.PaperSize = paperSize

            Finally
                Me.Free(xlPageSetup)
            End Try
        End Sub


        ''' <summary>シートの数</summary>
        Public ReadOnly Property SheetCount() As Integer
            Get
                Dim xlSheets As Object = m_xlBook.Sheets
                Try
                    Return xlSheets.Count()
                Finally
                    Me.Free(xlSheets)
                End Try
            End Get
        End Property

        ''' <summary>
        ''' アクティブシートのシート名を返す
        ''' </summary>
        ''' <returns>シート名</returns>
        ''' <remarks></remarks>
        Public Function GetActiveSheetName() As String
            Return m_activeSheet.Name
        End Function

        ''' <summary>
        ''' アクティブシートのシートindexを返す
        ''' </summary>
        ''' <returns>シートindex</returns>
        ''' <remarks></remarks>
        Public Function GetActiveSheetIndex() As Integer
            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                For i As Integer = 1 To xlSheets.Count
                    If m_activeSheet Is xlSheets(i) Then
                        Return i
                    End If
                Next
            Finally
                Free(xlSheets)
            End Try
            Return -1
        End Function

        ''' <summary>
        ''' #GetValueした値を二次元配列にして返す
        ''' </summary>
        ''' <param name="gotValue">#GetValueした値</param>
        ''' <returns>二次元配列</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvGotValueToStrings(ByVal gotValue As Object) As String()()
            Return CollectionUtil.ConvToTwoJaggedArrayAsString(gotValue)
        End Function

        ''' <summary>
        ''' シートは表示中かを返す
        ''' </summary>
        ''' <param name="index">対象のシートindex</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function IsVisibleSheet(ByVal index As Integer) As Boolean
            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                Return xlSheets.Item(index).Visible

            Finally
                Me.Free(xlSheets)
            End Try
        End Function

        ''' <summary>
        ''' シートは表示中かを返す
        ''' </summary>
        ''' <param name="name">対象のシート名</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function IsVisibleSheet(ByVal name As String) As Boolean
            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                Return xlSheets.Item(name).Visible

            Finally
                Me.Free(xlSheets)
            End Try
        End Function

        ''' <summary>
        ''' 範囲が有効であることを保証する
        ''' </summary>
        ''' <param name="startCol">開始列(1start)</param>
        ''' <param name="startRow">開始行(1start)</param>
        ''' <param name="endCol">終了列(1start)</param>
        ''' <param name="endRow">終了行(1start)</param>
        ''' <remarks></remarks>
        Public Sub AssertCellRange(ByVal startCol As Integer, ByVal startRow As Integer, _
                                        ByVal endCol As Integer, ByVal endRow As Integer)
            If startCol < 1 Then
                Throw New ArgumentOutOfRangeException("startCol", startCol, "1以上を指定すべき")
            End If
            If startRow < 1 Then
                Throw New ArgumentOutOfRangeException("startRow", startRow, "1以上を指定すべき")
            End If
            If endCol < 1 Then
                Throw New ArgumentOutOfRangeException("endCol", endCol, "1以上を指定すべき")
            End If
            If endRow < 1 Then
                Throw New ArgumentOutOfRangeException("endRow", endRow, "1以上を指定すべき")
            End If
        End Sub

        ''' <summary>
        ''' シートindexが有効であることを保証する
        ''' </summary>
        ''' <param name="sheetIndex">シートindex</param>
        ''' <remarks></remarks>
        Public Sub AssertSheetIndex(ByVal sheetIndex As Integer)
            If sheetIndex < 1 Then
                Throw New ArgumentOutOfRangeException(String.Format("シートindex {0} は無効", sheetIndex))
            End If
            If SheetCount < sheetIndex Then
                Throw New ArgumentOutOfRangeException(String.Format("シート数は {0} だが {1} を指定している", SheetCount, sheetIndex))
            End If
        End Sub

        ''' <summary>
        ''' オートシェイプライン追加
        ''' </summary>
        ''' <param name="startRow">ライン開始となるセル行</param>
        ''' <param name="startCol">ライン開始となるセル列</param>
        ''' <param name="endRow">ライン終端となるセル行</param>
        ''' <param name="endCol">ライン終端となるセル列</param>
        ''' <remarks>ラインは指定した開始・終端セルの左上座標を結ぶ</remarks>
        Public Sub AddShapeLineRC(ByVal startRow As Integer, ByVal startCol As Integer, _
                                  ByVal endRow As Integer, ByVal endCol As Integer)
            Dim xlRangeStart As Object = GetCellRC(startRow, startCol, startRow, startCol)
            Try
                Dim xlRangeEnd As Object = GetCellRC(endRow, endCol, endRow, endCol)
                Try
                    Dim xStart As Integer = xlRangeStart.Left
                    Dim yStart As Integer = xlRangeStart.Top
                    Dim xEnd As Integer = xlRangeEnd.Left
                    Dim yEnd As Integer = xlRangeEnd.Top
                    AddShapeLineBy(xStart, yStart, xEnd, yEnd)
                Finally
                    Free(xlRangeEnd)
                End Try
            Finally
                Free(xlRangeStart)
            End Try
        End Sub

        ''' <summary>
        ''' オートシェイプでラインを追加します.
        ''' </summary>
        ''' <param name="startPos">開始座標</param>
        ''' <param name="endPos">終了座標</param>
        Public Sub AddShapeLine(ByVal startPos As Point, ByVal endPos As Point)
            AddShapeLineBy(startPos.X, startPos.Y, endPos.X, endPos.Y)
        End Sub

        ''' <summary>
        ''' オートシェイプでラインを追加します.
        ''' </summary>
        ''' <param name="startX">開始座標</param>
        ''' <param name="startY">開始座標</param>
        ''' <param name="endX">終了座標</param>
        ''' <param name="endY">終了座標</param>
        Public Sub AddShapeLineBy(ByVal startX As Integer, ByVal startY As Integer, ByVal endX As Integer, ByVal endY As Integer)
            Dim shapes As Object = m_activeSheet.Shapes
            Try
                shapes.AddLine(startX, startY, endX, endY)

            Finally
                Free(shapes)
            End Try
        End Sub

        ''' <summary>
        ''' 楕円シェイプを追加します.
        ''' </summary>
        ''' <param name="left">開始座標</param>
        ''' <param name="top">開始座標</param>
        ''' <param name="width">幅</param>
        ''' <param name="height">高さ</param>
        Public Sub AddShapeOvalBy(ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
            AddShape(XlAutoShapeType.msoShapeOval, left, top, width, height)
        End Sub

        ''' <summary>
        ''' オートシェイプを追加します.
        ''' </summary>
        ''' <param name="shapeType">シェイプの型</param>
        ''' <param name="left">開始座標</param>
        ''' <param name="top">開始座標</param>
        ''' <param name="width">幅</param>
        ''' <param name="height">高さ</param>
        Public Sub AddShape(ByVal shapeType As XlAutoShapeType, ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
            Dim shapes As Object = m_activeSheet.Shapes
            Try
                shapes.AddShape(shapeType, left, top, width, height)

            Finally
                Free(shapes)
            End Try
        End Sub

        ''' <summary>
        ''' 指定された範囲のすべてのデータを削除します.
        ''' </summary>
        ''' <param name="startRow">対象セルの開始行インデックス</param>
        ''' <param name="startCol">対象セルの開始列インデックス</param>
        ''' <param name="endRow">対象セルの終了行インデックス</param>
        ''' <param name="endCol">対象セルの終了列インデックス</param>
        ''' <remarks></remarks>
        Public Sub ClearContentsRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer)
            Dim xlRange As Object = Me.GetCellRC(startRow, startCol, endRow, endCol)
            Try
                xlRange.ClearContents()
            Finally
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' セルの書式設定の内の「保護 | ロック」を設定します.
        ''' </summary>
        ''' <param name="row">対象セルの行インデックス</param>
        ''' <param name="col">対象セルの列インデックス</param>
        ''' <param name="isLocked">設定するロック状態</param>
        ''' <remarks></remarks>
        Public Sub SetLockedRC(ByVal row As Integer, ByVal col As Integer, ByVal isLocked As Boolean)
            Dim xlCell As Object = m_activeSheet.Cells(row, col)
            Try
                If xlCell.MergeCells Then
                    Dim xlMerge As Object = xlCell.MergeArea
                    Try
                        xlMerge.Locked = isLocked
                    Finally
                        Me.Free(xlMerge)
                    End Try
                Else
                    xlCell.Locked = isLocked
                End If
            Finally
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' セルの書式設定の内の「保護 | ロック」を一括設定します.
        ''' </summary>
        ''' <param name="startRow">対象セルの開始行インデックス</param>
        ''' <param name="startCol">対象セルの開始列インデックス</param>
        ''' <param name="endRow">対象セルの終了行インデックス</param>
        ''' <param name="endCol">対象セルの終了列インデックス</param>
        ''' <param name="isLocked">設定するロック状態</param>
        ''' <remarks></remarks>
        Public Sub SetLockedRangeRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer, ByVal isLocked As Boolean)
            Dim xlRange As Object = Me.GetCellRC(startRow, startCol, endRow, endCol)
            Try
                xlRange.Locked = isLocked
            Catch ex As Exception
                If "Range クラスの Locked プロパティを設定できません。".Equals(ex.Message) Then
                    ' 上記で設定できれば、それが速いが、マージセルが含まれるとエラーになるので以下で救済
                    For row As Integer = startRow To endRow
                        For col As Integer = startCol To endCol
                            SetLockedRC(row, col, isLocked)
                        Next
                    Next
                    Return
                End If
                Throw
            Finally
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' セルの左上座標を取得します.
        ''' </summary>
        ''' <param name="row">対象セルの行インデックス</param>
        ''' <param name="col">対象セルの列インデックス</param>
        ''' <remarks></remarks>
        Public Function GetCellTopLeftRC(ByVal row As Integer, ByVal col As Integer) As Point
            Dim pos As Point
            Dim xlRange As Object = Me.GetCellRC(row, col, row, col)
            Try
                pos.X = xlRange.Left
                pos.Y = xlRange.Top
            Finally
                Me.Free(xlRange)
            End Try
            Return pos
        End Function

        ''' <summary>
        ''' セルの右下座標を取得します.
        ''' </summary>
        ''' <param name="row">対象セルの行インデックス</param>
        ''' <param name="col">対象セルの列インデックス</param>
        ''' <remarks></remarks>
        Public Function GetCellBottomRightRC(ByVal row As Integer, ByVal col As Integer) As Point
            Dim pos As Point
            Dim xlRange As Object = Me.GetCellRC(row, col, row, col)
            Try
                pos.X = xlRange.Left + xlRange.Width
                pos.Y = xlRange.Top + +xlRange.Height
            Finally
                Me.Free(xlRange)
            End Try
            Return pos
        End Function

        ''' <summary>
        ''' セルの真ん中座標を取得します.
        ''' </summary>
        ''' <param name="row">対象セルの行インデックス</param>
        ''' <param name="col">対象セルの列インデックス</param>
        ''' <remarks></remarks>
        Public Function GetCellCenterRC(ByVal row As Integer, ByVal col As Integer) As Point
            Dim pos As Point
            Dim xlRange As Object = Me.GetCellRC(row, col, row, col)
            Try
                pos.X = xlRange.Left + CInt(xlRange.Width / 2)
                pos.Y = xlRange.Top + CInt(xlRange.Height / 2)
            Finally
                Me.Free(xlRange)
            End Try
            Return pos
        End Function

        ''' <summary>
        ''' セルの大きさを取得します.
        ''' </summary>
        ''' <param name="row">対象セルの行インデックス</param>
        ''' <param name="col">対象セルの列インデックス</param>
        ''' <remarks></remarks>
        Public Function GetCellSizeRC(ByVal row As Integer, ByVal col As Integer) As Size
            Dim sz As Size
            Dim xlRange As Object = Me.GetCellRC(row, col, row, col)
            Try
                sz.Width = xlRange.width
                sz.Height = xlRange.height
            Finally
                Me.Free(xlRange)
            End Try
            Return sz
        End Function

        ''' <summary>
        ''' ページ設定のヘッダーに文字列を設定します.
        ''' </summary>
        ''' <param name="left">左上</param>
        ''' <param name="center">中央上</param>
        ''' <param name="right">右上</param>
        ''' <remarks></remarks>
        Public Sub SetPageSetupHeader(ByVal left As String, ByVal center As String, ByVal right As String)
            m_activeSheet = m_xlBook.ActiveSheet
            Dim xlPageSetup As Object = m_activeSheet.PageSetup
            Try
                xlPageSetup.LeftHeader = left
                xlPageSetup.CenterHeader = center
                xlPageSetup.RightHeader = right
            Finally
                Me.Free(xlPageSetup)
            End Try

        End Sub

        ''' <summary>
        ''' 「次のページ数に合わせて印刷」を設定する
        ''' </summary>
        ''' <param name="wide">横ページ数</param>
        ''' <param name="tail">縦ページ数</param>
        ''' <remarks></remarks>
        Public Sub SetPageSetupFitToPages(ByVal wide As Integer?, ByVal tail As Integer?)
            m_activeSheet = m_xlBook.ActiveSheet
            Dim xlPageSetup As Object = m_activeSheet.PageSetup
            Try
                If wide.HasValue OrElse tail.HasValue Then
                    xlPageSetup.Zoom = False
                End If
                If wide.HasValue Then
                    xlPageSetup.FitToPagesWide = wide.Value
                Else
                    xlPageSetup.FitToPagesWide = False
                End If
                If tail.HasValue Then
                    xlPageSetup.FitToPagesTall = tail.Value
                Else
                    xlPageSetup.FitToPagesTall = False
                End If
            Finally
                Me.Free(xlPageSetup)
            End Try

        End Sub

        Private Shared Function AttachBordersIndexAsDictionary(ByVal aBorders As XlBorders) As Dictionary(Of XlBordersIndex, XlBorders.Border)
            Dim resultByIndex As New Dictionary(Of XlBordersIndex, XlBorders.Border)
            resultByIndex.Add(XlBordersIndex.xlDiagonalDown, aBorders.DiagonalDown)
            resultByIndex.Add(XlBordersIndex.xlDiagonalUp, aBorders.DiagonalUp)
            resultByIndex.Add(XlBordersIndex.xlEdgeLeft, aBorders.EdgeLeft)
            resultByIndex.Add(XlBordersIndex.xlEdgeTop, aBorders.EdgeTop)
            resultByIndex.Add(XlBordersIndex.xlEdgeBottom, aBorders.EdgeBottom)
            resultByIndex.Add(XlBordersIndex.xlEdgeRight, aBorders.EdgeRight)
            resultByIndex.Add(XlBordersIndex.xlInsideVertical, aBorders.InsideVertical)
            resultByIndex.Add(XlBordersIndex.xlInsideHorizontal, aBorders.InsideHorizontal)
            Return resultByIndex
        End Function

        ''' <summary>
        ''' 指定範囲の罫線情報群を取得します
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <remarks></remarks>
        Public Function GetBordersRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer) As XlBorders

            Dim xlRange As Object = Me.GetCellRC(startRow, startCol, endRow, endCol)
            Try
                Dim xlRangeBorders As Object = xlRange.Borders
                Try
                    Dim setValues As Action(Of Object, XlBorders.Border) =
                        Sub(srcXlBorder As Object, destBorder As XlBorders.Border)
                            If IsNumeric(srcXlBorder.LineStyle) Then
                                destBorder.LineStyle = [Enum].Parse(GetType(XlLineStyle), srcXlBorder.LineStyle)
                            End If
                            If IsNumeric(srcXlBorder.Weight) Then
                                destBorder.Weight = [Enum].Parse(GetType(XlBorderWeight), srcXlBorder.Weight)
                            End If
                        End Sub
                    Dim result As New XlBorders
                    setValues.Invoke(xlRangeBorders, result)
                    Dim propertyByIndex As Dictionary(Of XlBordersIndex, XlBorders.Border) = AttachBordersIndexAsDictionary(result)
                    For Each pair As KeyValuePair(Of XlBordersIndex, XlBorders.Border) In propertyByIndex
                        Dim xlBorder As Object = xlRangeBorders.Item(pair.Key)
                        Try
                            setValues.Invoke(xlBorder, pair.Value)
                        Finally
                            Me.Free(xlBorder)
                        End Try
                    Next
                    Return result
                Finally
                    Me.Free(xlRangeBorders)
                End Try
            Finally
                Me.Free(xlRange)
            End Try
        End Function

        ''' <summary>
        ''' 指定範囲に罫線を設定します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="aBorders">罫線情報(群)</param>
        ''' <remarks></remarks>
        Public Sub SetBordersRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer,
                                ByVal aBorders As XlBorders.Border)
            Dim xlRange As Object = Me.GetCellRC(startRow, startCol, endRow, endCol)
            Try
                Dim xlRangeBorders As Object = xlRange.Borders
                Try
                    Dim setValues As Action(Of XlBorders.Border, Object) =
                        Sub(srcBorder As XlBorders.Border, destXlBorder As Object)
                            If srcBorder.LineStyle.HasValue Then
                                destXlBorder.LineStyle = srcBorder.LineStyle.Value
                            End If
                            If srcBorder.Weight.HasValue Then
                                If Not XlLineStyle.xlLineStyleNone.Equals(srcBorder.LineStyle.Value) Then
                                    destXlBorder.Weight = srcBorder.Weight.Value
                                End If
                            End If
                        End Sub
                    setValues.Invoke(aBorders, xlRangeBorders)
                    If Not TypeOf aBorders Is XlBorders Then
                        Return
                    End If
                    Dim propertyByIndex As Dictionary(Of XlBordersIndex, XlBorders.Border) = AttachBordersIndexAsDictionary(aBorders)
                    For Each pair As KeyValuePair(Of XlBordersIndex, XlBorders.Border) In propertyByIndex
                        Dim xlBorder As Object = xlRangeBorders.Item(pair.Key)
                        Try
                            setValues.Invoke(pair.Value, xlBorder)
                        Finally
                            Me.Free(xlBorder)
                        End Try
                    Next
                Finally
                    Me.Free(xlRangeBorders)
                End Try
            Finally
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 指定範囲の指定場所に罫線を設定します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="index">罫線の場所</param>
        ''' <param name="lineStyle">罫線の種類[省略可]</param>
        ''' <param name="weight">罫線の太さ[省略可]</param>
        ''' <remarks></remarks>
        Public Sub SetLineRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer, _
                             ByVal index As XlBordersIndex, _
                             Optional ByVal lineStyle As XlLineStyle = XlLineStyle.xlContinuous, _
                             Optional ByVal weight As XlBorderWeight = XlBorderWeight.xlThin)

            Dim xlRange As Object = Me.GetCellRC(startRow, startCol, endRow, endCol)
            Try
                Dim xlBorders As Object = xlRange.Borders
                Try
                    Dim xlBorder As Object = xlBorders.Item(index)
                    Try
                        xlBorder.LineStyle = lineStyle
                        If Not XlLineStyle.xlLineStyleNone.Equals(lineStyle) Then
                            xlBorder.Weight = weight
                        End If
                    Finally
                        Me.Free(xlBorder)
                    End Try
                Finally
                    Me.Free(xlBorders)
                End Try
            Finally
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 指定範囲のすべての場所に罫線を設定します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="lineStyle">罫線の種類[省略可]</param>
        ''' <param name="weight">罫線の太さ[省略可]</param>
        ''' <remarks>スーパークラスの線を消せないバグ修正</remarks>
        Public Sub SetLineRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer, _
                             Optional ByVal lineStyle As XlLineStyle = XlLineStyle.xlContinuous, _
                             Optional ByVal weight As XlBorderWeight = XlBorderWeight.xlThin)

            Dim xlRange As Object = Me.GetCellRC(startRow, startCol, endRow, endCol)
            Try
                Dim xlBorders As Object = xlRange.Borders
                Try
                    xlBorders.LineStyle = lineStyle
                    If Not XlLineStyle.xlLineStyleNone.Equals(lineStyle) Then
                        xlBorders.Weight = weight
                    End If
                Finally
                    Me.Free(xlBorders)
                End Try
            Finally
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 指定範囲の罫線を解除します.
        ''' </summary>
        ''' <param name="startRow">開始行インデックス</param>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endRow">終了行インデックス</param>
        ''' <param name="endCol">終了列インデックス</param>
        ''' <param name="lineStyle">罫線の種類[省略可]</param>
        ''' <param name="weight">罫線の太さ[省略可]</param>
        ''' <remarks></remarks>
        Public Sub SetClearLineRC(ByVal startRow As Integer, ByVal startCol As Integer, ByVal endRow As Integer, ByVal endCol As Integer, _
                                  Optional ByVal lineStyle As XlLineStyle = XlLineStyle.xlLineStyleNone, _
                                  Optional ByVal weight As XlBorderWeight = XlBorderWeight.xlThin)
            Dim xlRange As Object = Nothing
            Dim xlBorders As Object = Nothing
            Dim xlBorder As Object = Nothing

            Try
                xlRange = Me.GetCellRC(startRow, startCol, endRow, endCol)
                xlBorders = xlRange.Borders(XlBordersIndex.xlInsideHorizontal)
                xlBorders.LineStyle = lineStyle
            Finally
                Me.Free(xlBorder)
                Me.Free(xlBorders)
                Me.Free(xlRange)

            End Try
        End Sub

        ''' <summary>
        ''' 対象シートを保護し,ユーザーの変更を禁止します.
        ''' </summary>
        ''' <param name="sheetName">設定を行うシート名</param>
        ''' <param name="password">保護解除の際に必要なパスワード</param>
        ''' <param name="drawingObjects">描画オブジェクトを保護</param>
        ''' <param name="contents">シートの内容を保護</param>
        ''' <param name="senarios">シナリオを保護</param>
        ''' <param name="userInterfaceOnly">マクロからの変更は可能にする</param>
        ''' <param name="allowFormattingCells">セル書式変更は許可する ※Excel2003以降</param>
        ''' <param name="allowFormattingRows">行高さ等の書式変更は許可する ※Excel2003以降</param>
        ''' <param name="allowFormattingColumns">列幅等の書式変更は許可する ※Excel2003以降</param>
        ''' <remarks></remarks>
        Public Sub Protect(ByVal sheetName As String, _
                           Optional ByVal password As String = "", _
                           Optional ByVal drawingObjects As Boolean = True, _
                           Optional ByVal contents As Boolean = True, _
                           Optional ByVal senarios As Boolean = True, _
                           Optional ByVal userInterfaceOnly As Boolean = False, _
                           Optional ByVal allowFormattingCells As Boolean = False,
                           Optional ByVal allowFormattingRows As Boolean = False,
                           Optional ByVal allowFormattingColumns As Boolean = False)
            Protect(GetSheetIndexByName(sheetName), password:=password, drawingObjects:=drawingObjects, _
                    contents:=contents, senarios:=senarios, userInterfaceOnly:=userInterfaceOnly, _
                    allowFormattingCells:=allowFormattingCells, allowFormattingRows:=allowFormattingRows,
                    allowFormattingColumns:=allowFormattingColumns)
        End Sub

        ''' <summary>
        ''' 対象シートを保護し,ユーザーの変更を禁止します.
        ''' </summary>
        ''' <param name="sheetIndex">シート位置index</param>
        ''' <param name="password">保護解除の際に必要なパスワード</param>
        ''' <param name="drawingObjects">描画オブジェクトを保護</param>
        ''' <param name="contents">シートの内容を保護</param>
        ''' <param name="senarios">シナリオを保護</param>
        ''' <param name="userInterfaceOnly">マクロからの変更は可能にする</param>
        ''' <param name="allowFormattingCells">セル書式変更は許可する ※Excel2003以降</param>
        ''' <param name="allowFormattingRows">行高さ等の書式変更は許可する ※Excel2003以降</param>
        ''' <param name="allowFormattingColumns">列幅等の書式変更は許可する ※Excel2003以降</param>
        ''' <remarks></remarks>
        Public Sub Protect(ByVal sheetIndex As Integer, _
                           Optional ByVal password As String = "", _
                           Optional ByVal drawingObjects As Boolean = True, _
                           Optional ByVal contents As Boolean = True, _
                           Optional ByVal senarios As Boolean = True, _
                           Optional ByVal userInterfaceOnly As Boolean = False, _
                           Optional ByVal allowFormattingCells As Boolean = False, _
                           Optional ByVal allowFormattingRows As Boolean = False, _
                           Optional ByVal allowFormattingColumns As Boolean = False)

            Dim xlSheets As Object = m_xlBook.Sheets

            Try
                Try
                    ' Excel2003以降
                    xlSheets.Item(sheetIndex).Protect(
                        password, drawingObjects, contents, senarios, userInterfaceOnly,
                        allowFormattingCells:=allowFormattingCells, allowFormattingRows:=allowFormattingRows,
                        allowFormattingColumns:=allowFormattingColumns)
                Catch ex As Exception
                    xlSheets.Item(sheetIndex).Protect(password, drawingObjects, contents, senarios, userInterfaceOnly)
                End Try
            Finally
                Me.Free(xlSheets)
            End Try
        End Sub

        ''' <summary>
        ''' セル内の部分文字列のフォントサイズを設定する
        ''' </summary>
        ''' <param name="row">対象セルの行インデックス</param>
        ''' <param name="col">対象セルの列インデックス</param>
        ''' <param name="start">対象文字列の開始位置</param>
        ''' <param name="length">対象文字列の長さ</param>
        ''' <param name="size">フォントサイズ</param>
        ''' <remarks></remarks>
        Public Sub SetCharactersFontSizeRC(ByVal row As Integer, _
                                           ByVal col As Integer, _
                                           ByVal start As Integer, _
                                           ByVal length As Integer, _
                                           ByVal size As Integer)
            Dim xlCell As Object = Me.GetCellRC(row, col)
            Try
                Dim xlCharacters As Object = xlCell.Characters(start, length)
                Try
                    xlCharacters.Font.Size = size
                Finally
                    Me.Free(xlCharacters)
                End Try
            Finally
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' (Accessが必要)コード39バーコードを追加する
        ''' </summary>
        ''' <param name="startRow">開始セル行</param>
        ''' <param name="startCol">開始セル列</param>
        ''' <param name="endRow">終端セル行</param>
        ''' <param name="endCol">終端セル列</param>
        ''' <param name="value">バーコード値</param>
        ''' <param name="appendsStartStopChar">スタート/ストップ文字を付加する場合、true</param>
        ''' <param name="showsData">データを表示する場合、true</param>
        ''' <remarks>バーコードの大きさは開始セル・終端セルそれぞれの左上座標を結ぶ</remarks>
        Public Sub AddBarCode39RC(ByVal startRow As Integer, ByVal startCol As Integer, _
                                  ByVal endRow As Integer, ByVal endCol As Integer, ByVal value As String, _
                                  Optional ByVal appendsStartStopChar As Boolean = True, Optional ByVal showsData As Boolean = True)
            Dim startCell As Point = GetCellTopLeftRC(startRow, startCol)
            Dim endCell As Point = GetCellTopLeftRC(endRow, endCol)
            AddBarCode39(startCell, New Size(endCell.X - startCell.X, endCell.Y - startCell.Y), value, appendsStartStopChar, showsData)
        End Sub

        ''' <summary>
        ''' (Accessが必要)コード39バーコードを追加する
        ''' </summary>
        ''' <param name="topLeftPoint">バーコード左上の位置</param>
        ''' <param name="aSize">バーコードの大きさ（幅と高さ）</param>
        ''' <param name="value">バーコード値</param>
        ''' <param name="appendsStartStopChar">スタート/ストップ文字を付加する場合、true</param>
        ''' <param name="showsData">データを表示する場合、true</param>
        ''' <remarks></remarks>
        Public Sub AddBarCode39(ByVal topLeftPoint As Point, ByVal aSize As Size, ByVal value As String, _
                                Optional ByVal appendsStartStopChar As Boolean = True, Optional ByVal showsData As Boolean = True)
            PerformAddBarCode(topLeftPoint, aSize, value, XlBarCodeStyle.CODE_39, 0, If(appendsStartStopChar, 1, 0), showsData, XlBarCodeDirection.NORMAL)
        End Sub

        ''' <summary>
        ''' バーコードの追加
        ''' </summary>
        ''' <param name="topLeftPoint">バーコード左上の位置</param>
        ''' <param name="aSize">バーコードの大きさ（幅と高さ）</param>
        ''' <param name="value">バーコード値</param>
        ''' <param name="style">バーコードのスタイル</param>
        ''' <param name="subStyle">スタイル毎に異なる「枝番」</param>
        ''' <param name="validation">スタイル毎に異なる「データの確認値」</param>
        ''' <param name="showsData">データを表示する場合、true</param>
        ''' <remarks>バーコードの大きさは開始セル・終端セルそれぞれの左上座標を結ぶ<br/>
        ''' バーコードコントロール詳細は http://msdn.microsoft.com/ja-jp/library/cc427149.aspx
        ''' </remarks>
        Private Sub PerformAddBarCode(ByVal topLeftPoint As Point, ByVal aSize As Size, _
                                      ByVal value As String, ByVal style As XlBarCodeStyle, ByVal subStyle As Integer, _
                                      ByVal validation As Integer, ByVal showsData As Boolean, ByVal direction As XlBarCodeDirection)
            Dim ole As Object = m_activeSheet.OLEObjects
            Try
                Dim barcode As Object
                Try
                    barcode = ole.Add(ClassType:="BARCODE.BarCodeCtrl.1")
                Catch ex As Exception
                    Throw New SystemException("バーコードを作成できません。Accessがインストールされているか確認してください。", ex)
                End Try
                Try
                    barcode.Locked = False

                    Dim barcode2 As Object = barcode.Object
                    Try
                        With barcode2
                            .Style = style
                            .SubStyle = subStyle
                            .Validation = validation
                            '.PrintObject = True
                            .ShowData = If(showsData, 1, 0)
                            '.LineWeight = 3
                            .Direction = direction
                            .Left = topLeftPoint.X
                            .Top = topLeftPoint.Y
                            .Width = aSize.Width
                            .Height = aSize.Height
                            .Value = value
                            .Refresh()
                        End With

                    Finally
                        Free(barcode2)
                    End Try

                    barcode.Visible = True
                Finally
                    Free(barcode)
                End Try
            Finally
                Free(ole)
            End Try

        End Sub

        ''' <summary>
        ''' PDF出力する
        ''' </summary>
        ''' <param name="saveFilePathName">出力ファイル名</param>
        ''' <remarks></remarks>
        <Obsolete("Excel2007以降が必要. それでも使用したいなら#Save2007AsPDFを使用して")> Public Sub SaveAsPDF(ByVal saveFilePathName As String)
            Save2007AsPDF(saveFilePathName)
        End Sub
        ''' <summary>
        ''' PDF出力する(Excel2007以降必須)
        ''' </summary>
        ''' <param name="saveFilePathName">出力ファイル名</param>
        ''' <remarks></remarks>
        Public Sub Save2007AsPDF(ByVal saveFilePathName As String)
            Dim ignorePrintArea As Boolean = False
            Dim openAfterPublish As Boolean = False
            Try
                m_xlBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, _
                                             saveFilePathName, _
                                             XlFixedFormatQuality.xlQualityMinimum, _
                                             True, ignorePrintArea, Type.Missing, Type.Missing, openAfterPublish, Type.Missing)
            Catch ex As System.MissingMemberException
                Throw New InvalidOperationException("Excel2007以降のインストールがないのでPDF出力できません", ex)
            End Try
        End Sub

        ''' <summary>
        ''' 選択セルを基点とした貼り付けをするリンク貼付
        ''' </summary>
        ''' <param name="row">行</param>
        ''' <param name="col">列</param>
        ''' <remarks></remarks>
        Public Sub PasteRC(ByVal row As Integer, ByVal col As Integer)

            Dim xlCell As Object = Me.GetCellRC(row, col)
            Try
                xlCell.Select()
                m_activeSheet.Paste()

            Finally
                Me.Free(xlCell)
            End Try

        End Sub

        ''' <summary>
        ''' 図の挿入
        ''' </summary>
        ''' <param name="row">行</param>
        ''' <param name="col">列</param>
        ''' <param name="fileName">挿入する図(ファイルパス付)</param>
        ''' <remarks></remarks>
        Public Sub InsertPictureRC(ByVal row As Integer, ByVal col As Integer, ByVal fileName As String)

            Dim xlCell As Object = Me.GetCellRC(row, col)
            Try
                Dim xlPictures As Object = m_activeSheet.Pictures
                Try
                    xlCell.Select()
                    xlPictures.Insert(fileName)
                Finally
                    Me.Free(xlPictures)
                End Try
            Finally
                Me.Free(xlCell)
            End Try

        End Sub

        ''' <summary>
        ''' Excelアプリケーションの表示を設定する
        ''' </summary>
        ''' <param name="visible">表示するなら、true</param>
        ''' <remarks></remarks>
        Public Sub SetAppVisible(ByVal visible As Boolean)
            m_xlApp.Visible = visible
        End Sub

        ''' <summary>
        ''' 改ページを挿入する
        ''' </summary>
        ''' <param name="row">行index</param>
        ''' <remarks></remarks>
        Public Sub SetPageBreakByRow(ByVal row As Integer)
            SetPageBreakRC(row, -1)
        End Sub

        ''' <summary>
        ''' 改ページを挿入する
        ''' </summary>
        ''' <param name="row">行</param>
        ''' <param name="col">列</param>
        ''' <remarks></remarks>
        Public Sub SetPageBreakRC(ByVal row As Integer, ByVal col As Integer)
            Dim xlCell As Object
            If 0 < col And 0 < row Then
                xlCell = Me.GetCellRC(row, col)
            ElseIf 0 < col Then
                xlCell = Me.GetColumn(col)
            ElseIf 0 < row Then
                xlCell = Me.GetRow(row)
            Else
                Throw New ArgumentException("colもrowも-1は不正")
            End If
            Try
                If 0 < row Then
                    Dim xlPageBreaks As Object = m_activeSheet.HPageBreaks
                    Try
                        xlPageBreaks.Add(Before:=xlCell)
                    Finally
                        Me.Free(xlPageBreaks)
                    End Try
                End If
                If 0 < col Then
                    Dim xlPageBreaks As Object = m_activeSheet.VPageBreaks
                    Try
                        xlPageBreaks.Add(Before:=xlCell)
                    Finally
                        Me.Free(xlPageBreaks)
                    End Try
                End If
            Finally
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' 「改ページ プレビュー」にする
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ChangeViewToPageBreakPreview()
            Dim xlWindow As Object = m_xlApp.ActiveWindow
            Try
                xlWindow.View = XlWindowView.xlPageBreakPreview
            Finally
                Me.Free(xlWindow)
            End Try
        End Sub

        ''' <summary>
        ''' 「標準」にする（「改ページ プレビュー」から戻す）
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ChangeViewToNormal()
            Dim xlWindow As Object = m_xlApp.ActiveWindow
            Try
                xlWindow.View = XlWindowView.xlNormalView
            Finally
                Me.Free(xlWindow)
            End Try
        End Sub

        ''' <summary>
        ''' 改頁プレビューの縦線を無しに（横1ページ印刷に）する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ClearVPageBreak()
            Dim xlVPageBreaks As Object = m_activeSheet.VPageBreaks
            Try
                If 0 < xlVPageBreaks.Count Then
                    Dim xlVPageBreak As Object = xlVPageBreaks(1)
                    Try
                        xlVPageBreak.DragOff(Direction:=XlDirection.xlToRight, RegionIndex:=1)
                    Catch ex As Exception
                        Throw New InvalidOperationException("0x800A03ECなら「改ページ プレビュー」になっていない可能性がある", ex)
                    Finally
                        Me.Free(xlVPageBreak)
                    End Try
                End If
            Finally
                Me.Free(xlVPageBreaks)
            End Try
        End Sub

        ''' <summary>(Excel設定の)標準フォント名</summary>
        ''' <remarks>Excelを再起動しないと設定した値が有効化されないので注意</remarks>
        Public Property StandardFont() As String
            Get
                Return m_xlApp.StandardFont
            End Get
            Set(ByVal value As String)
                m_xlApp.StandardFont = value
            End Set
        End Property

        ''' <summary>(Excel設定の)標準フォントサイズ</summary>
        ''' <remarks>Excelを再起動しないと設定した値が有効化されないので注意</remarks>
        Public Property StandardFontSize() As Double
            Get
                Return m_xlApp.StandardFontSize
            End Get
            Set(ByVal value As Double)
                m_xlApp.StandardFontSize = value
            End Set
        End Property

        ''' <summary>
        ''' シートの枠線を表示するか？を返す
        ''' </summary>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Function IsDisplayGridlines() As Boolean
            Dim xlWindow As Object = m_xlApp.ActiveWindow
            Try
                Return xlWindow.DisplayGridlines

            Finally
                Me.Free(xlWindow)
            End Try
        End Function

        ''' <summary>
        ''' シートの枠線表示を設定する
        ''' </summary>
        ''' <param name="displayGridlines">表示する場合、true</param>
        ''' <remarks></remarks>
        Public Sub SetDisplayGridlines(ByVal displayGridlines As Boolean)
            Dim xlWindow As Object = m_xlApp.ActiveWindow
            Try
                xlWindow.DisplayGridlines = displayGridlines

            Finally
                Me.Free(xlWindow)
            End Try
        End Sub

        ''' <summary>
        ''' シート数を設定する
        ''' </summary>
        ''' <param name="sheetCount">シート数</param>
        ''' <remarks></remarks>
        Public Sub SetSheetCount(ByVal sheetCount As Integer)
            Dim nowSheetCount As Integer = Me.SheetCount
            For sheetIndex As Integer = sheetCount + 1 To nowSheetCount
                Me.DeleteSheet(nowSheetCount - sheetIndex + sheetCount + 1)
            Next
        End Sub

        ''' <summary>
        ''' 指定した列範囲をグループ化します.
        ''' </summary>
        ''' <param name="startCol">開始列インデックス</param>
        ''' <param name="endCol">終了列インデックス[省略可]</param>
        ''' <remarks></remarks>
        Public Sub SetColumnsGroup(ByVal startCol As Integer, _
                                   Optional ByVal endCol As Integer = -1, _
                                   Optional ByVal isClose As Boolean = True)
            Dim xlColumn As Object = Nothing
            Dim xlSheets As Object = Nothing

            Try
                xlColumn = Me.GetColumn(startCol, endCol)
                xlColumn.Group()

                If isClose Then m_activeSheet.Outline.ShowLevels(ColumnLevels:=1)

            Finally
                Me.Free(xlSheets)
                Me.Free(xlColumn)
            End Try
        End Sub

        ''' <summary>イベントを動作させる場合、true</summary>
        Public Property EnableEvents() As Boolean
            Get
                Return m_xlApp.EnableEvents
            End Get
            Set(ByVal value As Boolean)
                m_xlApp.EnableEvents = value
            End Set
        End Property

        ''' <summary>表示更新をする場合、true</summary>
        Public Property ScreenUpdating() As Boolean
            Get
                Return m_xlApp.ScreenUpdating
            End Get
            Set(ByVal value As Boolean)
                m_xlApp.ScreenUpdating = value
            End Set
        End Property

        ''' <summary>
        ''' イメージを挿入する
        ''' </summary>
        ''' <param name="topLeftPoint"></param>
        ''' <param name="aSize"></param>
        ''' <param name="filePath"></param>
        ''' <remarks></remarks>
        Public Sub SetImage(ByVal topLeftPoint As Point, ByVal aSize As Size, ByVal filePath As String)
            Dim xlsShapes As Object = Nothing
            Dim xlsShape As Object = Nothing

            Try
                xlsShapes = m_activeSheet.Shapes
                xlsShape = xlsShapes.AddPicture(filePath, False, True, topLeftPoint.X, topLeftPoint.Y, aSize.Width, aSize.Height)
            Finally
                Me.Free(xlsShapes)
                Me.Free(xlsShape)
            End Try

        End Sub

        ''' <summary>
        ''' 現在のシートに存在するオートシェイプを指定したシートにコピーする
        ''' </summary>
        ''' <param name="sheetIndex">コピー先のシートインデックス</param>
        ''' <param name="shapeName">コピーするオートシェイプ名</param>
        ''' <remarks></remarks>
        Public Sub ShapeCopyPaste(ByVal sheetIndex As Integer, ByVal shapeName As String)
            m_activeSheet = m_xlBook.ActiveSheet

            Dim leftY As Double = Nothing
            Dim topX As Double = Nothing
            Dim heightH As Double = Nothing
            Dim widthW As Double = Nothing

            Dim copyShape As Object = Nothing
            Dim sheets As Object = Nothing
            Dim pasteSheet As Object = Nothing
            Dim pasteShape As Object = Nothing

            Try
                copyShape = m_activeSheet.Shapes(shapeName)
                leftY = copyShape.Left
                topX = copyShape.Top
                heightH = copyShape.Height
                widthW = copyShape.Width
                copyShape.Copy()

                sheets = m_xlBook.Sheets
                pasteSheet = sheets.Item(sheetIndex)
                pasteSheet.Paste()

                pasteShape = pasteSheet.Shapes(shapeName)
                pasteShape.Left = leftY
                pasteShape.Top = topX
                pasteShape.Height = heightH
                pasteShape.Width = widthW
            Finally
                Me.Free(pasteShape)
                Me.Free(pasteSheet)
                Me.Free(sheets)
                Me.Free(copyShape)
            End Try
        End Sub

        ''' <summary>
        ''' 印刷プレビュー
        ''' </summary>
        ''' <param name="enableChanges">印刷設定の変更を許可する場合、True</param>
        ''' <remarks></remarks>
        Public Sub PrintPreview(Optional ByVal enableChanges As Boolean = False)
            Dim workSheets As Object = m_xlBook.Sheets
            Try
                workSheets.PrintPreview(enableChanges)
            Finally
                Me.Free(workSheets)
            End Try
        End Sub

        ''' <summary>データ入力規則</summary>
        Private Enum XlDVType
            ''' <summary>すべての値</summary>
            xlValidateInputOnly = 0
            ''' <summary>整数</summary>
            xlValidateWholeNumber = 1
            ''' <summary>小数点数</summary>
            xlValidateDecimal = 2
            ''' <summary>リスト</summary>
            xlValidateList = 3
            ''' <summary>日付</summary>
            xlValidateDate = 4
            ''' <summary>時刻</summary>
            xlValidateTime = 5
            ''' <summary>文字列（長さ指定）</summary>
            xlValidateTextLength = 6
            ''' <summary>ユーザー設定</summary>
            xlValidateCustom = 7
        End Enum

        ''' <summary>
        ''' コンボボックス設定をする
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <param name="valueFormula">参照範囲 ex."=Sheet1!C5:C7"</param>
        ''' <remarks></remarks>
        Public Sub SetComboBoxFormula(axis As XlAxis, valueFormula As String)
            SetComboBoxFormula(axis.Row, axis.Column, valueFormula)
        End Sub
        ''' <summary>
        ''' コンボボックス設定をする
        ''' </summary>
        ''' <param name="row">行index</param>
        ''' <param name="column">列index</param>
        ''' <param name="valueFormula">カンマ区切りの値(ex."有,無")、または、セル範囲(ex."=Sheet1!C5:C7")</param>
        ''' <remarks></remarks>
        Public Sub SetComboBoxFormula(row As Integer, column As Integer, valueFormula As String)
            EzUtil.AssertParameterIsNotEmpty(valueFormula, "valueFormula")
            Dim xlCell As Object = GetCellRC(row, column)
            Try
                Dim xlValidation As Object = xlCell.Validation
                Try
                    xlValidation.Add(Type:=XlDVType.xlValidateList, Formula1:=valueFormula)
                Finally
                    Free(xlValidation)
                End Try
            Finally
                Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' コンボボックス設定の選択できる値すべてを取得する
        ''' </summary>
        ''' <param name="axis">座標</param>
        ''' <returns>選択値[]</returns>
        ''' <remarks></remarks>
        Public Function ExtractComboBoxValues(axis As XlAxis) As String()
            Return ExtractComboBoxValues(axis.Row, axis.Column)
        End Function
        ''' <summary>
        ''' コンボボックス設定の選択できる値すべてを取得する
        ''' </summary>
        ''' <param name="row">行index</param>
        ''' <param name="column">列index</param>
        ''' <returns>選択値[]</returns>
        ''' <remarks></remarks>
        Public Function ExtractComboBoxValues(row As Integer, column As Integer) As String()
            Dim xlCell As Object = GetCellRC(row, column)
            Try
                Dim xlValidation As Object = xlCell.Validation
                Try
                    If XlDVType.xlValidateList <> xlValidation.Type Then
                        Return Nothing
                    End If
                    Dim formula As String = xlValidation.Formula1
                    If StringUtil.IsEmpty(formula) Then
                        Throw New InvalidOperationException("リスト設定なのに選択値、またはセル範囲が未設定")
                    End If
                    If formula.StartsWith("=") Then
                        Dim xlRange As Object = m_xlApp.Evaluate(formula.Substring(1))
                        Try
                            Dim data As Object(,) = ConvCurrentRangeToTwoDimensionalArray(xlRange)
                            Return Enumerable.Range(0, UBound(data) + 1).Select(Function(i) StringUtil.ToString(data(i, 0))).ToArray
                        Finally
                            Free(xlRange)
                        End Try
                    Else
                        Return Split(formula, ",")
                    End If
                Finally
                    Free(xlValidation)
                End Try
            Finally
                Free(xlCell)
            End Try
        End Function

        ''' <summary>
        ''' Range#Row,Range#Columnの範囲を2次元配列にする（0はじまり）
        ''' </summary>
        ''' <param name="xlRange"></param>
        ''' <returns></returns>
        ''' <remarks>Range#Countは考慮していない</remarks>
        Private Function ConvCurrentRangeToTwoDimensionalArray(xlRange As Object) As Object(,)
            Dim rowCount As Integer = xlRange.Rows.Count
            Dim columnCount As Integer = xlRange.Columns.Count
            Dim resultData(rowCount - 1, columnCount - 1) As Object
            For row As Integer = 1 To rowCount
                For column As Integer = 1 To columnCount
                    Dim xlCell As Object = xlRange.Item(row, column)
                    Try
                        resultData(row - 1, column - 1) = xlCell.Value
                    Finally
                        Free(xlCell)
                    End Try
                Next
            Next
            Return resultData
        End Function

        ''' <summary>
        ''' 「データの入力規則」を設定している座標をすべて取得する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DetectAxisesOfAllValidation() As XlAxis()
            Dim xlCells As Object = m_activeSheet.Cells
            Try
                Dim xlRanges As Object = Nothing
                Try
                    xlRanges = xlCells.SpecialCells(XlCellType.xlCellTypeAllValidation)
                Catch ex As COMException
                    Me.Free(xlRanges)
                    Const NOT_FOUND As Integer = -2146827284
                    If ex.ErrorCode = NOT_FOUND Then
                        Return New XlAxis() {}
                    End If
                    Throw
                End Try
                Try
                    Dim xlAreas As Object = xlRanges.Areas
                    Try
                        Dim results As New List(Of XlAxis)
                        For i As Integer = 1 To xlAreas.Count
                            Dim xlRange As Object = xlAreas.Item(i)
                            Try
                                Dim xlRows As Object = xlRange.Rows
                                Try
                                    Dim xlColumns As Object = xlRange.Columns
                                    Try
                                        For row As Integer = 1 To xlRows.Count
                                            For column As Integer = 1 To xlColumns.Count
                                                Dim xlCell As Object = xlRange.Item(row, column)
                                                Try
                                                    results.Add(New XlAxis(xlCell.Row, xlCell.Column))
                                                Finally
                                                    Free(xlCell)
                                                End Try
                                            Next
                                        Next
                                    Finally
                                        Free(xlColumns)
                                    End Try
                                Finally
                                    Free(xlRows)
                                End Try
                            Finally
                                Free(xlRange)
                            End Try
                        Next
                        Return results.ToArray
                    Finally
                        Free(xlAreas)
                    End Try
                Finally
                    Me.Free(xlRanges)
                End Try
            Finally
                Me.Free(xlCells)
            End Try
        End Function

        Private Enum MsoTextOrientation
            msoTextOrientationHorizontal = 1
            msoTextOrientationVertical = 5
            ' 他にもある
        End Enum

        ''' <summary>
        ''' テキストボックスシェイプを追加する
        ''' </summary>
        ''' <param name="left">開始位置左</param>
        ''' <param name="top">開始位置上</param>
        ''' <param name="width">幅</param>
        ''' <param name="height">高さ</param>
        ''' <param name="text">テキスト文字列</param>
        ''' <returns>シェイプ名</returns>
        ''' <remarks></remarks>
        Public Function AddTextBox(ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single, ByVal text As String) As String
            Dim xlsShapes As Object = m_activeSheet.Shapes
            Try
                Dim xlsShape As Object = xlsShapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height)
                Try
                    Dim xlsTextFrame As Object = xlsShape.TextFrame
                    Try
                        Dim xlsCharacters As Object = xlsTextFrame.Characters
                        Try
                            xlsCharacters.Text = text
                        Finally
                            Me.Free(xlsCharacters)
                        End Try
                    Finally
                        Me.Free(xlsTextFrame)
                    End Try
                    Return xlsShape.Name
                Finally
                    Me.Free(xlsShape)
                End Try
            Finally
                Me.Free(xlsShapes)
            End Try
        End Function

        ''' <summary>
        ''' シェイプのテキスト値とシェイプ名を取得する
        ''' </summary>
        ''' <remarks></remarks>
        Public Function DetectShapeTextAndName() As Dictionary(Of String, String)
            Dim textByName As New Dictionary(Of String, String)
            Dim xlsShapes As Object = m_activeSheet.Shapes
            Try
                For i As Integer = 1 To xlsShapes.Count
                    Dim xlsShape As Object = xlsShapes.Item(i)
                    Try
                        Dim xlsTextFrame As Object = xlsShape.TextFrame
                        Try
                            Dim xlsCharacters As Object = xlsTextFrame.Characters
                            Try
                                textByName.Add(xlsShape.Name, xlsCharacters.Text)
                            Finally
                                Me.Free(xlsCharacters)
                            End Try
                        Catch ignore As COMException
                            ' nop
                        Finally
                            Me.Free(xlsTextFrame)
                        End Try
                    Finally
                        Me.Free(xlsShape)
                    End Try
                Next
            Finally
                Me.Free(xlsShapes)
            End Try
            Return textByName
        End Function

        ''' <summary>
        ''' 条件付き書式を追加する
        ''' </summary>
        ''' <param name="startRow">開始行</param>
        ''' <param name="startColumn">開始列</param>
        ''' <param name="endRow">終端行</param>
        ''' <param name="endColumn">終端列</param>
        ''' <param name="formula">式</param>
        ''' <param name="rgbColor">色</param>
        ''' <remarks></remarks>
        Public Sub AddFormatCondition(ByVal startRow As Integer, ByVal startColumn As Integer, ByVal endRow As Integer, ByVal endColumn As Integer, ByVal formula As String, ByVal rgbColor As Integer)
            Const xlExpression As Integer = 2
            Dim xlCell As Object = Me.GetCellRC(startRow, startColumn, endRow, endColumn)
            Try
                Dim xlFormatConditions As Object = xlCell.FormatConditions()
                Try
                    Dim xlFormatCondition As Object = xlFormatConditions.Add(Type:=xlExpression, Formula1:=formula)
                    Try
                        Dim xlInterior As Object = xlFormatCondition.Interior
                        Try
                            xlInterior.Color = rgbColor
                        Finally
                            Me.Free(xlInterior)
                        End Try
                        xlFormatCondition.StopIfTrue = True

                    Finally
                        Me.Free(xlFormatCondition)
                    End Try
                Finally
                    Me.Free(xlFormatConditions)
                End Try
            Finally
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' 条件付き書式を消去する
        ''' </summary>
        ''' <param name="startRow">開始行</param>
        ''' <param name="startColumn">開始列</param>
        ''' <param name="endRow">終端行</param>
        ''' <param name="endColumn">終端列</param>
        ''' <remarks></remarks>
        Public Sub ClearFormatCondition(ByVal startRow As Integer, ByVal startColumn As Integer, ByVal endRow As Integer, ByVal endColumn As Integer)
            Dim xlCell As Object = Me.GetCellRC(startRow, startColumn, endRow, endColumn)
            Try
                Dim xlFormatConditions As Object = xlCell.FormatConditions()
                Try
                    xlFormatConditions.Delete()
                Finally
                    Me.Free(xlFormatConditions)
                End Try
            Finally
                Me.Free(xlCell)
            End Try
        End Sub

        ''' <summary>
        ''' Shapeを複製する
        ''' </summary>
        ''' <param name="baseShapeName">複製元のShape名</param>
        ''' <param name="distShapeName">複製先のShape名</param>
        ''' <remarks></remarks>
        Public Sub CloneShape(baseShapeName As String, distShapeName As String)
            EzUtil.AssertParameterIsNotEmpty(baseShapeName, "baseShapeName")
            EzUtil.AssertParameterIsNotEmpty(distShapeName, "distShapeName")
            If baseShapeName.Equals(distShapeName) Then
                Throw New ArgumentException("元となるShapeと新しいShapeの名前が同じ")
            End If

            m_activeSheet = m_xlBook.ActiveSheet

            Dim baseShape As Object = m_activeSheet.Shapes(baseShapeName)
            Try

                Dim distShape As Object = baseShape.Duplicate()
                Try
                    distShape.Name = distShapeName
                Finally
                    Me.Free(distShape)
                End Try
            Finally
                Me.Free(baseShape)
            End Try
        End Sub

        ''' <summary>
        ''' Shapeを削除する
        ''' </summary>
        ''' <param name="targetShapeNames">削除したいShape名[]</param>
        ''' <remarks></remarks>
        Public Sub DeleteShapes(ParamArray targetShapeNames As String())
            m_activeSheet = m_xlBook.ActiveSheet

            Dim shapes As Object = m_activeSheet.Shapes
            Try
                For i As Integer = shapes.Count To 1 Step -1
                    Dim shape As Object = shapes(i)
                    Try
                        If Not targetShapeNames.Contains(shape.Name) Then
                            Continue For
                        End If
                        shape.Delete()
                    Finally
                        Me.Free(shape)
                    End Try
                Next
            Finally
                Me.Free(shapes)
            End Try
        End Sub

        ''' <summary>
        ''' セル範囲
        ''' </summary>
        ''' <remarks></remarks>
        Public Class RangeCell
            ''' <summary>開始列</summary>
            Public StartCol As Integer
            ''' <summary>開始行</summary>
            Public StartRow As Integer
            ''' <summary>終了列</summary>
            Public EndCol As Integer
            ''' <summary>終了行</summary>
            Public EndRow As Integer
        End Class

        Private Function MakeRangeParam(rangeCells As List(Of RangeCell)) As String
            If CollectionUtil.IsEmpty(rangeCells) Then
                Return Nothing
            End If
            Return Join(rangeCells.Select(Function(rangeCell) MakeRangeFormat(rangeCell)).ToArray, ",")
        End Function

        Private Function MakeRangeFormat(rangeCell As RangeCell) As String
            Return MakeRangeFormat(rangeCell.StartCol, rangeCell.StartRow, rangeCell.EndCol, rangeCell.EndRow)
        End Function

        Private Function MakeRangeFormat(startCol As Integer, startRow As Integer, endCol As Integer, endRow As Integer) As String
            Return String.Format("{0}{1}:{2}{3}", ConvertToLetter(startCol), startRow, ConvertToLetter(endCol), endRow)
        End Function

        ''' <summary>
        ''' 指定された複数のセル範囲をまとめて結合します.
        ''' </summary>
        ''' <param name="rangeCells">セル範囲[]</param>
        ''' <param name="merge"></param>
        ''' <remarks></remarks>
        Public Sub MergeMultiCells(rangeCells As List(Of RangeCell), merge As Boolean)
            Dim rangeParam As String = MakeRangeParam(rangeCells)
            If StringUtil.IsEmpty(rangeParam) Then
                Return
            End If
            Dim xlRange As Object = Me.GetCell(rangeParam)
            Try
                xlRange.MergeCells = merge
            Finally
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 指定された複数のセル範囲に罫線を設定します.
        ''' </summary>
        ''' <param name="rangeCells">セル範囲[]</param>
        ''' <param name="lineStyle">罫線の種類[省略可]</param>
        ''' <param name="weight">罫線の太さ[省略可]</param>
        ''' <remarks></remarks>
        Public Sub SetLineMultiCells(rangeCells As List(Of RangeCell), _
                                     Optional lineStyle As XlLineStyle = XlLineStyle.xlContinuous, Optional weight As XlBorderWeight = XlBorderWeight.xlThin)
            Dim rangeStr As String = MakeRangeParam(rangeCells)
            If StringUtil.IsEmpty(rangeStr) Then
                Return
            End If
            Dim xlRange As Object = Me.GetCell(rangeStr)
            Try
                Dim xlBorders As Object = xlRange.Borders
                Try
                    xlBorders.LineStyle = lineStyle
                    If Not XlLineStyle.xlLineStyleNone.Equals(lineStyle) Then
                        xlBorders.Weight = weight
                    End If
                Finally
                    Me.Free(xlBorders)
                End Try
            Finally
                Me.Free(xlRange)
            End Try
        End Sub

        ''' <summary>
        ''' 指定されたシートのヘッダー/フッター設定を行います.
        ''' </summary>
        ''' <param name="names">シート名[]</param>
        ''' <param name="headerString">ヘッダーに設定する内容(任意)</param>
        ''' <param name="footerString">フッターに設定する内容(任意)</param>
        ''' <remarks></remarks>
        Public Sub SetMultiSheetHeaderFooter(names As String(), Optional headerString As String = "", Optional footerString As String = "")
            SetMultiSheetHeaderFooter(GetSheetIndexesByNames(names), headerString, footerString)
        End Sub

        ''' <summary>
        ''' 指定されたシートのヘッダー/フッター設定を行います.
        ''' </summary>
        ''' <param name="indexes">シートインデックス[]</param>
        ''' <param name="headerString">ヘッダーに設定する内容(任意)</param>
        ''' <param name="footerString">フッターに設定する内容(任意)</param>
        ''' <remarks></remarks>
        Public Sub SetMultiSheetHeaderFooter(indexes As Integer(), Optional headerString As String = "", Optional footerString As String = "")
            Dim xlSheets As Object = m_xlBook.Sheets
            Try
                m_xlApp.PrintCommunication = False
                For Each index As Integer In indexes
                    Dim xlSheet As Object = xlSheets.Item(index)
                    Try
                        With xlSheet.PageSetup
                            .RightHeader = headerString
                            .RightFooter = footerString
                        End With
                    Finally
                        Me.Free(xlSheet)
                    End Try
                Next
            Finally
                Me.Free(xlSheets)
                m_xlApp.PrintCommunication = True
            End Try
        End Sub

    End Class

#Region "定数"

    ''' <summary>ファイル形式</summary>
    Public Enum XlFileFormat
        ''' <summary>Microsoft Office Excel アドイン </summary>
        xlAddIn = 18
        ''' <summary>???</summary>
        xlAddIn8 = 18
        ''' <summary>コンマ区切りの値 </summary>
        xlCSV = 6
        ''' <summary>コンマ区切りの値 </summary>
        xlCSVMac = 22
        ''' <summary>コンマ区切りの値</summary>
        xlCSVMSDOS = 24
        ''' <summary>コンマ区切りの値 </summary>
        xlCSVWindows = 23
        ''' <summary>テキスト形式の種類を指定します。 </summary>
        xlCurrentPlatformText = -4158
        ''' <summary>Dbase 2 形式 </summary>
        xlDBF2 = 7
        ''' <summary>Dbase 3 形式 </summary>
        xlDBF3 = 8
        ''' <summary>Dbase 4 形式 </summary>
        xlDBF4 = 11
        ''' <summary>データ交換形式 </summary>
        xlDIF = 9
        ''' <summary>Excel 12.0 </summary>
        xlExcel12 = 50
        ''' <summary>Excel 2.0 </summary>
        xlExcel2 = 16
        ''' <summary>Excel 2.0 (東アジア版) </summary>
        xlExcel2FarEast = 27
        ''' <summary>Excel 3.0 </summary>
        xlExcel3 = 29
        ''' <summary>Excel 4.0 </summary>
        xlExcel4 = 33
        ''' <summary>Excel 4.0、ワークブック形式 </summary>
        xlExcel4Workbook = 35
        ''' <summary>Excel 5.0 </summary>
        xlExcel5 = 39
        ''' <summary>Excel 95 </summary>
        xlExcel7 = 39
        ''' <summary>Excel 95 </summary>
        xlExcel8 = 56
        ''' <summary>Excel 95 および Excel 97 </summary>
        xlExcel9795 = 43
        ''' <summary>Web ページ形式 </summary>
        xlHtml = 44
        ''' <summary>Microsoft Office Excel アドインのインターナショナルな形式 </summary>
        xlIntlAddIn = 26
        ''' <summary>不適切な形式 </summary>
        xlIntlMacro = 25
        ''' <summary>???</summary>
        xlOpenXMLAddIn = 55
        ''' <summary>???</summary>
        xlOpenXMLTemplate = 54
        ''' <summary>???</summary>
        xlOpenXMLTemplateMacroEnabled = 53
        ''' <summary>???</summary>
        xlOpenXMLWorkbook = 51
        ''' <summary>???</summary>
        xlOpenXMLWorkbookMacroEnabled = 52
        ''' <summary>シンボリック リンク形式 </summary>
        xlSYLK = 2
        ''' <summary>Excel テンプレートの形式 </summary>
        xlTemplate = 17
        ''' <summary>???</summary>
        xlTemplate8 = 17
        ''' <summary>テキスト形式の種類を指定します。 </summary>
        xlTextMac = 19
        ''' <summary>テキスト形式の種類を指定します。</summary>
        xlTextMSDOS = 21
        ''' <summary>テキスト形式の種類を指定します。 </summary>
        xlTextPrinter = 36
        ''' <summary>テキスト形式の種類を指定します。</summary>
        xlTextWindows = 20
        ''' <summary>テキスト形式の種類を指定します。 </summary>
        xlUnicodeText = 42
        ''' <summary>MHT 形式 </summary>
        xlWebArchive = 45
        ''' <summary>不適切な形式 </summary>
        xlWJ2WD1 = 14
        ''' <summary>不適切な形式 </summary>
        xlWJ3 = 40
        ''' <summary>不適切な形式 </summary>
        xlWJ3FJ3 = 41
        ''' <summary>Lotus 1-2-3 形式 </summary>
        xlWK1 = 5
        ''' <summary>Lotus 1-2-3 形式 </summary>
        xlWK1ALL = 31
        ''' <summary>Lotus 1-2-3 形式 </summary>
        xlWK1FMT = 30
        ''' <summary>Lotus 1-2-3 形式 </summary>
        xlWK3 = 15
        ''' <summary>Lotus 1-2-3 形式 </summary>
        xlWK3FM3 = 32
        ''' <summary>Lotus 1-2-3 形式 </summary>
        xlWK4 = 38
        ''' <summary>Lotus 1-2-3 形式 </summary>
        xlWKS = 4
        ''' <summary>???</summary>
        xlWorkbookDefault = 51
        ''' <summary>Excel ブック形式 </summary>
        xlWorkbookNormal = -4143
        ''' <summary>Microsoft Works 2.0 形式 </summary>
        xlWorks2FarEast = 28
        ''' <summary>Quattro Pro 形式 </summary>
        xlWQ1 = 34
        ''' <summary>Excel シート形式 </summary>
        xlXMLSpreadsheet = 46
    End Enum



    ''' <summary>罫線の場所</summary>
    Public Enum XlBordersIndex
        ''' <summary>セル範囲の各セルの左上隅から右下隅への罫線</summary>
        xlDiagonalDown = 5
        ''' <summary>セル範囲の各セルの左下隅から右上隅への罫線</summary>
        xlDiagonalUp = 6
        ''' <summary>セル範囲の左側の罫線</summary>
        xlEdgeLeft = 7
        ''' <summary>セル範囲の上側の罫線</summary>
        xlEdgeTop = 8
        ''' <summary>セル範囲の下側の罫線</summary>
        xlEdgeBottom = 9
        ''' <summary>セル範囲の右側の罫線</summary>
        xlEdgeRight = 10
        ''' <summary>セル範囲の外枠を除く、すべてのセルの垂直方向の罫線</summary>
        xlInsideVertical = 11
        ''' <summary>セル範囲の外枠を除く、すべてのセルの水平方向の罫線</summary>
        xlInsideHorizontal = 12
    End Enum



    ''' <summary>罫線の線種</summary>
    Public Enum XlLineStyle
        ''' <summary>実線</summary>
        xlContinuous = 1
        ''' <summary>一点鎖線</summary>
        xlDashDot = 4
        ''' <summary>二点鎖線</summary>
        xlDashDotDot = 5
        ''' <summary>斜線</summary>
        xlSlantDashDot = 13
        ''' <summary>破線</summary>
        xlDash = -4115
        ''' <summary>点線</summary>
        xlDot = -4118
        ''' <summary>二重線</summary>
        xlDouble = -4119
        ''' <summary>線なし</summary>
        xlLineStyleNone = -4142
    End Enum



    ''' <summary>罫線の太さ</summary>
    Public Enum XlBorderWeight
        ''' <summary>極細(最も細い罫線)</summary>
        xlHairline = 1
        ''' <summary>細い</summary>
        xlThin = 2
        ''' <summary>太い (最も太い罫線)</summary>
        xlThick = 4
        ''' <summary>中</summary>
        xlMedium = -4138
    End Enum



    ''' <summary>横配置</summary>
    Public Enum XlHAlign
        ''' <summary>標準</summary>
        xlHAlignGeneral = 1
        ''' <summary>繰り返し</summary>
        xlHAlignFill = 5
        ''' <summary>選択範囲内</summary>
        xlHAlignCenterAcrossSelection = 7
        ''' <summary>中央揃え</summary>
        xlHAlignCenter = -4108
        ''' <summary>均等割り付け</summary>
        xlHAlignDistributed = -4117
        ''' <summary>両端揃え</summary>
        xlHAlignJustify = -4130
        ''' <summary>左詰め</summary>
        xlHAlignLeft = -4131
        ''' <summary>右詰め</summary>
        xlHAlignRight = -4152
    End Enum



    ''' <summary>縦配置</summary>
    Public Enum XlVAlign
        ''' <summary>下詰め</summary>
        xlVAlignBottom = -4107
        ''' <summary>中央揃え</summary>
        xlVAlignCenter = -4108
        ''' <summary>均等割り付け</summary>
        xlVAlignDistributed = -4117
        ''' <summary>両端揃え</summary>
        xlVAlignJustify = -4130
        ''' <summary>上詰め</summary>
        xlVAlignTop = -4160
    End Enum



    ''' <summary>セルタイプ</summary>
    Public Enum XlCellType
        xlCellTypeAllFormatConditions = -4172
        xlCellTypeAllValidation = -4174
        xlCellTypeBlanks = 4
        xlCellTypeComments = -4144
        xlCellTypeConstants = 2
        xlCellTypeFormulas = -4123
        xlCellTypeLastCell = 11
        xlCellTypeSameFormatConditions = -4173
        xlCellTypeSameValidation = -4175
        xlCellTypeVisible = 12
    End Enum



    ''' <summary>検索対象の種類</summary>
    Public Enum XlFindLookIn
        ''' <summary>数式</summary>
        xlFormulas = -4123
        ''' <summary>コメント</summary>
        xlComments = -4144
        ''' <summary>値</summary>
        xlValues = -4163
    End Enum



    ''' <summary>検索方法</summary>
    Public Enum XlLookAt
        ''' <summary>全てが一致するセルを検索</summary>
        xlWhole = 1
        ''' <summary>一部が一致するセルを検索</summary>
        xlPart = 2
    End Enum



    ''' <summary>検索方向</summary>
    Public Enum XlSearchOrder
        ''' <summary>行</summary>
        xlByRows = 1
        ''' <summary>列</summary>
        xlByColumns = 2
    End Enum



    ''' <summary>検索順序</summary>
    Public Enum XlSearchDirection
        ''' <summary>順方向</summary>
        xlNext = 1
        ''' <summary>逆方向</summary>
        xlPrevious = 2
    End Enum



    ''' <summary>3 ステートのブール型 (Boolean) の値を指定します。</summary>
    Public Enum MsoTriState
        ''' <summary>サポートされていません。</summary>
        msoTriStateToggle = -3
        ''' <summary>サポートされていません。</summary>
        msoTriStateMixed = -2
        ''' <summary>真 (True)</summary>
        msoTrue = -1
        ''' <summary>偽 (False)</summary>
        msoFalse = 0
        ''' <summary>サポートされていません。</summary>
        msoCTrue = 1
    End Enum



    ''' <summary>図形を拡大または縮小したときに、位置を固定する部分を指定します。</summary>
    Public Enum MsoScaleFrom
        ''' <summary>図形の左上端の位置を保持します。</summary>
        msoScaleFromTopLeft = 0
        ''' <summary>図形の中点の位置を保持します。</summary>
        msoScaleFromMiddle = 1
        ''' <summary>図形の右下端の位置を保持します。</summary>
        msoScaleFromBottomRight = 2
    End Enum



    ''' <summary>計算方法</summary>
    Public Enum XlCalculation
        ''' <summary>自動</summary>
        xlCalculationAutomatic = -4105
        ''' <summary>手動</summary>
        xlCalculationManual = -4135
        ''' <summary>テーブル以外自動</summary>
        xlCalculationSemiautomatic = 2
    End Enum

    ''' <summary>
    ''' 文字列の向きを指定します
    ''' </summary>
    Public Enum XlOrientation
        ''' <summary>下向き</summary>
        xlDownward = -4170
        ''' <summary>水平</summary>
        xlHorizontal = -4128
        ''' <summary>上向き</summary>
        xlUpward = -4171
        ''' <summary>垂直 (縦書き)</summary>
        xlVertical = -4166
    End Enum

    ''' <summary>用紙サイズ</summary>
    Public Enum XlPaperSize
        xlPaperLetter = 1
        xlPaperLetterSmall = 2
        xlPaperTabloid = 3
        xlPaperLedger = 4
        xlPaperLegal = 5
        xlPaperStatement = 6
        xlPaperExecutive = 7
        ''' <summary>A3</summary>
        xlPaperA3 = 8
        ''' <summary>A4</summary>
        xlPaperA4 = 9
        xlPaperA4Small = 10
        xlPaperA5 = 11
        xlPaperB4 = 12
        xlPaperB5 = 13
        xlPaperFolio = 14
        xlPaperQuarto = 15
        xlPaper10x14 = 16
        xlPaper11x17 = 17
        xlPaperNote = 18
        xlPaperEnvelope9 = 19
        xlPaperEnvelope10 = 20
        xlPaperEnvelope11 = 21
        xlPaperEnvelope12 = 22
        xlPaperEnvelope14 = 23
        xlPaperCsheet = 24
        xlPaperDsheet = 25
        xlPaperEsheet = 26
        xlPaperEnvelopeDL = 27
        xlPaperEnvelopeC5 = 28
        xlPaperEnvelopeC3 = 29
        xlPaperEnvelopeC4 = 30
        xlPaperEnvelopeC6 = 31
        xlPaperEnvelopeC65 = 32
        xlPaperEnvelopeB4 = 33
        xlPaperEnvelopeB5 = 34
        xlPaperEnvelopeB6 = 35
        xlPaperEnvelopeItaly = 36
        xlPaperEnvelopeMonarch = 37
        xlPaperEnvelopePersonal = 38
        xlPaperFanfoldUS = 39
        xlPaperFanfoldStdGerman = 40
        xlPaperFanfoldLegalGerman = 41
        xlPaperUser = 256
    End Enum

    ''' <summary>印刷の向き</summary>
    ''' <remarks>ExcelVBAオブジェクトブラウザ: msoOrientation</remarks>
    Public Enum XlPageOrientation
        ''' <summary>縦</summary>
        xlPortrait = 1
        ''' <summary>横</summary>
        xlLandscape = 2
    End Enum

    ''' <summary>シェイプタイプ</summary>
    ''' <remarks>ExcelVBAオブジェクトブラウザ: msoAutoShapeType</remarks>
    Public Enum XlAutoShapeType
        ''' <summary>楕円</summary>
        msoShapeOval = 9
    End Enum

#End Region

End Namespace
