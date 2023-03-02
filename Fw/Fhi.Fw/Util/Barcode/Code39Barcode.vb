Imports System.Text
Imports System.Drawing.Imaging

Namespace Util.Barcode
    ''' <summary>
    ''' CODE39バーコードを担うクラス
    ''' </summary>
    ''' <remarks>ex. <code>Call (New Code39Barcode With {.PrintsContent = True}).Create("ABC123").Save("abc123.bmp")</code></remarks>
    Public Class Code39Barcode

        Private Shared ReadOnly BITS_BY_CHAR As Dictionary(Of Char, String)
        Shared Sub New()
            BITS_BY_CHAR = New Dictionary(Of Char, String)
            BITS_BY_CHAR.Add("0"c, "101000111011101")
            BITS_BY_CHAR.Add("1"c, "111010001010111")
            BITS_BY_CHAR.Add("2"c, "101110001010111")
            BITS_BY_CHAR.Add("3"c, "111011100010101")
            BITS_BY_CHAR.Add("4"c, "101000111010111")
            BITS_BY_CHAR.Add("5"c, "111010001110101")
            BITS_BY_CHAR.Add("6"c, "101110001110101")
            BITS_BY_CHAR.Add("7"c, "101000101110111")
            BITS_BY_CHAR.Add("8"c, "111010001011101")
            BITS_BY_CHAR.Add("9"c, "101110001011101")
            BITS_BY_CHAR.Add("A"c, "111010100010111")
            BITS_BY_CHAR.Add("B"c, "101110100010111")
            BITS_BY_CHAR.Add("C"c, "111011101000101")
            BITS_BY_CHAR.Add("D"c, "101011100010111")
            BITS_BY_CHAR.Add("E"c, "111010111000101")
            BITS_BY_CHAR.Add("F"c, "101110111000101")
            BITS_BY_CHAR.Add("G"c, "101010001110111")
            BITS_BY_CHAR.Add("H"c, "111010100011101")
            BITS_BY_CHAR.Add("I"c, "101110100011101")
            BITS_BY_CHAR.Add("J"c, "101011100011101")
            BITS_BY_CHAR.Add("K"c, "111010101000111")
            BITS_BY_CHAR.Add("L"c, "101110101000111")
            BITS_BY_CHAR.Add("M"c, "111011101010001")
            BITS_BY_CHAR.Add("N"c, "101011101000111")
            BITS_BY_CHAR.Add("O"c, "111010111010001")
            BITS_BY_CHAR.Add("P"c, "101110111010001")
            BITS_BY_CHAR.Add("Q"c, "101010111000111")
            BITS_BY_CHAR.Add("R"c, "111010101110001")
            BITS_BY_CHAR.Add("S"c, "101110101110001")
            BITS_BY_CHAR.Add("T"c, "101011101110001")
            BITS_BY_CHAR.Add("U"c, "111000101010111")
            BITS_BY_CHAR.Add("V"c, "100011101010111")
            BITS_BY_CHAR.Add("W"c, "111000111010101")
            BITS_BY_CHAR.Add("X"c, "100010111010111")
            BITS_BY_CHAR.Add("Y"c, "111000101110101")
            BITS_BY_CHAR.Add("Z"c, "100011101110101")
            BITS_BY_CHAR.Add("-"c, "100010101110111")
            BITS_BY_CHAR.Add("."c, "111000101011101")
            BITS_BY_CHAR.Add(" "c, "100011101011101")
            BITS_BY_CHAR.Add("$"c, "100010001000101")
            BITS_BY_CHAR.Add("/"c, "100010001010001")
            BITS_BY_CHAR.Add("+"c, "100010100010001")
            BITS_BY_CHAR.Add("%"c, "101000100010001")
        End Sub

#Region "Public properties..."
        ''' <summary>バーコードの細バーの幅</summary>
        Public NarrowBarWidth As Integer = 1
        ''' <summary>バーコードの高さ</summary>
        Public BarHeight As Integer = 40
        ''' <summary>バーコードの色</summary>
        Public BarColor As Color = Color.Black
        ''' <summary>バーコードの背景色</summary>
        Public BackGroundColor As Color = Color.White
        ''' <summary>バーコードにクワイエットゾーンが不要ならtrue</summary>
        Public IgnoresQuietZone As Boolean = False
        ''' <summary>前後の`*`を出力するならtrue</summary>
        Public PrintsCheckAster As Boolean = False
        ''' <summary>文字のフォントの色</summary>
        Public FontColor As Color = Color.Black
        ''' <summary>文字のフォントの名前</summary>
        Public FontName As String = "Tahoma"
        ''' <summary>文字のフォントの大きさ</summary>
        Public FontSize As Integer = 8

        Private Enum PrintContentStatus
            Unnecessary
            Overlap
            Outside
        End Enum
        Private _printContentStatus As PrintContentStatus = PrintContentStatus.Unnecessary
        ''' <summary>文字を出力するならtrue</summary>
        Public Property PrintsContent() As Boolean
            Get
                Return _printContentStatus <> PrintContentStatus.Unnecessary
            End Get
            Set(ByVal value As Boolean)
                If value Then
                    If _printContentStatus = PrintContentStatus.Unnecessary Then
                        _printContentStatus = PrintContentStatus.Overlap
                    End If
                Else
                    _printContentStatus = PrintContentStatus.Unnecessary
                End If
            End Set
        End Property
        ''' <summary>バーコード領域の外側に文字を出力するならtrue</summary>
        Public Property PrintsContentOutsideOfArea() As Boolean
            Get
                Return _printContentStatus = PrintContentStatus.Outside
            End Get
            Set(ByVal value As Boolean)
                If value Then
                    _printContentStatus = PrintContentStatus.Outside
                Else
                    _printContentStatus = PrintContentStatus.Overlap
                End If
            End Set
        End Property
#End Region
        ''' <summary>
        ''' 出力内容を符号化する
        ''' </summary>
        ''' <param name="content">内容</param>
        ''' <param name="quietZoneSize">クワイエットゾーンの幅(細バー比) ※10以上じゃないと無意味</param>
        ''' <returns>符号化した出力内容</returns>
        ''' <remarks></remarks>
        Private Shared Function EncodeBits(ByVal content As String, ByVal quietZoneSize As Integer) As String
            Const QUIET_SIGN As Char = "0"c
            Const ASTERISK As String = "100010111011101"
            Const START_CHARACTER As String = ASTERISK
            Const STOP_CHARACTER As String = ASTERISK
            Const GAP_SIGN As Char = "0"c

            Dim result As New StringBuilder()
            result.Append(QUIET_SIGN, quietZoneSize)
            result.Append(START_CHARACTER).Append(GAP_SIGN)
            For Each c As Char In content.ToUpper
                If Not BITS_BY_CHAR.ContainsKey(c) Then
                    Throw New ArgumentOutOfRangeException("content", String.Format("'{0}' はCODE39化できません.", c))
                End If
                result.Append(BITS_BY_CHAR(c)).Append(GAP_SIGN)
            Next
            result.Append(STOP_CHARACTER)
            result.Append(QUIET_SIGN, quietZoneSize)
            Return result.ToString
        End Function

        ''' <summary>
        ''' CODE39を作成する
        ''' </summary>
        ''' <param name="content">出力内容</param>
        ''' <returns>Bitmap</returns>
        ''' <remarks></remarks>
        Public Function Create(ByVal content As String) As Bitmap
            Dim encodedContent As String = EncodeBits(content, If(IgnoresQuietZone, 0, 15))
            Dim width As Integer = encodedContent.Length * NarrowBarWidth
            Dim contentHeight As Integer = If(PrintsContentOutsideOfArea, GetTextHeight, 0)
            Dim canvas As New Bitmap(width, BarHeight + contentHeight)

            Using aGraphics As Graphics = Graphics.FromImage(canvas)
                DrawBarcode(aGraphics, encodedContent, width)
                DrawContentIfNecessary(aGraphics, content, width)
            End Using

            Return canvas
        End Function

        ''' <summary>
        ''' CODE39をメタファイルで作成する
        ''' </summary>
        ''' <param name="content">出力内容</param>
        ''' <returns>metafile</returns>
        ''' <remarks>
        ''' 当メソッドで作成したメタファイルを保存する場合の注意点
        ''' Metafile#Save()メソッドで保存した場合、画質が大幅に劣化する(PNG形式で保存されるらしい MSDN情報)
        ''' </remarks>
        Public Function CreateAsMeta(ByVal content As String) As Metafile
            Dim encodedContent As String = EncodeBits(content, If(IgnoresQuietZone, 0, 15))
            Dim width As Integer = encodedContent.Length * NarrowBarWidth
            Dim contentHeight As Integer = If(PrintsContentOutsideOfArea, GetTextHeight, 0)
            Dim canvas As New Bitmap(width, BarHeight + contentHeight)
            Dim meta As Metafile

            Using canvasGraphics As Graphics = Graphics.FromImage(canvas)
                Dim hdc As IntPtr = canvasGraphics.GetHdc()
                meta = New Metafile(hdc, System.Drawing.Imaging.EmfType.EmfOnly)
                canvasGraphics.ReleaseHdc(hdc)
                Using emfGraphics As Graphics = Graphics.FromImage(meta)
                    DrawBarcode(emfGraphics, encodedContent, width)
                    DrawContentIfNecessary(emfGraphics, content, width)
                End Using
            End Using

            Return meta
        End Function

        Private Sub DrawBarcode(ByVal aGraphics As Graphics, ByVal encodedContent As String, ByVal width As Integer)
            Using aBarColor As New SolidBrush(Me.BarColor)
                Using bgColor As New SolidBrush(BackGroundColor)
                    aGraphics.FillRectangle(bgColor, New Rectangle(0, 0, width, BarHeight))

                    Dim offsetLeft As Integer = 0
                    For Each c As Char In encodedContent
                        offsetLeft += NarrowBarWidth
                        Dim aRectangle As New Rectangle(offsetLeft, 0, NarrowBarWidth, BarHeight)
                        aGraphics.FillRectangle(If(c = "0"c, bgColor, aBarColor), aRectangle)
                    Next
                End Using
            End Using
        End Sub

        Private Sub DrawContentIfNecessary(ByVal aGraphics As Graphics, ByVal content As String, ByVal width As Integer)
            If Not PrintsContent Then
                Return
            End If
            Dim contentHeight As Integer = If(PrintsContentOutsideOfArea, GetTextHeight, 0)
            Dim aFont As Font = New Font(FontName, FontSize)
            Dim text As String = String.Format(If(PrintsCheckAster, "*{0}*", "{0}"), content)
            Dim textLayerSize As SizeF = aGraphics.MeasureString(text, aFont)
            Dim textLayerSizeWidth As Integer = CInt(Math.Ceiling(textLayerSize.Width))

            Dim textLayerSizeHeight As Integer = CInt(Math.Ceiling(textLayerSize.Height))

            Dim aRectangle As RectangleF = New RectangleF(CSng((width - textLayerSizeWidth) / 2), BarHeight - textLayerSizeHeight + contentHeight, _
                                                          textLayerSizeWidth, textLayerSizeHeight)
            Using bgColor As New SolidBrush(BackGroundColor)
                aGraphics.FillRectangle(bgColor, aRectangle)
            End Using
            Using aFontColor As New SolidBrush(Me.FontColor)
                aGraphics.DrawString(text, aFont, aFontColor, aRectangle)
            End Using
        End Sub

        ''' <summary>
        ''' 保存する
        ''' </summary>
        ''' <param name="saveMetafile">保存するメタファイル</param>
        ''' <param name="filePath">ファイルパス</param>
        ''' <returns>ファイル名をもつメタファイル</returns>
        ''' <remarks>※Metafile#Save()だとPNG形式で保存されるらしい(MSDN情報)</remarks>
        Public Shared Function Save(ByVal saveMetafile As Metafile, ByVal filePath As String) As Metafile
            Return Save(saveMetafile, filePath, saveMetafile.Width)
        End Function

        ''' <summary>
        ''' 保存する
        ''' </summary>
        ''' <param name="saveMetafile">保存するメタファイル</param>
        ''' <param name="filePath">ファイルパス</param>
        ''' <param name="width">サイズ幅</param>
        ''' <returns>サイズ変更後のメタファイル</returns>
        ''' <remarks>※Metafile#Save()だとPNG形式で保存されるらしい(MSDN情報)</remarks>
        Public Shared Function Save(ByVal saveMetafile As Metafile, ByVal filePath As String, ByVal width As Integer) As Metafile
            Return Save(saveMetafile, filePath, width, saveMetafile.Height)
        End Function

        ''' <summary>
        ''' 保存する
        ''' </summary>
        ''' <param name="saveMetafile">保存するメタファイル</param>
        ''' <param name="filePath">ファイルパス</param>
        ''' <param name="width">サイズ幅</param>
        ''' <param name="height">サイズ高さ</param>
        ''' <returns>サイズ変更後のメタファイル</returns>
        ''' <remarks>※Metafile#Save()だとPNG形式で保存されるらしい(MSDN情報)</remarks>
        Public Shared Function Save(ByVal saveMetafile As Metafile, ByVal filePath As String, ByVal width As Integer, ByVal height As Integer) As Metafile
            Dim meta As Metafile
            Dim canvas As New Bitmap(width, height)
            Using bmpg As Graphics = Graphics.FromImage(canvas)
                Dim hdc As IntPtr = bmpg.GetHdc()
                meta = New Metafile(filePath, hdc, EmfType.EmfOnly)
                bmpg.ReleaseHdc(hdc)
                Using emfg As Graphics = Graphics.FromImage(meta)
                    emfg.DrawImage(saveMetafile, 0, 0, canvas.Width, canvas.Height)
                End Using
            End Using
            Return meta
        End Function

        Private Function GetTextHeight() As Integer
            Const content As String = "A"
            Using aGraphics As Graphics = Graphics.FromImage(New Bitmap(Me.FontSize, BarHeight))
                Dim aFont As Font = New Font(FontName, FontSize)
                Dim textLayerSize As SizeF = aGraphics.MeasureString(content, aFont)
                Return CInt(Math.Ceiling(textLayerSize.Height))
            End Using
        End Function

    End Class
End Namespace