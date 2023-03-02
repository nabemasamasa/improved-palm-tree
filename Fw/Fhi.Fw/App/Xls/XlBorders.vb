Namespace App.Xls
    ''' <summary>
    ''' 罫線情報群 (VBAのBordersオブジェクト模倣)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class XlBorders : Implements Border
        ''' <summary>実線</summary>
        Public Shared ReadOnly Continuous As Border = New ReadonlyBorder(XlLineStyle.xlContinuous, XlBorderWeight.xlThin)
        ''' <summary>二重線</summary>
        Public Shared ReadOnly [Double] As Border = New ReadonlyBorder(XlLineStyle.xlDouble, XlBorderWeight.xlThick)
        ''' <summary>点線</summary>
        Public Shared ReadOnly Dot As Border = New ReadonlyBorder(XlLineStyle.xlDot, XlBorderWeight.xlThin)
        ''' <summary>破線</summary>
        Public Shared ReadOnly Dash As Border = New ReadonlyBorder(XlLineStyle.xlDash, XlBorderWeight.xlThin)

        ' 定型として必要なら追加してほしい

#Region "Nested classes..."
        ''' <summary>
        ''' 罫線情報 (VBAのBorderオブジェクト模倣)
        ''' </summary>
        ''' <remarks></remarks>
        Public Interface Border
            ''' <summary>罫線の種類</summary>
            Property LineStyle As XlLineStyle?
            ''' <summary>罫線の太さ</summary>
            Property Weight As XlBorderWeight?
        End Interface
        Private Class ReadWriteBorder : Implements Border
            Public Property LineStyle As XlLineStyle? Implements Border.LineStyle
            Public Property Weight As XlBorderWeight? Implements Border.Weight
        End Class
        Private Class ReadonlyBorder : Implements Border
            Private ReadOnly _LineStyle As XlLineStyle?
            Private ReadOnly _Weight As XlBorderWeight?
            Public Sub New(ByVal lineStyle As XlLineStyle, ByVal weight As XlBorderWeight)
                _LineStyle = lineStyle
                _Weight = weight
            End Sub
            Private Sub RaiseInvalidProgramException()
                Throw New InvalidProgramException("読み取り専用")
            End Sub

            Public Property LineStyle As XlLineStyle? Implements Border.LineStyle
                Get
                    Return _LineStyle
                End Get
                Set(value As XlLineStyle?)
                    RaiseInvalidProgramException()
                End Set
            End Property

            Public Property Weight As XlBorderWeight? Implements Border.Weight
                Get
                    Return _Weight
                End Get
                Set(value As XlBorderWeight?)
                    RaiseInvalidProgramException()
                End Set
            End Property
        End Class
#End Region

        ''' <summary>セル範囲の各セルの左上隅から右下隅への罫線</summary>
        Public DiagonalDown As Border = New ReadWriteBorder
        ''' <summary>セル範囲の各セルの左下隅から右上隅への罫線</summary>
        Public DiagonalUp As Border = New ReadWriteBorder
        ''' <summary>セル範囲の左側の罫線</summary>
        Public EdgeLeft As Border = New ReadWriteBorder
        ''' <summary>セル範囲の上側の罫線</summary>
        Public EdgeTop As Border = New ReadWriteBorder
        ''' <summary>セル範囲の下側の罫線</summary>
        Public EdgeBottom As Border = New ReadWriteBorder
        ''' <summary>セル範囲の右側の罫線</summary>
        Public EdgeRight As Border = New ReadWriteBorder
        ''' <summary>セル範囲の外枠を除く、すべてのセルの垂直方向の罫線</summary>
        Public InsideVertical As Border = New ReadWriteBorder
        ''' <summary>セル範囲の外枠を除く、すべてのセルの水平方向の罫線</summary>
        Public InsideHorizontal As Border = New ReadWriteBorder

        ''' <summary>罫線の種類</summary>
        Public Property LineStyle As XlLineStyle? Implements Border.LineStyle
        ''' <summary>罫線の太さ</summary>
        Public Property Weight As XlBorderWeight? Implements Border.Weight

    End Class

End Namespace