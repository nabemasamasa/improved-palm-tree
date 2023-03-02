Imports Fhi.Fw.Util

Namespace App.Xls
    ''' <summary>
    ''' Excelセル座標のみを表すクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class XlAxis : Inherits AbstractImmutablePair(Of Integer, Integer)

        ''' <summary>行index 1スタート</summary>
        Public ReadOnly Property Row As Integer
            Get
                Return PairA
            End Get
        End Property

        ''' <summary>列index 1スタート</summary>
        Public ReadOnly Property Column As Integer
            Get
                Return PairB
            End Get
        End Property

        Public Sub New(ByVal row As Integer, ByVal column As Integer)
            MyBase.New(PairA:=row, PairB:=column)
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("RC=({0},{1})", Row, Column)
        End Function

    End Class
End Namespace