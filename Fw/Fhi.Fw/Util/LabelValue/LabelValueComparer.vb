Namespace Util.LabelValue
    ''' <summary>
    ''' LabelValueVoのIComparer実装クラス
    ''' </summary>
    ''' <remarks>Nothingを考慮してソートする</remarks>
    Public Class LabelValueComparer : Implements IComparer(Of LabelValueVo)
        ''' <summary>
        ''' ラベルでソートするComparerを返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NewLabelSort() As LabelValueComparer
            Return New LabelValueComparer
        End Function
        ''' <summary>
        ''' 値でソートするComparerを返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NewValueSort() As LabelValueComparer
            Return New LabelValueComparer(False)
        End Function
        ''' <summary>
        ''' 値を数値としてソートするComparerを返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NewIntValueSort() As LabelValueComparer
            Return New LabelValueComparer(False, True)
        End Function

        Private isSortByLabel As Boolean
        Private isSortByInt As Boolean
        ''' <summary>
        ''' コンストラクタ(ラベルでソートする)
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me.New(True)
        End Sub
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="isSortByLabel">ラベルでソートする場合、true</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal isSortByLabel As Boolean)
            Me.New(isSortByLabel, False)
        End Sub
        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="isSortByLabel">ラベルでソートする場合、true</param>
        ''' <param name="isSortByInt">ソート対象を数値でソートする場合、true</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal isSortByLabel As Boolean, ByVal isSortByInt As Boolean)
            Me.isSortByLabel = isSortByLabel
            Me.isSortByInt = isSortByInt
        End Sub

        ''' <summary>
        ''' ラベルの比較結果を返す
        ''' </summary>
        ''' <param name="x">値x</param>
        ''' <param name="y">値y</param>
        ''' <returns>結果</returns>
        ''' <remarks></remarks>
        Private Function CompareLabel(ByVal x As LabelValueVo, ByVal y As LabelValueVo) As Integer
            If x.Label Is Nothing And y.Label IsNot Nothing Then
                Return 1
            ElseIf x.Label IsNot Nothing And y.Label Is Nothing Then
                Return -1
            ElseIf x.Label Is Nothing And y.Label Is Nothing Then
                Return 0
            End If
            If isSortByInt AndAlso IsNumeric(x.Label) AndAlso IsNumeric(y.Label) Then
                Return Convert.ToInt32(x.Label).CompareTo(Convert.ToInt32(y.Label))
            End If
            Return x.Label.CompareTo(y.Label)
        End Function
        ''' <summary>
        ''' 値の比較結果を返す
        ''' </summary>
        ''' <param name="x">値x</param>
        ''' <param name="y">値y</param>
        ''' <returns>結果</returns>
        ''' <remarks></remarks>
        Private Function CompareValue(ByVal x As LabelValueVo, ByVal y As LabelValueVo) As Integer
            If x.Value Is Nothing And y.Value IsNot Nothing Then
                Return 1
            ElseIf x.Value IsNot Nothing And y.Value Is Nothing Then
                Return -1
            ElseIf x.Value Is Nothing And y.Value Is Nothing Then
                Return 0
            End If
            If isSortByInt AndAlso IsNumeric(x.Value) AndAlso IsNumeric(y.Value) Then
                Return Convert.ToInt32(x.Value).CompareTo(Convert.ToInt32(y.Value))
            End If
            Return x.Value.CompareTo(y.Value)
        End Function
        Public Function Compare(ByVal x As LabelValueVo, ByVal y As LabelValueVo) As Integer _
                Implements IComparer(Of LabelValueVo).Compare
            If isSortByLabel Then
                Dim result As Integer = CompareLabel(x, y)
                If result <> 0 Then
                    Return result
                End If
                Return CompareValue(x, y)
            End If

            Dim result2 As Integer = CompareValue(x, y)
            If result2 <> 0 Then
                Return result2
            End If
            Return CompareLabel(x, y)
        End Function

    End Class
End Namespace