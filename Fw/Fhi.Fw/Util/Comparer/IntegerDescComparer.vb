Namespace Util.Comparer

    ''' <summary>
    ''' 数値(Integer)の降順ソート
    ''' </summary>
    ''' <remarks></remarks>
    Public Class IntegerDescComparer : Implements IComparer(Of Integer)
        Public Function Compare(ByVal x As Integer, ByVal y As Integer) As Integer Implements IComparer(Of Integer).Compare
            Return y.CompareTo(x)
        End Function
    End Class

End Namespace
