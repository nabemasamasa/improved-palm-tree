Imports Fhi.Fw.Domain

''' <summary>
''' 比較を行うユーティリティ
''' </summary>
''' <remarks></remarks>
Public Class CompareUtil
    ''' <summary>比較結果</summary>
    Public Enum CompareResult
        ''' <summary>両辺が等しい(＝)</summary>
        Equal
        ''' <summary>右辺が左辺より大きい(＜)</summary>
        GreaterThan
        ''' <summary>右辺が左辺より小さい(＞)</summary>
        LessThan
    End Enum

    ''' <summary>
    ''' 比較する
    ''' </summary>
    ''' <param name="leftSide">左辺</param>
    ''' <param name="rightSide">右辺</param>
    ''' <param name="ignorePvo">PVOかを無視（内部値同士で比較）するか？</param>
    ''' <returns>比較結果</returns>
    ''' <remarks></remarks>
    Public Shared Function Compare(leftSide As Object, rightSide As Object, Optional ignorePvo As Boolean = False) As CompareResult
        If leftSide Is Nothing OrElse rightSide Is Nothing Then
            Throw New ArgumentNullException()
        End If
        If ignorePvo Then
            Dim lsValue As Object = If(TypeOf leftSide Is PrimitiveValueObject, DirectCast(leftSide, PrimitiveValueObject).Value, leftSide)
            Dim rsValue As Object = If(TypeOf rightSide Is PrimitiveValueObject, DirectCast(rightSide, PrimitiveValueObject).Value, rightSide)
            Return Compare(lsValue, rsValue)
        End If
        If leftSide.GetType() IsNot rightSide.GetType() Then
            Throw New ArgumentException("左辺と右辺の型は同じでないとダメ")
        End If

        If TypeOf leftSide Is PrimitiveValueObject Then
            Dim lsValue As Object = DirectCast(leftSide, PrimitiveValueObject).Value
            Dim rsValue As Object = DirectCast(rightSide, PrimitiveValueObject).Value
            Return Compare(lsValue, rsValue)
        ElseIf TypeOf leftSide Is String Then
            Return CompareString(leftSide.ToString(), rightSide.ToString())
        ElseIf TypeOf leftSide Is Integer Then
            Return CompareInteger(CInt(leftSide), CInt(rightSide))
        ElseIf TypeOf leftSide Is Decimal Then
            Return CompareDecimal(CDec(leftSide), CDec(rightSide))
        ElseIf TypeOf leftSide Is Single Then
            Return CompareSingle(CSng(leftSide), CSng(rightSide))
        ElseIf TypeOf leftSide Is Double Then
            Return CompareDouble(CDbl(leftSide), CDbl(rightSide))
        ElseIf TypeOf leftSide Is Date Then
            Return CompareDate(CDate(leftSide), CDate(rightSide))
        End If
        Throw New ArgumentException(String.Format("{0}型は対応していない", leftSide.GetType()))
    End Function

    Private Shared Function CompareString(leftSide As String, rightSide As String) As CompareResult
        Dim result As Integer = String.Compare(leftSide, rightSide)
        If result = 0 Then
            Return CompareResult.Equal
        ElseIf result < 0 Then
            Return CompareResult.GreaterThan
        End If
        Return CompareResult.LessThan
    End Function

    Private Shared Function CompareInteger(leftSide As Integer, rightSide As Integer) As CompareResult
        If leftSide = rightSide Then
            Return CompareResult.Equal
        ElseIf leftSide < rightSide Then
            Return CompareResult.GreaterThan
        End If
        Return CompareResult.LessThan
    End Function

    Private Shared Function CompareDecimal(leftSide As Decimal, rightSide As Decimal) As CompareResult
        If leftSide = rightSide Then
            Return CompareResult.Equal
        ElseIf leftSide < rightSide Then
            Return CompareResult.GreaterThan
        End If
        Return CompareResult.LessThan
    End Function

    Private Shared Function CompareSingle(leftSide As Single, rightSide As Single) As CompareResult
        If leftSide = rightSide Then
            Return CompareResult.Equal
        ElseIf leftSide < rightSide Then
            Return CompareResult.GreaterThan
        End If
        Return CompareResult.LessThan
    End Function

    Private Shared Function CompareDouble(leftSide As Double, rightSide As Double) As CompareResult
        If leftSide = rightSide Then
            Return CompareResult.Equal
        ElseIf leftSide < rightSide Then
            Return CompareResult.GreaterThan
        End If
        Return CompareResult.LessThan
    End Function

    Private Shared Function CompareDate(leftSide As DateTime, rightSide As DateTime) As CompareResult
        Dim result As Integer = Date.Compare(leftSide, rightSide)
        If result = 0 Then
            Return CompareResult.Equal
        ElseIf result < 0 Then
            Return CompareResult.GreaterThan
        End If
        Return CompareResult.LessThan
    End Function

    ''' <summary>
    ''' 左辺と右辺が等しいか？
    ''' leftSide＝rightSide
    ''' </summary>
    ''' <param name="leftSide">左辺</param>
    ''' <param name="rightSide">右辺</param>
    ''' <param name="ignorePvo">PVOかを無視（内部値同士で比較）するか？</param>
    ''' <returns>等しければTrue</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEqual(leftSide As Object, rightSide As Object, Optional ignorePvo As Boolean = False) As Boolean
        Return Compare(leftSide, rightSide, ignorePvo) = CompareResult.Equal
    End Function

    ''' <summary>
    ''' 左辺より右辺が大きいか？
    ''' leftSide＜rightSide
    ''' </summary>
    ''' <param name="leftSide">左辺</param>
    ''' <param name="rightSide">右辺</param>
    ''' <param name="ignorePvo">PVOかを無視（内部値同士で比較）するか？</param>
    ''' <returns>右辺のほうが大きければTrue</returns>
    ''' <remarks></remarks>
    Public Shared Function IsGreaterThan(leftSide As Object, rightSide As Object, Optional ignorePvo As Boolean = False) As Boolean
        Return Compare(leftSide, rightSide, ignorePvo) = CompareResult.GreaterThan
    End Function

    ''' <summary>
    ''' 左辺より右辺が小さいか？
    ''' leftSide＞rightSide
    ''' </summary>
    ''' <param name="leftSide">左辺</param>
    ''' <param name="rightSide">右辺</param>
    ''' <param name="ignorePvo">PVOかを無視（内部値同士で比較）するか？</param>
    ''' <returns>右辺のほうが小さければTrue</returns>
    ''' <remarks></remarks>
    Public Shared Function IsLessThan(leftSide As Object, rightSide As Object, Optional ignorePvo As Boolean = False) As Boolean
        Return Compare(leftSide, rightSide, ignorePvo) = CompareResult.LessThan
    End Function

    ''' <summary>
    ''' 右辺が左辺以上か？
    ''' leftSide≦rightSide
    ''' </summary>
    ''' <param name="leftSide">左辺</param>
    ''' <param name="rightSide">右辺</param>
    ''' <param name="ignorePvo">PVOかを無視（内部値同士で比較）するか？</param>
    ''' <returns>左辺以上ならTrue</returns>
    ''' <remarks></remarks>
    Public Shared Function IsGreaterEqual(leftSide As Object, rightSide As Object, Optional ignorePvo As Boolean = False) As Boolean
        Dim result As CompareResult = Compare(leftSide, rightSide, ignorePvo)
        Return result = CompareResult.Equal OrElse result = CompareResult.GreaterThan
    End Function

    ''' <summary>
    ''' 右辺が左辺以下か？
    ''' leftSide≧rightSide
    ''' </summary>
    ''' <param name="leftSide">左辺</param>
    ''' <param name="rightSide">右辺</param>
    ''' <param name="ignorePvo">PVOかを無視（内部値同士で比較）するか？</param>
    ''' <returns>左辺以下ならTrue</returns>
    ''' <remarks></remarks>
    Public Shared Function IsLessEqual(leftSide As Object, rightSide As Object, Optional ignorePvo As Boolean = False) As Boolean
        Dim result As CompareResult = Compare(leftSide, rightSide, ignorePvo)
        Return result = CompareResult.Equal OrElse result = CompareResult.LessThan
    End Function

End Class
