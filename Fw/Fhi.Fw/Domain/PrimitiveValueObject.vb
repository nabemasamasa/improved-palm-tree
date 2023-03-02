Imports System.Runtime.CompilerServices

Namespace Domain
    ''' <summary>
    ''' 基本となる値を保持する値オブジェクト
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <remarks></remarks>
    <DebuggerDisplay("{Value} {GetClassName(),nq}")>
    Public MustInherit Class PrimitiveValueObject(Of T) : Inherits ValueObject : Implements PrimitiveValueObject, IComparable(Of Object)

        Private ReadOnly _value As T
        Private ReadOnly behavior As PrimitiveValueObject.IBehavior

        ''' <summary>値</summary>
        Protected ReadOnly Property Value() As T
            Get
                Return _value
            End Get
        End Property

        ''' <summary>値</summary>
        Private ReadOnly Property PrimitiveValueObject_Value() As Object Implements PrimitiveValueObject.Value
            Get
                Return _value
            End Get
        End Property

        Public Sub New(value As T)
            Me._value = value
            Me.behavior = DirectCast(Me, PrimitiveValueObject).That(GetType(T))
        End Sub

        Protected Overrides Function GetAtomicValues() As IEnumerable(Of Object)
            Return New Object() {_value}
        End Function

        Private Function GetClassName() As String
            ' 未使用表示だけど、DebuggerDisplayで使用中
            Return "{" & Me.GetType.Name & "}"
        End Function

        Public Overrides Function ToString() As String
            Return StringUtil.ToString(_value)
        End Function

        Private Function IComparable_CompareTo(ByVal other As Object) As Integer Implements IComparable(Of Object).CompareTo
            Return Me.CompareTo(other)
        End Function

        Public Overridable Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo
            If obj Is Nothing Then
                Return EzUtil.CompareObjectNullsLast(Me, obj).Value
            End If
            If Not Me.GetType Is obj.GetType Then
                Return Me.GetType.FullName.CompareTo(obj.GetType.FullName)
            End If
            Dim value1 As Object = Me.Value
            Dim value2 As Object = DirectCast(obj, PrimitiveValueObject).Value
            Dim resultIfContainsNull As Integer? = EzUtil.CompareObjectNullsLast(value2, value1)
            If resultIfContainsNull.HasValue Then
                Return resultIfContainsNull.Value
            End If
            Return behavior.Compare(value1, value2)
        End Function

    End Class
    ''' <summary>
    ''' 基本となるObject値を保持する値オブジェクト
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface PrimitiveValueObject : Inherits IComparable
        Friend Interface IBehavior
            ''' <summary>
            ''' 比較する
            ''' </summary>
            ''' <param name="x">比較値 x</param>
            ''' <param name="y">比較値 y</param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Compare(x As Object, y As Object) As Integer
            ''' <summary>
            ''' 値が無い、もしくは初期値か？を返す
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function IsEmpty(val As Object) As Boolean
        End Interface
        ''' <summary>値</summary>
        ReadOnly Property Value As Object
    End Interface
    Public Module PrimitiveValueObjectExtension
#Region "Nested classes..."
        Private Class StringImpl : Implements PrimitiveValueObject.IBehavior
            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements PrimitiveValueObject.IBehavior.Compare
                Return String.Compare(DirectCast(x, String), DirectCast(y, String))
            End Function
            Public Function IsEmpty(ByVal val As Object) As Boolean Implements PrimitiveValueObject.IBehavior.IsEmpty
                Return StringUtil.IsEmpty(val)
            End Function
        End Class
        Private Class IntegerImpl : Implements PrimitiveValueObject.IBehavior
            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements PrimitiveValueObject.IBehavior.Compare
                Return DirectCast(x, Integer).CompareTo(DirectCast(y, Integer))
            End Function
            Public Function IsEmpty(ByVal val As Object) As Boolean Implements PrimitiveValueObject.IBehavior.IsEmpty
                Return val Is Nothing
            End Function
        End Class
        Private Class DecimalImpl : Implements PrimitiveValueObject.IBehavior
            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements PrimitiveValueObject.IBehavior.Compare
                Return DirectCast(x, Decimal).CompareTo(DirectCast(y, Decimal))
            End Function
            Public Function IsEmpty(ByVal val As Object) As Boolean Implements PrimitiveValueObject.IBehavior.IsEmpty
                Return val Is Nothing
            End Function
        End Class
        Private Class DateTimeImpl : Implements PrimitiveValueObject.IBehavior
            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements PrimitiveValueObject.IBehavior.Compare
                Return DirectCast(x, DateTime).CompareTo(DirectCast(y, DateTime))
            End Function
            Public Function IsEmpty(ByVal val As Object) As Boolean Implements PrimitiveValueObject.IBehavior.IsEmpty
                Return val Is Nothing
            End Function
        End Class
        Private Class LongImpl : Implements PrimitiveValueObject.IBehavior
            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements PrimitiveValueObject.IBehavior.Compare
                Return DirectCast(x, Long).CompareTo(DirectCast(y, Long))
            End Function
            Public Function IsEmpty(ByVal val As Object) As Boolean Implements PrimitiveValueObject.IBehavior.IsEmpty
                Return val Is Nothing
            End Function
        End Class
        Private Class ByteImpl : Implements PrimitiveValueObject.IBehavior
            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements PrimitiveValueObject.IBehavior.Compare
                Return DirectCast(x, Byte).CompareTo(DirectCast(y, Byte))
            End Function
            Public Function IsEmpty(ByVal val As Object) As Boolean Implements PrimitiveValueObject.IBehavior.IsEmpty
                Return val Is Nothing
            End Function
        End Class
        Private Class ShortImpl : Implements PrimitiveValueObject.IBehavior
            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements PrimitiveValueObject.IBehavior.Compare
                Return DirectCast(x, Short).CompareTo(DirectCast(y, Short))
            End Function
            Public Function IsEmpty(ByVal val As Object) As Boolean Implements PrimitiveValueObject.IBehavior.IsEmpty
                Return val Is Nothing
            End Function
        End Class
        Private Class SingleImpl : Implements PrimitiveValueObject.IBehavior
            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements PrimitiveValueObject.IBehavior.Compare
                Return DirectCast(x, Single).CompareTo(DirectCast(y, Single))
            End Function
            Public Function IsEmpty(ByVal val As Object) As Boolean Implements PrimitiveValueObject.IBehavior.IsEmpty
                Return val Is Nothing
            End Function
        End Class
        Private Class DoubleImpl : Implements PrimitiveValueObject.IBehavior
            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements PrimitiveValueObject.IBehavior.Compare
                Return DirectCast(x, Double).CompareTo(DirectCast(y, Double))
            End Function
            Public Function IsEmpty(ByVal val As Object) As Boolean Implements PrimitiveValueObject.IBehavior.IsEmpty
                Return val Is Nothing
            End Function
        End Class
        Private Class UnknownTypeImpl : Implements PrimitiveValueObject.IBehavior
            Private ReadOnly originType As Type
            Friend Sub New(originType As Type)
                Me.originType = originType
            End Sub
            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements PrimitiveValueObject.IBehavior.Compare
                Throw New NotImplementedException(originType.FullName.ToString & "同士は未対応. 個別に実装して")
            End Function
            Public Function IsEmpty(ByVal val As Object) As Boolean Implements PrimitiveValueObject.IBehavior.IsEmpty
                If Not originType.IsPrimitive Then
                    Return val Is Nothing
                End If
                Throw New NotImplementedException(originType.FullName.ToString & "は未対応. 個別に実装して")
            End Function
        End Class
#End Region
        Private behaviorByType As Dictionary(Of Type, PrimitiveValueObject.IBehavior)
        <Extension()>
        Private Function That(pvo As PrimitiveValueObject) As PrimitiveValueObject.IBehavior
            If Not pvo.HasValue Then
                Throw New ArgumentException("値がNothingなので引数必要")
            End If
            Return pvo.That(pvo.Value.GetType)
        End Function
        ''' <summary>
        ''' 振る舞いを返す
        ''' </summary>
        ''' <param name="pvo">値オブジェクト</param>
        ''' <param name="rawType">Raw値の型</param>
        ''' <returns>振る舞い</returns>
        ''' <remarks></remarks>
        <Extension()>
        Friend Function That(pvo As PrimitiveValueObject, rawType As Type) As PrimitiveValueObject.IBehavior
            If behaviorByType Is Nothing Then
                behaviorByType = New Dictionary(Of Type, PrimitiveValueObject.IBehavior) _
                    From {{GetType(String), New StringImpl}, {GetType(Integer), New IntegerImpl}, {GetType(Decimal), New DecimalImpl},
                          {GetType(DateTime), New DateTimeImpl}, {GetType(Long), New LongImpl}, {GetType(Byte), New ByteImpl},
                          {GetType(Short), New ShortImpl}, {GetType(Single), New SingleImpl}, {GetType(Double), New DoubleImpl}}
            End If
            Dim originType As Type = TypeUtil.GetTypeIfNullable(rawType)
            If Not behaviorByType.ContainsKey(originType) Then
                Return New UnknownTypeImpl(originType)
            End If
            Return behaviorByType(originType)
        End Function
        ''' <summary>
        ''' 値があるか？を返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function HasValue(pvo As PrimitiveValueObject) As Boolean
            Return pvo IsNot Nothing AndAlso pvo.Value IsNot Nothing
        End Function
        ''' <summary>
        ''' 値が無い、もしくは初期値か？を返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function IsEmpty(pvo As PrimitiveValueObject) As Boolean
            If Not pvo.HasValue Then
                Return True
            End If
            Return pvo.That.IsEmpty(pvo.Value)
        End Function
        ''' <summary>
        ''' 値があるか？を返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function IsNotEmpty(pvo As PrimitiveValueObject) As Boolean
            Return Not pvo.IsEmpty()
        End Function
    End Module
End Namespace