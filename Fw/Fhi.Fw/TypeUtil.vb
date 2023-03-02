Imports Fhi.Fw.Domain

''' <summary>
''' 型に関するユーティリティクラス
''' </summary>
''' <remarks></remarks>
Public Class TypeUtil

    ''' <summary>IEnumerable型</summary>
    Public Shared ReadOnly TypeIEnumerable As Type = GetType(IEnumerable)
    ''' <summary>ICollection型</summary>
    Public Shared ReadOnly TypeICollection As Type = GetType(ICollection)
    ''' <summary>String型</summary>
    Public Shared ReadOnly TypeString As Type = GetType(String)
    ''' <summary>DateTime型</summary>
    Public Shared ReadOnly TypeDateTime As Type = GetType(DateTime)
    ''' <summary>Decimal型</summary>
    Public Shared ReadOnly TypeDecimal As Type = GetType(Decimal)
    ''' <summary>Integer型</summary>
    Public Shared ReadOnly TypeInteger As Type = GetType(Int32)
    ''' <summary>Long型</summary>
    Public Shared ReadOnly TypeLong As Type = GetType(Int64)
    ''' <summary>Single型</summary>
    Public Shared ReadOnly TypeSingle As Type = GetType(Single)
    ''' <summary>Double型</summary>
    Public Shared ReadOnly TypeDouble As Type = GetType(Double)
    ''' <summary>Boolean型</summary>
    Public Shared ReadOnly TypeBoolean As Type = GetType(Boolean)
    ''' <summary>Object型</summary>
    Public Shared ReadOnly TypeObject As Type = GetType(Object)
    ''' <summary>ValueObject型</summary>
    Public Shared ReadOnly TypeValueObject As Type = GetType(ValueObject)

    ''' <summary>
    ''' 匿名型かを返す
    ''' </summary>
    ''' <param name="obj">判定するインスタンス</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsTypeAnonymous(ByVal obj As Object) As Boolean
        Return IsTypeAnonymous(obj.GetType)
    End Function

    ''' <summary>
    ''' 匿名型かを返す
    ''' </summary>
    ''' <param name="aType">判定するType</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsTypeAnonymous(ByVal aType As Type) As Boolean
        ' "VB":VisualBasicコンパイラによって生成されている型を表す
        ' "$" :VisualBasicで許可されていない記号. ユーザー定義の型と競合しないようになっている
        Return aType.Name.StartsWith("VB$AnonymousType")
    End Function

    ''' <summary>
    ''' Nullable(Of T) かを返す
    ''' </summary>
    ''' <param name="aType">判定するType</param>
    ''' <returns>結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsTypeNullable(ByVal aType As Type) As Boolean
        Return aType.Name.StartsWith("Nullable`")
    End Function

    ''' <summary>
    ''' 型を返す. Nullable(Of T) なら T の型を返す.
    ''' </summary>
    ''' <param name="aType">判定する Type</param>
    ''' <returns>型 (Nullable以外)</returns>
    ''' <remarks></remarks>
    Public Shared Function GetTypeIfNullable(ByVal aType As Type) As Type
        If IsTypeNullable(aType) Then
            Return aType.GetGenericArguments(0)
        End If
        Return aType
    End Function

    ''' <summary>
    ''' 不変オブジェクトかを返す（NullableもTrue）
    ''' </summary>
    ''' <param name="propertyType">型</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsTypeImmutable(ByVal propertyType As Type) As Boolean
        Return IsTypeNullable(propertyType) _
                OrElse propertyType.IsValueType _
                OrElse propertyType Is TypeUtil.TypeString _
                OrElse IsTypeValueObjectOrSubClass(propertyType)
    End Function

    ''' <summary>
    ''' ValueObject、もしくはそのサブクラスか？を返す
    ''' </summary>
    ''' <param name="aType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsTypeValueObjectOrSubClass(ByVal aType As Type) As Boolean
        Return TypeUtil.TypeValueObject.IsAssignableFrom(aType)
    End Function

    ''' <summary>
    ''' AはGenericなB、またはそのサブクラスか？を返す
    ''' </summary>
    ''' <param name="aType">型A</param>
    ''' <param name="aGenericType">Genericな型B</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsTypeGenericOrSubClass(ByVal aType As Type, ByVal aGenericType As Type) As Boolean
        Return (aType.IsGenericType AndAlso (aType.GetGenericTypeDefinition Is aGenericType OrElse aGenericType.IsAssignableFrom(aType))) _
            OrElse (aType.BaseType IsNot Nothing AndAlso IsTypeGenericOrSubClass(aType.BaseType, aGenericType))
    End Function

    ''' <summary>
    ''' 配列、またはコレクションか？を返す
    ''' </summary>
    ''' <param name="arrayOrCollection">配列、またはコレクション</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsArrayOrCollection(ByVal arrayOrCollection As Object) As Boolean
        Return TypeOf arrayOrCollection Is ICollection
    End Function

    ''' <summary>
    ''' 配列型かを返す
    ''' </summary>
    ''' <param name="aType">型</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsTypeArray(ByVal aType As Type) As Boolean
        Return aType.IsArray
    End Function

    ''' <summary>
    ''' コレクション型（含む配列）かを返す
    ''' </summary>
    ''' <param name="aType">型</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsTypeCollection(ByVal aType As Type) As Boolean
        Return TypeUtil.TypeICollection.IsAssignableFrom(aType) OrElse IsTypeGenericOrSubClass(aType, GetType(ICollection(Of )))
    End Function

    ''' <summary>
    ''' コレクションの要素型を返す
    ''' </summary>
    ''' <param name="collectionValue">コレクション値</param>
    ''' <returns>要素型</returns>
    ''' <remarks></remarks>
    Public Shared Function DetectElementType(ByVal collectionValue As Object) As Type
        Return DetectElementType(collectionValue.GetType)
    End Function

    ''' <summary>
    ''' コレクションの要素型を返す
    ''' </summary>
    ''' <param name="collectionType">コレクション型</param>
    ''' <returns>要素型</returns>
    ''' <remarks></remarks>
    Public Shared Function DetectElementType(ByVal collectionType As Type) As Type
        AssertImplementsICollection(collectionType)
        If collectionType.IsGenericType Then
            Return collectionType.GetGenericArguments(0)
        End If
        If TypeUtil.IsTypeArray(collectionType) Then
            Return collectionType.GetElementType
        End If
        Throw New InvalidOperationException(String.Format("引数 value は {0} 型。未対応。", collectionType.Name))
    End Function

    ''' <summary>
    ''' ICollection型を実装している事を保証する
    ''' </summary>
    ''' <param name="aType">型</param>
    ''' <remarks></remarks>
    Public Shared Sub AssertImplementsICollection(ByVal aType As Type)
        If Not IsTypeCollection(aType) Then
            Throw New ArgumentException(String.Format("引数 value は {0} 型。ICollection型を実装していません。", aType.Name))
        End If
    End Sub
End Class