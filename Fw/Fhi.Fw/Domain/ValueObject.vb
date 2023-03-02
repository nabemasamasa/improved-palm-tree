Imports System.Reflection

Namespace Domain
    ''' <summary>
    ''' 値オブジェクトの基底クラス
    ''' </summary>
    ''' <remarks>当クラス実装のときは、Immutableクラスにしなければいけない</remarks>
    ''' <see>https://docs.microsoft.com/ja-jp/dotnet/standard/microservices-architecture/microservice-ddd-cqrs-patterns/implement-value-objects</see>
    Public MustInherit Class ValueObject

        ''' <summary>
        ''' 同値判定に含めるメンバー変数値を列挙して返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected MustOverride Function GetAtomicValues() As IEnumerable(Of Object)

        Protected Shared Function EqualsOperator(left As ValueObject, right As ValueObject) As Boolean
            If ReferenceEquals(left, Nothing) Xor ReferenceEquals(right, Nothing) Then
                Return False
            End If
            Return ReferenceEquals(left, Nothing) OrElse left.Equals(right)
        End Function

        Protected Shared Function NotEqualsOperator(left As ValueObject, right As ValueObject) As Boolean
            Return Not EqualsOperator(left, right)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj Is Nothing OrElse Not (obj.GetType Is Me.GetType) Then
                Return False
            End If
            Dim other As ValueObject = DirectCast(obj, ValueObject)
            Dim thisValues As IEnumerator(Of Object) = GetAtomicValues.GetEnumerator
            Dim otherValues As IEnumerator(Of Object) = other.GetAtomicValues.GetEnumerator
            While thisValues.MoveNext AndAlso otherValues.MoveNext
                If ReferenceEquals(thisValues.Current, Nothing) Xor ReferenceEquals(otherValues.Current, Nothing) Then
                    Return False
                End If
                If thisValues.Current IsNot Nothing AndAlso Not thisValues.Current.Equals(otherValues.Current) Then
                    Return False
                End If
            End While
            Return Not thisValues.MoveNext AndAlso Not otherValues.MoveNext
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return EzUtil.MakeKey(GetAtomicValues.Select(
                                  Function(x) If(x IsNot Nothing, EzUtil.MakeKey({x.GetHashCode, x.ToString()}), "0")).ToArray()
                              ).GetHashCode()
        End Function

        ''' <summary>
        ''' ValueObjectのPublicプロパティを最も初期化できるコンストラクタを返す
        ''' </summary>
        ''' <param name="aType"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Shared Function DetectConstructor(ByVal aType As Type) As ConstructorInfo
            If Not TypeUtil.IsTypeValueObjectOrSubClass(aType) Then
                Throw New ArgumentException("aTypeはValueObjectじゃない." & aType.FullName)
            End If

            Dim fields As FieldInfo() = aType.GetFields(BindingFlags.Public Or BindingFlags.Instance)
            Dim propertyInfos As PropertyInfo() = aType.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
            Dim memberTypes As List(Of Type) = fields.Select(Function(f) f.FieldType).Concat(propertyInfos.Select(Function(p) p.PropertyType)).ToList
            Return VoUtil.DetectConstructorInfo(aType, memberTypes)
        End Function

    End Class
End Namespace