Imports System.Reflection
Imports System.Linq.Expressions
Imports Fhi.Fw.Domain
Imports Fhi.Fw.Util

Public Class VoUtil

    ''' <summary>自身を参照するプロパティ名</summary>
    Public Const PROPERTY_NAME_SELF_VALUE As String = "Value"

    ''' <summary>
    ''' 同名プロパティ同士の値をコピーする
    ''' </summary>
    ''' <param name="src">コピー元のVo</param>
    ''' <param name="dest">コピー先のVo</param>
    ''' <remarks>同名プロパティは</remarks>
    Public Shared Sub CopyProperties(ByVal src As Object, ByVal dest As Object)
        CopyPropertiesWithIgnoreProperties(Of Object)(src, dest, Nothing)
    End Sub

    ''' <summary>
    ''' 無視したいプロパティ以外の同名プロパティ同士の値をコピーする
    ''' </summary>
    ''' <param name="src">コピー元のVo</param>
    ''' <param name="dest">コピー先のVo</param>
    ''' <remarks>同名プロパティは</remarks>
    Public Shared Sub CopyPropertiesWithIgnoreProperties(Of T)(src As Object, dest As Object, configure As PropertySpecifierConfigure(Of T))
        Dim ignoreProperties As String() = If(configure IsNot Nothing, New PropertySpecifier(Of T)(configure).Results, New String() {})
        Dim destSetterProperties As New Dictionary(Of String, PropertyInfo)
        For Each aProperty As PropertyInfo In dest.GetType.GetProperties
            If aProperty.GetSetMethod Is Nothing Then
                Continue For
            End If
            If aProperty.GetSetMethod.GetParameters.Length = 0 Then
                Continue For
            End If
            If ignoreProperties.Contains(aProperty.Name) Then
                Continue For
            End If
            destSetterProperties.Add(aProperty.Name, aProperty)
        Next
        For Each aProperty As PropertyInfo In src.GetType.GetProperties
            If aProperty.GetGetMethod Is Nothing Then
                Continue For
            End If
            If Not destSetterProperties.ContainsKey(aProperty.Name) Then
                Continue For
            End If

            Dim value As Object = aProperty.GetValue(src, Nothing)
            Dim destProperty As PropertyInfo = destSetterProperties(aProperty.Name)
            destProperty.SetValue(dest, ResolveValue(value, destProperty.PropertyType, True), Nothing)
        Next
    End Sub

    ''' <summary>
    ''' 新しいインスタンスを生成する
    ''' </summary>
    ''' <returns>新しいインスタンス</returns>
    ''' <remarks>ただし引数無しコンストラクタを持つクラスに限る</remarks>
    Public Shared Function NewInstance(ByVal aType As Type) As Object
        Return Activator.CreateInstance(aType)
    End Function

    ''' <summary>
    ''' 新しいインスタンスを生成する
    ''' </summary>
    ''' <typeparam name="T">生成するインスタンスの型</typeparam>
    ''' <param name="ignoreScope">非Publicコンストラクタでもインスタンス生成する場合、true</param>
    ''' <returns>新しいインスタンス</returns>
    ''' <remarks>ただし引数無しコンストラクタを持つクラスに限る</remarks>
    Public Shared Function NewInstance(Of T)(Optional ignoreScope As Boolean = False) As T
        Return NewInstance(Of T)(Nothing, ignoreScope)
    End Function

    ''' <summary>
    ''' 新しいインスタンスを生成する
    ''' </summary>
    ''' <typeparam name="T">生成するインスタンスの型</typeparam>
    ''' <param name="propertiesObj">コピーさせたいプロパティ値をもつobject</param>
    ''' <param name="ignoreScope">非Publicコンストラクタでもインスタンス生成する場合、true</param>
    ''' <returns>新しいインスタンス</returns>
    ''' <remarks>ただし引数無しコンストラクタを持つクラスに限る</remarks>
    Public Shared Function NewInstance(Of T)(ByVal propertiesObj As T, Optional ignoreScope As Boolean = False) As T
        Return DirectCast(NewInstance(GetType(T), propertiesObj, ignoreScope), T)
    End Function

    ''' <summary>
    ''' 新しいインスタンスを生成する
    ''' </summary>
    ''' <param name="aType">生成するインスタンスの型</param>
    ''' <param name="propertiesObj">コピーさせたいプロパティ値をもつobject</param>
    ''' <param name="ignoreScope">非Publicコンストラクタでもインスタンス生成する場合、true</param>
    ''' <returns>新しいインスタンス</returns>
    ''' <remarks>ただし引数無しコンストラクタを持つクラスに限る</remarks>
    Public Shared Function NewInstance(aType As Type, propertiesObj As Object, Optional ignoreScope As Boolean = False) As Object
        If aType.IsArray Then
            Dim elementType As Type = aType.GetElementType
            Dim propertiesArray As Array = DirectCast(DirectCast(propertiesObj, Object), Array)
            Dim results As Array = Array.CreateInstance(elementType, If(propertiesArray Is Nothing, 0, propertiesArray.Length))
            For i As Integer = 0 To If(results Is Nothing, 0, results.Length) - 1
                results.SetValue(ResolveValue(propertiesArray.GetValue(i), elementType), i)
            Next
            Return results
        End If
        Dim newInsta As Object = Activator.CreateInstance(aType, nonPublic:=ignoreScope)
        If propertiesObj IsNot Nothing Then
            CopyProperties(propertiesObj, newInsta)
        End If
        Return newInsta
    End Function

    ''' <summary>
    ''' 適切なコンストラクタを探索する
    ''' </summary>
    ''' <param name="aType">コンストラクタを持つ型</param>
    ''' <param name="requiredArgumentTypes">必要な引数の型[]</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DetectConstructorInfo(ByVal aType As Type, requiredArgumentTypes As IEnumerable(Of Type)) As ConstructorInfo
        Dim memberTypes As New List(Of Type)(requiredArgumentTypes)
        If memberTypes.Count = 0 AndAlso aType.GetInterfaces.Contains(GetType(PrimitiveValueObject)) Then
            Dim baseType As Type = aType
            Do
                baseType = baseType.BaseType
                Dim genericArguments As Type() = baseType.GetGenericArguments
                If genericArguments.Length = 1 AndAlso GetType(PrimitiveValueObject(Of )).MakeGenericType(genericArguments(0)).IsAssignableFrom(aType) Then
                    memberTypes.AddRange(genericArguments)
                    Exit Do
                End If
            Loop While baseType.BaseType IsNot Nothing
        End If

        Dim constructors As ConstructorInfo() = aType.GetConstructors(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
        If CollectionUtil.IsEmpty(constructors) Then
            Throw New ArgumentException("Publicなコンストラクタがないので無理." & aType.FullName)
        End If
        If constructors.Length = 1 Then
            Return constructors(0)
        End If
        Dim infoByUnmatchParameters As New Dictionary(Of List(Of Type), ConstructorInfo)
        For Each info As ConstructorInfo In constructors
            Dim parameterTypes As List(Of Type) = info.GetParameters.Select(Function(p) p.ParameterType).ToList
            infoByUnmatchParameters.Add(parameterTypes, info)
            For Each memberType As Type In memberTypes
                If parameterTypes.Contains(memberType) Then
                    parameterTypes.Remove(memberType)
                End If
            Next
        Next
        Dim bestMatchs As IEnumerable(Of KeyValuePair(Of List(Of Type), ConstructorInfo)) = infoByUnmatchParameters.Where(Function(pair) pair.Key.Count = 0 AndAlso 0 < pair.Value.GetParameters.Count)
        If CollectionUtil.IsNotEmpty(bestMatchs) Then
            Return bestMatchs(0).Value
        End If
        Dim parameterCountMatchs As List(Of ConstructorInfo) = constructors.Where(Function(info) info.GetParameters.Length = memberTypes.Count).ToList
        If parameterCountMatchs.Count = 1 Then
            Return parameterCountMatchs(0)
        ElseIf 1 < parameterCountMatchs.Count Then
            Dim keyValuePairs As List(Of KeyValuePair(Of List(Of Type), ConstructorInfo)) = infoByUnmatchParameters.Where(Function(pair) parameterCountMatchs.Contains(pair.Value)).OrderBy(Function(pair) pair.Key.Count).ToList
            Return keyValuePairs(0).Value
        End If
        Return constructors(0)
    End Function

    ''' <summary>
    ''' コレクション要素用の新しい空インスタンスを作成する
    ''' </summary>
    ''' <param name="elementType">コレクション要素の型</param>
    ''' <returns>空要素</returns>
    ''' <remarks></remarks>
    Public Shared Function NewInstanceForElement(ByVal elementType As Type) As Object
        If elementType Is TypeUtil.TypeString Then
            Return Nothing
        End If
        Return NewInstance(elementType)
    End Function

    ''' <summary>
    ''' （コレクション型の）要素を作成する ※但し不変オブジェクトは作成しない
    ''' </summary>
    ''' <param name="collectionType">コレクション型</param>
    ''' <returns>新しい空の要素インスタンス</returns>
    ''' <remarks></remarks>
    Public Shared Function NewElementInstanceFromCollectionType(ByVal collectionType As Type) As Object
        Return NewInstanceForElement(TypeUtil.DetectElementType(collectionType))
    End Function

    ''' <summary>
    ''' 「型」からコレクションを作成する
    ''' </summary>
    ''' <param name="collectionType">コレクション型（含む配列型）</param>
    ''' <param name="size">コレクションの大きさ</param>
    ''' <returns>新しいコレクションインスタンス</returns>
    ''' <remarks></remarks>
    Public Shared Function NewCollectionInstance(ByVal collectionType As Type, ByVal size As Integer) As Object
        Return NewCollectionInstanceWithInitialElement(collectionType, size, Nothing)
    End Function

    ''' <summary>
    ''' 「型」からコレクションを作成する
    ''' </summary>
    ''' <param name="collectionType">コレクション型（含む配列型）</param>
    ''' <param name="size">コレクションの大きさ</param>
    ''' <returns>新しいコレクションインスタンス</returns>
    ''' <remarks></remarks>
    Public Shared Function NewCollectionInstanceWithInitialElement(ByVal collectionType As Type, ByVal size As Integer, ByVal initialValues As ICollection(Of Object)) As Object

        TypeUtil.AssertImplementsICollection(collectionType)

        Dim initials As New List(Of Object)
        If initialValues IsNot Nothing Then
            initials.AddRange(initialValues)
        End If

        Dim elementType As Type = TypeUtil.DetectElementType(collectionType)
        Dim anArray As Array = Array.CreateInstance(elementType, size)
        For i As Integer = 0 To size - 1
            If i < initials.Count Then
                anArray.SetValue(initials(i), i)
            Else
                anArray.SetValue(NewInstanceForElement(elementType), i)
            End If
        Next

        If TypeUtil.IsTypeArray(collectionType) Then
            Return anArray
        End If
        Return Activator.CreateInstance(collectionType, anArray)
    End Function

    ''' <summary>
    ''' 新しい配列インスタンスを生成する
    ''' </summary>
    ''' <typeparam name="T">生成する配列インスタンスの要素型</typeparam>
    ''' <param name="size">配列サイズ</param>
    ''' <returns>新しいインスタンス</returns>
    ''' <remarks></remarks>
    Public Shared Function NewArrayInstance(Of T)(ByVal size As Integer) As T()
        If size = 0 Then
            Return DirectCast(Array.CreateInstance(GetType(T), 0), T())
        End If
        Dim result(size - 1) As T
        Return result
    End Function

    ''' <summary>
    ''' 新しいインスタンスを作成する（ディープコピー）
    ''' </summary>
    ''' <typeparam name="T">生成するインスタンスの型</typeparam>
    ''' <param name="obj">コピーしたいプロパティ値をもつobject</param>
    ''' <returns>新しいインスタンス</returns>
    ''' <remarks></remarks>
    Public Shared Function NewInstanceDeep(Of T)(obj As T) As T
        Return EzUtil.CopyForSerializable(Of T)(obj)
    End Function

    ''' <summary>
    ''' プロパティ値が同じかを返す
    ''' </summary>
    ''' <param name="vo1">値obj</param>
    ''' <param name="vo2">値obj</param>
    ''' <param name="ignoreAttributeTypes">同値判定から除外するプロパティに付与してる属性[]</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function IsEquals(ByVal vo1 As Object, ByVal vo2 As Object, ParamArray ignoreAttributeTypes As Type()) As Boolean
        If vo1 Is Nothing OrElse vo2 Is Nothing Then
            Throw New ArgumentException("Null値は対象外")
        End If
        Return PerformIsEquals(vo1, vo2, ignoreAttributeTypes)
    End Function

    ''' <summary>
    ''' プロパティ値が同じかを返す
    ''' </summary>
    ''' <param name="vo1">値obj</param>
    ''' <param name="vo2">値obj</param>
    ''' <param name="ignoreAttributeTypes">同値判定から除外するプロパティに付与してる属性[]</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Private Shared Function PerformIsEquals(ByVal vo1 As Object, ByVal vo2 As Object, ParamArray ignoreAttributeTypes As Type()) As Boolean
        If vo1 Is Nothing AndAlso vo2 Is Nothing Then
            Return True
        ElseIf vo1 Is Nothing OrElse vo2 Is Nothing Then
            Return False
        End If
        If vo1.GetType IsNot vo2.GetType Then
            Return False
        End If

        Dim voType As Type = vo1.GetType
        If TypeUtil.IsTypeImmutable(voType) Then
            Return vo1.Equals(vo2)

        ElseIf TypeUtil.IsTypeCollection(voType) Then
            Dim vo1Array As Object() = ConvObjectToArray(vo1)
            Dim vo2Array As Object() = ConvObjectToArray(vo2)
            If vo1Array.Length <> vo2Array.Length Then
                Return False
            End If
            For i As Integer = 0 To vo1Array.Length - 1
                If Not PerformIsEquals(vo1Array(i), vo2Array(i)) Then
                    Return False
                End If
            Next

        ElseIf TypeOf vo1 Is ICollectionObject Then
            Dim vo1Objects As ICollectionObject = DirectCast(vo1, ICollectionObject)
            Dim vo2Objects As ICollectionObject = DirectCast(vo2, ICollectionObject)
            If vo1Objects.Count <> vo2Objects.Count Then
                Return False
            End If
            For i As Integer = 0 To vo1Objects.Count - 1
                If Not PerformIsEquals(vo1Objects(i), vo2Objects(i)) Then
                    Return False
                End If
            Next

        Else
            For Each aProperty As PropertyInfo In voType.GetProperties()
                If Not aProperty.CanRead Then
                    Continue For
                End If
                Dim customAttributes As Object() = aProperty.GetCustomAttributes(inherit:=False)
                If ignoreAttributeTypes IsNot Nothing AndAlso customAttributes IsNot Nothing _
                        AndAlso customAttributes.Any(Function(attr) ignoreAttributeTypes.Contains(attr.GetType)) Then
                    Continue For
                End If

                If Not PerformIsEquals(aProperty.GetValue(vo1, Nothing), aProperty.GetValue(vo2, Nothing)) Then
                    Return False
                End If
            Next
        End If

        Return True
    End Function

    ''' <summary>
    ''' 値を取得する
    ''' </summary>
    ''' <param name="vo">値object</param>
    ''' <param name="name">プロパティ名</param>
    ''' <param name="voIfNull">voがnullだった時の、戻り値</param>
    ''' <param name="includesField">フィールド値を含める場合、true</param>
    ''' <returns>値</returns>
    ''' <remarks>括弧"XxxList(3)"を使用しcollection要素の参照可。"."経由で更にプロパティ値の参照可。</remarks>
    Public Shared Function GetValue(ByVal vo As Object, ByVal name As String, Optional ByVal voIfNull As Object = Nothing, _
                                    Optional ByVal includesField As Boolean = False) As Object
        Const OPEN_BRACKET As Char = "("c
        Const CLOSE_BRACKET As Char = ")"c
        Const PROPERTY_SEPARATOR As Char = "."c
        If vo Is Nothing Then
            If voIfNull Is Nothing Then
                Throw New NullReferenceException("パラメータ名 " & name & " をもつはずの ValueObject が null.")
            End If
            Return voIfNull
        End If
        Dim firstElem As Integer = name.IndexOf(OPEN_BRACKET)
        Dim firstPeriod As Integer = name.IndexOf(PROPERTY_SEPARATOR)
        If firstPeriod < 0 AndAlso firstElem < 0 Then
            Dim aProperty As PropertyInfo = vo.GetType.GetProperty(name)
            If aProperty Is Nothing Then
                If includesField Then
                    Dim aField As FieldInfo = vo.GetType.GetField(name)
                    If aField IsNot Nothing Then
                        Return aField.GetValue(vo)
                    End If
                End If
                If PROPERTY_NAME_SELF_VALUE.Equals(name) Then
                    Return vo
                End If
                Throw New ArgumentException("パラメータ名 " & name & " に相当するプロパティは、パラメータ値 " & vo.GetType.Name & " に無い.")
            End If
            Return aProperty.GetValue(vo, Nothing)

        ElseIf firstPeriod < 0 Then
            Dim o As Object = GetValue(vo, name.Substring(0, firstElem), voIfNull)
            If o Is Nothing Then
                Return o
            End If

            Dim values As List(Of Object)
            If TypeOf o Is ICollection Then
                Dim value As ICollection = DirectCast(o, ICollection)
                values = value.Cast(Of Object)().ToList()
            ElseIf TypeOf o Is ICollectionObject Then
                Dim value As ICollectionObject = DirectCast(o, ICollectionObject)
                values = Enumerable.Range(0, value.Count).Select(Function(i) value(i)).ToList()
            Else
                Throw New ArgumentException("パラメータ名 " & name.Substring(0, firstElem) & " はコレクションではない. " & o.GetType.Name)
            End If

            Dim firstCloseElem As Integer = name.IndexOf(CLOSE_BRACKET)
            Dim strIndex As String = name.Substring(firstElem + 1, firstCloseElem - firstElem - 1)
            If Not IsNumeric(strIndex) Then
                Throw New ArgumentException("パラメータ名 " & name.Substring(0, firstElem) & " の要素が数値じゃない. " & strIndex)
            End If
            Dim index As Integer = CInt(strIndex)
            If values.Count <= index Then
                Throw New ArgumentException("パラメータ名 " & name.Substring(0, firstElem) & " の要素数は、" & values.Count & " なのに指定の要素は " & strIndex)
            End If
            Return values(index)

        ElseIf firstElem < 0 Then
            Return GetValue(GetValue(vo, name.Substring(0, firstPeriod), voIfNull), name.Substring(firstPeriod + 1), voIfNull)

        ElseIf firstPeriod < firstElem Then
            Return GetValue(GetValue(vo, name.Substring(0, firstPeriod), voIfNull), name.Substring(firstPeriod + 1), voIfNull)

        Else ' If firstElem < firstPeriod Then
            Return GetValue(GetValue(vo, name.Substring(0, firstPeriod), voIfNull), name.Substring(firstPeriod + 1), voIfNull)
        End If
    End Function

    ''' <summary>
    ''' 配列、またはコレクションを、Object型コレクションにして返す
    ''' </summary>
    ''' <param name="arrayOrCollection">配列、またはコレクション</param>
    ''' <returns>Object型のコレクション</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvObjectToCollection(ByVal arrayOrCollection As Object) As ICollection(Of Object)
        If Not TypeUtil.IsArrayOrCollection(arrayOrCollection) Then
            Throw New ArgumentException("コレクションではない. " & arrayOrCollection.GetType.Name)
        End If
        Dim values As ICollection = DirectCast(arrayOrCollection, ICollection)
        Return values.Cast(Of Object)().ToList()
    End Function

    ''' <summary>
    ''' 配列、またはコレクションを、配列型にして返す
    ''' </summary>
    ''' <param name="arrayOrCollection">配列、またはコレクション</param>
    ''' <returns>配列型</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvObjectToArray(ByVal arrayOrCollection As Object) As Object()
        If Not TypeUtil.IsArrayOrCollection(arrayOrCollection) Then
            Throw New ArgumentException("コレクションではない. " & arrayOrCollection.GetType.Name)
        End If
        Dim values As ICollection = DirectCast(arrayOrCollection, ICollection)
        Return values.Cast(Of Object)().ToArray()
    End Function

    ''' <summary>
    ''' 値を型にあわせて変換する
    ''' </summary>
    ''' <param name="value">値</param>
    ''' <param name="aType">型</param>
    ''' <returns>適切な型に変換した値</returns>
    ''' <remarks></remarks>
    Public Shared Function ResolveValue(ByVal value As Object, ByVal aType As Type) As Object
        Return ResolveValue(value, aType, False)
    End Function

    ''' <summary>
    ''' 値を型にあわせて変換する
    ''' </summary>
    ''' <param name="value">値</param>
    ''' <param name="aType">型</param>
    ''' <param name="exceptionsIfInvalid">値が不正なら例外にする場合、true</param>
    ''' <returns>適切な型に変換した値</returns>
    ''' <remarks></remarks>
    Public Shared Function ResolveValue(ByVal value As Object, ByVal aType As Type, ByVal exceptionsIfInvalid As Boolean) As Object
        If value Is Nothing Then
            Return Nothing
        End If
        Dim valueType As Type = value.GetType

        If aType.IsArray Then
            If Not valueType.IsArray Then
                ThrowResolveValue(value, aType)
            End If
            Dim elementType As Type = aType.GetElementType
            Dim valueElementType As Type = valueType.GetElementType
            Dim valueArray As Array = DirectCast(value, Array)
            If elementType.IsAssignableFrom(valueElementType) Then
                Return value
            End If
            Dim results As Array = Array.CreateInstance(elementType, valueArray.Length)
            For i As Integer = 0 To valueArray.Length - 1
                results.SetValue(ResolveValue(valueArray.GetValue(i), elementType, exceptionsIfInvalid), i)
            Next
            Return results
        End If

        Dim propertyType As Type = TypeUtil.GetTypeIfNullable(aType)

        If propertyType.IsAssignableFrom(valueType) Then
            'value が propertyType のサブクラスならそのまま返す
            Return value

        ElseIf propertyType Is TypeUtil.TypeString Then
            Return StringUtil.ToString(value)

        ElseIf propertyType Is GetType(Int32) Then
            Dim result As Integer
            If Integer.TryParse(value.ToString, result) Then
                Return result
            ElseIf exceptionsIfInvalid AndAlso StringUtil.IsNotEmpty(value) Then
                ThrowResolveValue(value, aType)
            End If

        ElseIf propertyType Is TypeUtil.TypeDecimal Then
            Dim result As Decimal
            If Decimal.TryParse(value.ToString, result) Then
                Return result
            ElseIf exceptionsIfInvalid AndAlso StringUtil.IsNotEmpty(value) Then
                ThrowResolveValue(value, aType)
            End If

        ElseIf propertyType Is TypeUtil.TypeDateTime Then
            If IsDate(value) Then
                ' Convert.ToDateTime("11:22:33") は結果が [今日] 11:22:33 になる. CDateはならない.
                Return CDate(value)
            ElseIf exceptionsIfInvalid AndAlso StringUtil.IsNotEmpty(value) Then
                ThrowResolveValue(value, aType)
            End If

        ElseIf propertyType Is GetType(Int64) Then
            Dim result As Long
            If Long.TryParse(value.ToString, result) Then
                Return result
            ElseIf exceptionsIfInvalid AndAlso StringUtil.IsNotEmpty(value) Then
                ThrowResolveValue(value, aType)
            End If

        ElseIf propertyType Is GetType(Double) Then
            Dim result As Double
            If Double.TryParse(value.ToString, result) Then
                Return result
            ElseIf exceptionsIfInvalid AndAlso StringUtil.IsNotEmpty(value) Then
                ThrowResolveValue(value, aType)
            End If

        ElseIf propertyType Is GetType(Boolean) Then
            Return EzUtil.IsTrue(value)

        ElseIf propertyType Is GetType(Single) Then
            Dim result As Single
            If Single.TryParse(value.ToString, result) Then
                Return result
            ElseIf exceptionsIfInvalid AndAlso StringUtil.IsNotEmpty(value) Then
                ThrowResolveValue(value, aType)
            End If

        ElseIf propertyType Is GetType(Byte) Then
            Dim result As Byte
            If Byte.TryParse(value.ToString, result) Then
                Return result
            ElseIf exceptionsIfInvalid AndAlso StringUtil.IsNotEmpty(value) Then
                ThrowResolveValue(value, aType)
            End If

        ElseIf propertyType.IsEnum Then
            If IsUnsupportedTypeForEnum(valueType) _
                    OrElse IsEnumButTypeDifferent(valueType, propertyType) _
                    OrElse Not [Enum].IsDefined(propertyType, value) Then
                ThrowResolveValue(value, aType)
            End If
            Return [Enum].Parse(propertyType, StringUtil.ToString(value))

        ElseIf exceptionsIfInvalid Then
            ThrowResolveValue(value, aType)
        Else
            Return value
        End If

        Return Nothing
    End Function

    Private Shared Sub ThrowResolveValue(ByVal value As Object, ByVal aType As Type)
        Throw New ArgumentException(String.Format("{0}型（値={1}）を {2} 型へ変換できない", value.GetType.Name, value, aType.Name))
    End Sub

    ''' <summary>
    ''' Enumで変換サポートされていない型かどうか
    ''' </summary>
    ''' <param name="aType">型</param>
    ''' <returns>Enumで変換サポートされていない型ならTrue、それ以外はFalse</returns>
    ''' <remarks></remarks>
    Private Shared Function IsUnsupportedTypeForEnum(ByVal aType As Type) As Boolean
        Return _
            aType IsNot GetType(String) _
            AndAlso aType IsNot GetType(SByte) _
            AndAlso aType IsNot GetType(Byte) _
            AndAlso aType IsNot GetType(Int16) _
            AndAlso aType IsNot GetType(Int32) _
            AndAlso aType IsNot GetType(Int64) _
            AndAlso aType IsNot GetType(UInt16) _
            AndAlso aType IsNot GetType(UInt32) _
            AndAlso aType IsNot GetType(UInt64)
    End Function

    ''' <summary>
    ''' Enum同士だが型が異なるかどうか
    ''' </summary>
    ''' <param name="type1">型1</param>
    ''' <param name="type2">型2</param>
    ''' <returns>Enum同士だが型が異なる場合True、それ以外False</returns>
    ''' <remarks></remarks>
    Private Shared Function IsEnumButTypeDifferent(ByVal type1 As Type, ByVal type2 As Type) As Boolean
        Return type1.IsEnum AndAlso type2.IsEnum AndAlso (type1 IsNot type2)
    End Function

    ''' <summary>
    ''' 式のinstanceを差し替えて評価する
    ''' </summary>
    ''' <param name="vo">差し替えるinstance</param>
    ''' <param name="anExpression">式</param>
    ''' <returns>評価した結果</returns>
    ''' <remarks></remarks>
    Public Shared Function InvokeExpressionBy(Of T)(ByVal vo As Object, ByVal anExpression As Expression(Of Func(Of T))) As Object
        Return InvokeExpressionBy(vo, DirectCast(anExpression, Expression))
    End Function
    ''' <summary>
    ''' 式のinstanceを差し替えて評価する
    ''' </summary>
    ''' <param name="vo">差し替えるinstance</param>
    ''' <param name="anExpression">式</param>
    ''' <returns>評価した結果</returns>
    ''' <remarks></remarks>
    Public Shared Function InvokeExpressionBy(ByVal vo As Object, ByVal anExpression As Expression) As Object
        If vo Is Nothing Then
            Return Nothing
        End If
        Select Case anExpression.NodeType
            Case ExpressionType.Call
                Dim methodCall As MethodCallExpression = DirectCast(anExpression, MethodCallExpression)
                Dim args As New List(Of Object)
                For Each argExpression As Expression In methodCall.Arguments
                    If argExpression.NodeType = ExpressionType.Lambda Then
                        Throw New NotSupportedException("内部のLambdaは未対応")
                    End If
                    args.Add(InvokeExpressionBy(vo, argExpression))
                Next
                Dim resolvedInstance As Object = If(methodCall.Object Is Nothing OrElse vo.GetType Is methodCall.Object.Type, _
                                                    vo, InvokeExpressionBy(vo, methodCall.Object))
                Return methodCall.Method.Invoke(resolvedInstance, args.ToArray)
            Case ExpressionType.MemberAccess
                Dim memberExpr As MemberExpression = DirectCast(anExpression, MemberExpression)
                If TypeUtil.IsTypeImmutable(vo.GetType) _
                        AndAlso memberExpr.Member.MemberType = MemberTypes.Field _
                        AndAlso vo.GetType Is DirectCast(memberExpr.Member, FieldInfo).FieldType Then
                    Return vo
                End If
                Dim resolvedInstance As Object = If(memberExpr.Expression Is Nothing, Nothing,
                                                    If(vo.GetType Is memberExpr.Expression.Type,
                                                       vo, InvokeExpressionBy(vo, memberExpr.Expression)))
                Return GetValueFrom(memberExpr.Member, resolvedInstance)
            Case ExpressionType.Convert, ExpressionType.ConvertChecked
                Return ResolveUnary(vo, DirectCast(anExpression, UnaryExpression))
            Case ExpressionType.ArrayLength
                Dim result As Object = ResolveUnary(vo, DirectCast(anExpression, UnaryExpression))
                Return DirectCast(result, Array).Length
            Case ExpressionType.Constant
                Return DirectCast(anExpression, ConstantExpression).Value
            Case ExpressionType.ArrayIndex
                Dim binary As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                Dim valueArray As Array = DirectCast(binary.Left, Array)
                Dim index As Integer = CInt(binary.Right)
                Return valueArray.GetValue(index)
            Case ExpressionType.Conditional
                Dim conditional As ConditionalExpression = DirectCast(anExpression, ConditionalExpression)
                If DirectCast(InvokeExpressionBy(vo, conditional.Test), Boolean) Then
                    Return InvokeExpressionBy(vo, conditional.IfTrue)
                Else
                    Return InvokeExpressionBy(vo, conditional.IfFalse)
                End If
            Case ExpressionType.Equal
                Dim binary As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                Return If(binary.Left Is Nothing, binary.Right Is Nothing, binary.Left.Equals(binary.Right))
            Case ExpressionType.NotEqual
                Dim binary As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                Return Not If(binary.Left Is Nothing, binary.Right Is Nothing, binary.Left.Equals(binary.Right))
            Case ExpressionType.LessThan
                Dim binary As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                Return CDec(binary.Left) < CDec(binary.Right)
            Case ExpressionType.LessThanOrEqual
                Dim binary As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                Return CDec(binary.Left) <= CDec(binary.Right)
            Case ExpressionType.GreaterThan
                Dim binary As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                Return CDec(binary.Right) < CDec(binary.Left)
            Case ExpressionType.GreaterThanOrEqual
                Dim binary As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                Return CDec(binary.Right) <= CDec(binary.Left)
            Case ExpressionType.Coalesce
                Dim binary As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                Return If(binary.Left, binary.Right)
            Case ExpressionType.Lambda
                Return InvokeExpressionBy(vo, DirectCast(anExpression, LambdaExpression).Body)
            Case ExpressionType.New
                Dim newExpr As NewExpression = DirectCast(anExpression, NewExpression)
                Return newExpr.Constructor.Invoke(newExpr.Arguments.Select(Function(arg) InvokeExpressionBy(vo, arg)).ToArray)
            Case ExpressionType.MemberInit
                Dim memberInit As MemberInitExpression = DirectCast(anExpression, MemberInitExpression)
                Dim newInstance As Object = InvokeExpressionBy(vo, memberInit.NewExpression)
                For Each binding As MemberBinding In memberInit.Bindings
                    Dim memberValue As Object
                    Select Case binding.BindingType
                        Case MemberBindingType.Assignment
                            memberValue = InvokeExpressionBy(vo, DirectCast(binding, MemberAssignment).Expression)
                        Case Else
                            Throw New NotSupportedException("With句の初期化以外は未対応")
                    End Select
                    SetValueFrom(binding.Member, newInstance, memberValue)
                Next
                Return newInstance
            Case ExpressionType.MultiplyChecked
                Dim result As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                AssertIntegerResult(result)
                Return CInt(result.Left) * CInt(result.Right)
            Case ExpressionType.AddChecked
                Dim result As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                AssertIntegerResult(result)
                Return CInt(result.Left) + CInt(result.Right)
            Case ExpressionType.SubtractChecked
                Dim result As BinaryResult = ResolveBinary(vo, DirectCast(anExpression, BinaryExpression))
                AssertIntegerResult(result)
                Return CInt(result.Left) - CInt(result.Right)
            Case ExpressionType.Divide
                Throw New NotSupportedException(String.Format("除数(right={0})がDoubleになるので未対応", DirectCast(anExpression, BinaryExpression).Right))
            Case Else
                Throw New NotSupportedException(String.Format("必要なら対応して！ 未対応のNodeType:{0}({1})", anExpression.NodeType.ToString, CInt(anExpression.NodeType)))
        End Select
    End Function

    Private Shared Sub AssertIntegerResult(result As BinaryResult)
        If Not TypeOf result.Left Is Integer OrElse Not TypeOf result.Right Is Integer Then
            Throw New NotSupportedException(String.Format("四則演算はInt型同士だけ可能. 必要なら拡張すべし (left, right) = ({0}, {1})", result.Left.GetType.Name, result.Right.GetType.Name))
        End If
    End Sub

    Private Shared Sub SetValueFrom(memberInfo As MemberInfo, instance As Object, memberValue As Object)
        If TypeOf memberInfo Is PropertyInfo Then
            DirectCast(memberInfo, PropertyInfo).SetValue(instance, memberValue, Nothing)
            Return
        ElseIf TypeOf memberInfo Is FieldInfo Then
            DirectCast(memberInfo, FieldInfo).SetValue(instance, memberValue)
            Return
        End If
        Throw New InvalidOperationException("PropertyInfoでもFieldInfoでもない.あり得ないハズ. memberInfo.GetType=" & memberInfo.GetType.Name)
    End Sub

    Private Shared Function GetValueFrom(memberInfo As MemberInfo, instance As Object) As Object
        If TypeOf memberInfo Is PropertyInfo Then
            Return DirectCast(memberInfo, PropertyInfo).GetValue(instance, Nothing)
        ElseIf TypeOf memberInfo Is FieldInfo Then
            Return DirectCast(memberInfo, FieldInfo).GetValue(instance)
        End If
        Throw New InvalidOperationException("PropertyInfoでもFieldInfoでもない.あり得ないハズ. memberInfo.GetType=" & memberInfo.GetType.Name)
    End Function

    Private Shared Function ResolveUnary(ByVal vo As Object, ByVal anExpression As UnaryExpression) As Object
        Dim value As Object = InvokeExpressionBy(vo, anExpression.Operand)
        If anExpression.Method Is Nothing Then
            Return value
        End If
        Return anExpression.Method.Invoke(vo, {value})
    End Function

    Private Class BinaryResult
        Public Left As Object
        Public Right As Object
    End Class

    Private Shared Function ResolveBinary(ByVal vo As Object, ByVal anExpression As BinaryExpression) As BinaryResult
        Dim left As Object = InvokeExpressionBy(vo, anExpression.Left)
        Dim right As Object = InvokeExpressionBy(vo, anExpression.Right)
        Return New BinaryResult With {.Left = left, .Right = right}
    End Function

    ''' <summary>
    ''' プロパティ名一覧を取得する
    ''' </summary>
    ''' <param name="vo">対象の値Object</param>
    ''' <param name="condition">(option) 絞り込み条件</param>
    ''' <returns>名前一覧</returns>
    ''' <remarks></remarks>
    Public Shared Function GetPropertyNames(Of T)(vo As T, Optional condition As Func(Of String, Boolean) = Nothing) As String()
        Return GetPropertyNames(GetType(T), condition)
    End Function

    ''' <summary>
    ''' プロパティ名一覧を取得する
    ''' </summary>
    ''' <typeparam name="T">対象の型</typeparam>
    ''' <param name="condition">(option) 絞り込み条件</param>
    ''' <returns>名前一覧</returns>
    ''' <remarks></remarks>
    Public Shared Function GetPropertyNames(Of T)(Optional condition As Func(Of String, Boolean) = Nothing) As String()
        Return GetPropertyNames(Of T)(Nothing, condition)
    End Function

    ''' <summary>
    ''' プロパティ名一覧を取得する
    ''' </summary>
    ''' <param name="type">プロパティ名一覧を取得したい型</param>
    ''' <param name="condition">(option) 絞り込み条件</param>
    ''' <returns>名前一覧</returns>
    ''' <remarks></remarks>
    Public Shared Function GetPropertyNames(type As Type, Optional condition As Func(Of String, Boolean) = Nothing) As String()
        If type Is Nothing Then
            Throw New ArgumentNullException("type")
        End If
        Return type.GetProperties.Select(Function(aProperty) aProperty.Name).Where(Function(name) If(condition, Function(aName) True).Invoke(name)).ToArray
    End Function

    ''' <summary>
    ''' Voへ値を設定する
    ''' </summary>
    ''' <typeparam name="T">適用したいVoの型</typeparam>
    ''' <param name="vo">適用先Vo</param>
    ''' <param name="propertyName">適用先のプロパティ名</param>
    ''' <param name="value">適用させたい値</param>
    ''' <remarks></remarks>
    Public Shared Sub SetValue(Of T)(vo As T, propertyName As String, value As Object)
        If vo Is Nothing Then
            Throw New ArgumentNullException("vo")
        End If

        Dim info As PropertyInfo = GetType(T).GetProperties.FirstOrDefault(Function(anInfo) anInfo.Name.Equals(propertyName))
        If info Is Nothing Then
            Throw New NotSupportedException(String.Format("プロパティ名:{0}は型:{1}でサポートされていません", propertyName, GetType(T).Name))
        End If
        info.SetValue(vo, VoUtil.ResolveValue(value, info.PropertyType), Nothing)
    End Sub

    ''' <summary>
    ''' プロパティを持っているか？
    ''' </summary>
    ''' <typeparam name="T">Voの型</typeparam>
    ''' <param name="vo">Vo</param>
    ''' <param name="propertyName">プロパティ名</param>
    ''' <returns>プロパティがあればTrue</returns>
    ''' <remarks></remarks>
    Public Shared Function HasProperty(Of T)(vo As T, propertyName As String) As Boolean
        If vo Is Nothing Then
            Throw New ArgumentNullException("vo")
        End If
        Return HasProperty(Of T)(propertyName)
    End Function

    ''' <summary>
    ''' プロパティを持っているか？
    ''' </summary>
    ''' <typeparam name="T">Voの型</typeparam>
    ''' <param name="propertyName">プロパティ名</param>
    ''' <returns>プロパティがあればTrue</returns>
    ''' <remarks></remarks>
    Public Shared Function HasProperty(Of T)(propertyName As String) As Boolean
        Return GetType(T).GetProperty(propertyName) IsNot Nothing
    End Function

    ''' <summary>
    ''' 値を(PVO含む)型へ解決する
    ''' </summary>
    ''' <param name="value">値</param>
    ''' <param name="aType">変換後の型</param>
    ''' <returns>変換後の値</returns>
    ''' <remarks></remarks>
    Friend Shared Function ResolveValueContainsPVO(value As Object, aType As Type) As Object
        If Not TypeUtil.IsTypeValueObjectOrSubClass(aType) Then
            Return ResolveValue(value, aType)
        End If
        If value Is Nothing Then
            Return Nothing
        End If
        Dim constructor As ConstructorInfo = ValueObject.DetectConstructor(aType)
        Dim parameterInfos As ParameterInfo() = constructor.GetParameters
        If parameterInfos.Length <> 1 Then
            Throw New InvalidOperationException("引数のないPrimitiveValueObjectは対応できない")
        End If
        Return constructor.Invoke(New Object() {value})
    End Function

    ''' <summary>
    ''' ラムダ式でプロパティ名を取得する
    ''' </summary>
    ''' <typeparam name="T">Voの型</typeparam>
    ''' <param name="expression">ラムダ式</param>
    ''' <returns>プロパティ名</returns>
    Public Shared Function GetPropertyName(Of T)(expression As Expression(Of Func(Of T, Object))) As String
        Dim marker As New VoPropertyMarker
        marker.CreateMarkedVo(Of T)()
        Return marker.GetPropertyInfo(expression).Name
    End Function

    ''' <summary>
    ''' プロパティをラムダ式で一括指定してプロパティ名を取得する
    ''' </summary>
    ''' <typeparam name="T">プロパティ名を取得したいVoの型</typeparam>
    ''' <param name="configure">プロパティを指定する為のDelegate</param>
    ''' <returns>プロパティ名[]</returns>
    ''' <remarks></remarks>
    Public Shared Function GetSpecifyPropertyNames(Of T)(configure As PropertySpecifierConfigure(Of T)) As String()
        Dim specifiedProperties As String() = New PropertySpecifier(Of T)(configure).Results
        Return specifiedProperties.Intersect(GetPropertyNames(Of T)()).ToArray
    End Function

     ''' <summary>プロパティを指定する為のDelegate</summary>
    Public Delegate Function PropertySpecifierConfigure(Of T)(define As IPropertySpecifier, vo As T) As IPropertySpecifier
    Public Interface IPropertySpecifier
        ''' <summary>プロパティを指定する</summary>
        Function AppendProperties(ParamArray props As Object()) As IPropertySpecifier
        ''' <summary>Delegateでプロパティを指定する</summary>
        ''' <remarks>Boolean三項目以上持つVoだと内部使用しているVoPropertyMarkerでプロパティを特定できない為、こっちを使ってほしい</remarks>
        Function AppendPropertyWithFunc(Of T)(expression As Expression(Of Func(Of T))) As IPropertySpecifier
    End Interface
    Private Class PropertySpecifier(Of T) : Implements IPropertySpecifier
        Private ReadOnly marker As New VoPropertyMarker
        Private ReadOnly _results As New List(Of String)
        Public Sub New(callback As PropertySpecifierConfigure(Of T))
            Dim vo As T = VoUtil.NewInstance(Of T)()
            marker.MarkVo(vo)
            callback.Invoke(Me, vo)
        End Sub
        Public ReadOnly Property Results As String()
            Get
                Return _results.ToArray
            End Get
        End Property
        Public Function AppendProperties(ParamArray props As Object()) As IPropertySpecifier Implements IPropertySpecifier.AppendProperties
            _results.AddRange(props.Select(Function(prop) marker.GetPropertyInfo(prop).Name))
            Return Me
        End Function
        Public Function AppendPropertyWithFunc(Of T1)(expression As Expression(Of Func(Of T1))) As IPropertySpecifier Implements IPropertySpecifier.AppendPropertyWithFunc
            _results.Add(marker.GetPropertyInfo(expression).Name)
            Return Me
        End Function
    End Class

End Class
