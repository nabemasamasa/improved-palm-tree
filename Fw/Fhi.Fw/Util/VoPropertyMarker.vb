Imports System.Reflection
Imports System.Linq.Expressions
Imports Fhi.Fw.Domain

Namespace Util
    ''' <summary>
    ''' Voのプロパティをマーキングする役割を担うクラス
    ''' </summary>
    ''' <remarks>
    ''' 1. プロパティ値がコレクションのときは要素数0のまま（内部要素は別途マーキングしてもらう方針）
    ''' 2. Readonlyなプロパティやフィールドを優先して、極力コンストラクタで初期値を設定する
    '''     - 書き込み可能プロパティはインスタンス生成後に初期値を設定する
    ''' </remarks>
    Public Class VoPropertyMarker

        Private ReadOnly propertyInfoByValue As New Dictionary(Of Object, PropertyInfo)
        Private ReadOnly fieldInfoByValue As New Dictionary(Of Object, FieldInfo)
        Private ReadOnly accessorByValue As New Dictionary(Of Object, Func(Of Object, Object))
        Private wasInitialized As Boolean
        Private properties As List(Of PropertyInfo)
        Private isOutOfSupportBoolean As Boolean = False
        Private markCount As Integer
        Private boolCount As Integer

        ''' <summary>
        ''' Voのプロパティをマーキングする
        ''' </summary>
        ''' <param name="vo">対象Vo</param>
        ''' <remarks></remarks>
        Public Sub MarkVo(ByVal vo As Object)
            Initialize()
            RecurMarkVoExecute(vo)
            BuildPropertyParentAccessor(vo)
            BuildFieldInfoAndAccessor(vo)
        End Sub

        Private Sub Initialize()
            If 0 < propertyInfoByValue.Count OrElse 0 < fieldInfoByValue.Count Then
                Throw New InvalidOperationException("このインスタンスはマーク済. #Clear()するか別のインスタンスでマーキングして")
            End If
            markCount = 0
            boolCount = 0
            properties = New List(Of PropertyInfo)
            wasInitialized = True
        End Sub

        ''' <summary>
        ''' マーキング済みのVoを作成する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CreateMarkedVo(Of T)() As T
            Dim aType As Type = GetType(T)
            If TypeUtil.IsTypeValueObjectOrSubClass(aType) Then
                Initialize()
                Dim result As Object = MakeMarkedValueObject(aType)
                RecurMarkVoExecute(result)
                BuildPropertyParentAccessor(result)
                BuildFieldInfoAndAccessor(result)
                BuildNonPublicIfNecessary(result)
                Return DirectCast(result, T)
            ElseIf TypeUtil.TypeString Is aType Then
                Initialize()
                Return DirectCast(DirectCast(aType.Name, Object), T)
            End If
            Dim newInstance As T = VoUtil.NewInstance(Of T)(ignoreScope:=True)
            MarkVo(newInstance)
            Return newInstance
        End Function

        ''' <summary>
        ''' Voのプロパティをマーキングする
        ''' </summary>
        ''' <param name="vo">対象Vo</param>
        ''' <param name="callingCount">再起呼出し回数</param>
        ''' <remarks></remarks>
        Private Sub RecurMarkVoExecute(ByVal vo As Object, Optional ByVal callingCount As Integer = 0)
            AssertCallLimitNumber(callingCount)
            Dim aType As Type = vo.GetType
            Dim propertyInfos As List(Of PropertyInfo) = aType.GetProperties.Where(Function(p) p.GetSetMethod IsNot Nothing).ToList
            For Each [property] As PropertyInfo In propertyInfos
                Dim aProperty As PropertyInfo = [property]
                MakeAndSetValue(callingCount, aProperty.PropertyType,
                                setValueCallback:=Sub(val)
                                                      aProperty.SetValue(vo, val, Nothing)
                                                      propertyInfoByValue.Add(val, aProperty)
                                                  End Sub)
            Next
            properties.AddRange(propertyInfos)
            For Each info As FieldInfo In aType.GetFields(BindingFlags.Public Or BindingFlags.Instance)
                Dim anInfo As FieldInfo = info
                MakeAndSetValue(callingCount, anInfo.FieldType,
                                setValueCallback:=Sub(val)
                                                      If anInfo.IsInitOnly Then
                                                          Return
                                                      End If
                                                      anInfo.SetValue(vo, val)
                                                  End Sub)
            Next
        End Sub

        Private Sub MakeAndSetValue(ByVal callingCount As Integer, ByVal aType As Type, _
                                    ByVal setValueCallback As Action(Of Object))
            Dim value As Object = ResolveValue(aType, callingCount)
            If value Is Nothing Then
                Return
            End If
            setValueCallback.Invoke(value)
        End Sub

        Private Shared Sub AssertCallLimitNumber(ByVal callingCount As Integer)
            Const RECURSIVE_CALL_LIMIT_NUM As Integer = 100
            If RECURSIVE_CALL_LIMIT_NUM <= callingCount Then
                Throw New InvalidOperationException("Voの階層が深すぎます。無限ループかも")
            End If
        End Sub

        Private Function ResolveValue(ByVal aType As Type, ByVal callingCount As Integer) As Object
            Dim aPropertyType As Type = TypeUtil.GetTypeIfNullable(aType)
            Dim value As Object
            If aType.IsArray Then
                value = Array.CreateInstance(aType.GetElementType, 0)
            ElseIf aPropertyType Is TypeUtil.TypeInteger Then
                value = markCount
            ElseIf aPropertyType Is TypeUtil.TypeString Then
                value = CStr(markCount)
            ElseIf aPropertyType Is TypeUtil.TypeDecimal Then
                value = CDec(markCount)
            ElseIf aPropertyType Is TypeUtil.TypeDateTime Then
                value = DateAdd(DateInterval.Day, markCount, New DateTime())
            ElseIf aPropertyType Is TypeUtil.TypeLong Then
                value = CLng(markCount)
            ElseIf aPropertyType Is TypeUtil.TypeSingle Then
                value = CSng(markCount)
            ElseIf aPropertyType Is TypeUtil.TypeDouble Then
                value = CDbl(markCount)
            ElseIf aPropertyType Is TypeUtil.TypeBoolean Then
                If 2 <= boolCount Then
                    isOutOfSupportBoolean = True
                    Return Nothing
                End If
                value = (boolCount = 1)
                boolCount += 1
            ElseIf TypeUtil.IsTypeValueObjectOrSubClass(aPropertyType) Then
                value = MakeMarkedValueObject(aPropertyType, callingCount + 1)
            ElseIf Not aPropertyType.IsPrimitive Then
                If IsTypeCollection(aPropertyType) AndAlso aPropertyType.IsInterface Then
                    Dim genericArguments As Type() = aPropertyType.GetGenericArguments
                    If genericArguments.Length <> 1 Then
                        Throw New NotSupportedException("未対応." & aPropertyType.FullName)
                    End If
                    value = Activator.CreateInstance(GetType(List(Of )).MakeGenericType(genericArguments(0)))
                Else
                    value = Activator.CreateInstance(aPropertyType)
                End If
                If Not aPropertyType Is TypeUtil.TypeObject AndAlso Not IsTypeCollection(aPropertyType) Then
                    RecurMarkVoExecute(value, callingCount + 1)
                End If
            ElseIf aPropertyType Is GetType(Int16) Then
                value = CShort(markCount)
            ElseIf aPropertyType Is GetType(Byte) Then
                value = CByte(markCount)
            Else
                Throw New NotSupportedException("未対応のプロパティ型です. " & aType.ToString)
            End If
            markCount += 1
            Return value
        End Function

        Private Shared Function IsTypeCollection(ByVal aPropertyType As Type) As Boolean
            Return TypeUtil.IsTypeCollection(aPropertyType) _
                OrElse (TypeUtil.TypeString IsNot aPropertyType AndAlso TypeUtil.IsTypeGenericOrSubClass(aPropertyType, GetType(IEnumerable(Of )))) _
                OrElse GetType(ICollectionObject).IsAssignableFrom(aPropertyType)
        End Function

        ''' <summary>
        ''' マーキングしたキーとプロパティ情報の一覧を取得する
        ''' </summary>
        ''' <returns>キーとプロパティ情報の一覧</returns>
        ''' <remarks></remarks>
        Public Function GetMarkedKeysAndProperties() As Dictionary(Of Object, PropertyInfo)
            Return propertyInfoByValue
        End Function

        ''' <summary>
        ''' 適切なプロパティ値なら処理を実行する
        ''' </summary>
        ''' <param name="aField">Voのプロパティ値</param>
        ''' <param name="performCallback">Callback処理</param>
        ''' <returns>Callback処理の結果</returns>
        ''' <remarks></remarks>
        Private Function PerformIfValid(Of T)(ByVal aField As Object, ByVal performCallback As Func(Of Object, T), ByVal ignoresOutOfSupportBoolean As Boolean) As T
            If Not wasInitialized Then
                Throw New NotSupportedException("MarkVoが実行されていません.")
            End If

            If Not ignoresOutOfSupportBoolean Then
                Assert3OrMoreBoolean(aField)
            End If

            Return performCallback(aField)
        End Function

        ''' <summary>
        ''' Boolean型が3項目以上無いことを保証する
        ''' </summary>
        ''' <param name="aField">マーク値</param>
        ''' <remarks></remarks>
        Public Sub Assert3OrMoreBoolean(ByVal aField As Object)
            Dim isLambda As Boolean = TypeOf aField Is [Delegate]
            If TypeOf aField Is Boolean AndAlso isOutOfSupportBoolean AndAlso (Not isLambda) Then
                Throw New NotSupportedException("Boolean型が3項目以上の場合は、ラムダ式を利用してください")
            End If
        End Sub

        ''' <summary>
        ''' プロパティ情報を取得する
        ''' </summary>
        ''' <param name="aField">Voのプロパティ値</param>
        ''' <returns>プロパティ情報(PropertyInfo)</returns>
        ''' <remarks></remarks>
        Public Function GetPropertyInfo(ByVal aField As Object) As PropertyInfo
            Return PerformIfValid(aField, Function(value) propertyInfoByValue(value), ignoresOutOfSupportBoolean:=False)
        End Function

        ''' <summary>
        ''' プロパティ情報を取得する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="aExpression">ラムダ式</param>
        ''' <returns>プロパティ情報</returns>
        ''' <remarks></remarks>
        Public Function GetPropertyInfo(Of T)(ByVal aExpression As Expression(Of Func(Of T))) As PropertyInfo
            Return properties.First(Function(prop) prop.Name.Equals(GetPropertyName(aExpression)))
        End Function

        ''' <summary>
        ''' プロパティ情報を取得する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="aExpression">ラムダ式</param>
        ''' <returns>プロパティ情報</returns>
        ''' <remarks></remarks>
        Public Function GetPropertyInfo(Of T)(ByVal aExpression As Expression(Of Func(Of T, Object))) As PropertyInfo
            Return properties.First(Function(prop) prop.Name.Equals(GetPropertyName(aExpression)))
        End Function

        ''' <summary>
        ''' プロパティ情報が含まれるかを返す
        ''' </summary>
        ''' <param name="aField">Voのプロパティ値</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Contains(ByVal aField As Object) As Boolean
            Return PerformIfValid(aField, Function(value) propertyInfoByValue.ContainsKey(value), ignoresOutOfSupportBoolean:=True)
        End Function

        ''' <summary>
        ''' プロパティ情報が含まれるかを返す
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="aExpression">ラムダ式</param>
        ''' <returns>プロパティ情報</returns>
        ''' <remarks></remarks>
        Public Function Contains(Of T)(ByVal aExpression As Expression(Of Func(Of T))) As Boolean
            Return properties.Any(Function(prop) prop.Name.Equals(GetPropertyName(aExpression)))
        End Function

        ''' <summary>
        ''' プロパティ名を取得する
        ''' </summary>
        ''' <param name="aExpression">ラムダ式</param>
        ''' <returns>プロパティ名</returns>
        ''' <remarks></remarks>
        Private Function GetPropertyName(ByVal aExpression As Expression) As String
            Return GetMemberExpression(aExpression).Member.Name
        End Function

        ''' <summary>
        ''' MemberExpressionを取得する
        ''' </summary>
        ''' <param name="aExpression">ラムダ式</param>
        ''' <returns>MemberExpression</returns>
        ''' <remarks></remarks>
        Private Function GetMemberExpression(ByVal aExpression As Expression) As MemberExpression
            Dim lambda As LambdaExpression = TryCast(aExpression, LambdaExpression)
            If lambda Is Nothing Then
                Throw New ArgumentNullException
            End If

            Dim memberExpression As MemberExpression = Nothing
            If lambda.Body.NodeType = ExpressionType.Convert Then
                memberExpression = TryCast(DirectCast(lambda.Body, UnaryExpression).Operand, MemberExpression)
            ElseIf lambda.Body.NodeType = ExpressionType.MemberAccess Then
                memberExpression = TryCast(lambda.Body, MemberExpression)
            End If

            If memberExpression Is Nothing Then
                Throw New ArgumentException
            End If

            Return memberExpression
        End Function

        ''' <summary>
        ''' マーキング情報をクリアする
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Clear()
            propertyInfoByValue.Clear()
            fieldInfoByValue.Clear()
            accessorByValue.Clear()
            wasInitialized = False
        End Sub

        Private Function MakeMarkedValueObject(aType As Type, Optional ByVal callingCount As Integer = 0) As Object
            AssertCallLimitNumber(callingCount)
            Dim constructor As ConstructorInfo = ValueObject.DetectConstructor(aType)
            Dim parameterInfos As ParameterInfo() = constructor.GetParameters
            Dim parameterValues As Object() = parameterInfos.Select(Function(info) ResolveValue(info.ParameterType, callingCount)).ToArray
            Return constructor.Invoke(parameterValues)
        End Function

        Private Sub BuildFieldInfoAndAccessor(ByVal value As Object, Optional accessor As Func(Of Object, Object) = Nothing)
            Dim anAccessor As Func(Of Object, Object) = If(accessor, Function(obj) obj)
            For Each info As FieldInfo In value.GetType.GetFields(BindingFlags.Public Or BindingFlags.Instance)
                Dim fieldValue As Object = info.GetValue(value)
                If fieldValue Is Nothing Then
                    Continue For
                End If
                Dim anInfo As FieldInfo = info
                If TypeUtil.IsTypeValueObjectOrSubClass(info.FieldType) Then
                    BuildFieldInfoAndAccessor(fieldValue, Function(obj) anInfo.GetValue(anAccessor(obj)))
                    BuildPropertyParentAccessor(fieldValue, Function(obj) anInfo.GetValue(anAccessor(obj)))
                ElseIf Not TypeUtil.IsTypeImmutable(info.FieldType) AndAlso Not IsTypeCollection(info.FieldType) Then
                    BuildPropertyParentAccessor(fieldValue, Function(obj) anInfo.GetValue(anAccessor(obj)))
                End If
                If Not accessorByValue.ContainsKey(fieldValue) Then
                    accessorByValue.Add(fieldValue, anAccessor)
                End If
                If Not fieldInfoByValue.ContainsKey(fieldValue) Then
                    fieldInfoByValue.Add(fieldValue, info)
                End If
            Next
        End Sub

        Private Sub BuildPropertyParentAccessor(ByVal value As Object, Optional accessor As Func(Of Object, Object) = Nothing)
            Dim anAccessor As Func(Of Object, Object) = If(accessor, Function(obj) obj)
            For Each info As PropertyInfo In value.GetType.GetProperties
                If Not info.CanRead OrElse 0 < info.GetIndexParameters.Length Then
                    Continue For
                End If
                Dim propertyValue As Object = info.GetValue(value, Nothing)
                If propertyValue Is Nothing Then
                    Continue For
                End If
                Dim anInfo As PropertyInfo = info
                If TypeUtil.IsTypeValueObjectOrSubClass(info.PropertyType) Then
                    BuildFieldInfoAndAccessor(propertyValue, Function(obj) anInfo.GetValue(anAccessor(obj), Nothing))
                    BuildPropertyParentAccessor(propertyValue, Function(obj) anInfo.GetValue(anAccessor(obj), Nothing))
                ElseIf Not TypeUtil.IsTypeImmutable(info.PropertyType) AndAlso Not IsTypeCollection(info.PropertyType) Then
                    BuildPropertyParentAccessor(propertyValue, Function(obj) anInfo.GetValue(anAccessor(obj), Nothing))
                End If
                If Not accessorByValue.ContainsKey(propertyValue) Then
                    accessorByValue.Add(propertyValue, anAccessor)
                End If
                If Not propertyInfoByValue.ContainsKey(propertyValue) Then
                    propertyInfoByValue.Add(propertyValue, info)
                End If
            Next
        End Sub

        ''' <summary>
        ''' "Value"のプロパティ情報を取得する
        ''' </summary>
        ''' <param name="aType">"Value"をもつ型</param>
        ''' <returns>※無ければNothing</returns>
        ''' <remarks></remarks>
        Public Shared Function GetPropertyInfoOfValueByType(aType As Type) As PropertyInfo
            Const VALUE_NAME As String = "Value"
            Const ATTR As BindingFlags = BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance
            Dim propertyInfos2 As PropertyInfo() = aType.GetProperties(ATTR)
            Return propertyInfos2.FirstOrDefault(Function(info) VALUE_NAME.Equals(info.Name, StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>
        ''' "Value"のフィールド情報を取得する
        ''' </summary>
        ''' <param name="aType">"Value"をもつ型</param>
        ''' <returns>※無ければNothing</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFieldInfoOfValueByType(aType As Type) As FieldInfo
            Const VALUE_NAME As String = "Value"
            Const ATTR As BindingFlags = BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance
            Dim fieldInfos2 As FieldInfo() = aType.GetFields(ATTR)
            Return fieldInfos2.FirstOrDefault(Function(info) VALUE_NAME.Equals(info.Name, StringComparison.OrdinalIgnoreCase))
        End Function

        Private Sub BuildNonPublicIfNecessary(value As Object)
            If CollectionUtil.IsNotEmpty(propertyInfoByValue) OrElse CollectionUtil.IsNotEmpty(accessorByValue) Then
                Return
            End If
            Dim anAccessor As Func(Of Object, Object) = Function(obj) obj
            Dim aType As Type = value.GetType
            Dim aPropertyInfo As PropertyInfo = GetPropertyInfoOfValueByType(aType)
            If aPropertyInfo IsNot Nothing AndAlso aPropertyInfo.CanRead Then
                Dim propertyValue As Object = aPropertyInfo.GetValue(value, Nothing)
                If Not accessorByValue.ContainsKey(propertyValue) Then
                    accessorByValue.Add(propertyValue, anAccessor)
                End If
                If Not propertyInfoByValue.ContainsKey(propertyValue) Then
                    propertyInfoByValue.Add(propertyValue, aPropertyInfo)
                End If
                Return
            End If

            Dim aFieldInfo As FieldInfo = GetFieldInfoOfValueByType(aType)
            If aFieldInfo IsNot Nothing Then
                Dim fieldValue As Object = aFieldInfo.GetValue(value)
                If Not accessorByValue.ContainsKey(fieldValue) Then
                    accessorByValue.Add(fieldValue, anAccessor)
                End If
                If Not fieldInfoByValue.ContainsKey(fieldValue) Then
                    fieldInfoByValue.Add(fieldValue, aFieldInfo)
                End If
            End If
        End Sub

        ''' <summary>
        ''' マーク値で特定したフィールド値/プロパティ値を取得する
        ''' </summary>
        ''' <param name="aField">フィールド/プロパティを特定する値</param>
        ''' <param name="vo">Vo</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetValue(ByVal aField As Object, ByVal vo As Object) As Object
            Dim accessor As Func(Of Object, Object) =
                PerformIfValid(aField, Function(value) If(accessorByValue.ContainsKey(value), accessorByValue(value), Function(o) o), ignoresOutOfSupportBoolean:=False)
            If propertyInfoByValue.ContainsKey(aField) Then
                Dim info As PropertyInfo = GetPropertyInfo(aField)
                Return info.GetValue(accessor(vo), Nothing)
            ElseIf fieldInfoByValue.ContainsKey(aField) Then
                Dim info As FieldInfo = PerformIfValid(aField, Function(value) fieldInfoByValue(value), ignoresOutOfSupportBoolean:=False)
                Return info.GetValue(accessor(vo))
            End If
            Return accessor(vo)
        End Function

        ''' <summary>
        ''' 名前を取得する
        ''' </summary>
        ''' <param name="aField">フィールド値</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMemberName(aField As Object) As String
            If propertyInfoByValue.ContainsKey(aField) Then
                Return propertyInfoByValue(aField).Name
            ElseIf fieldInfoByValue.ContainsKey(aField) Then
                Return fieldInfoByValue(aField).Name
            End If
            Return Nothing
        End Function

    End Class
End Namespace