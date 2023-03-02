Imports System.Reflection

Namespace TestUtil
    ''' <summary>
    ''' 非公開のクラス・メソッド・プロパティのためのユーティリティクラス
    ''' </summary>
    ''' <remarks>外部ツール等を利用したUnitTestでどうしても非公開情報にアクセスしないとTestできない場合の利用が対象</remarks>
    Public Class NonPublicUtil

        Private Sub New()
        End Sub

        ''' <summary>
        ''' 非公開のクラス情報を取得する
        ''' </summary>
        ''' <typeparam name="TFriend">非公開クラスのAssemblyにあるクラス型（どれか一つ）</typeparam>
        ''' <param name="nonPublicClassName">非公開クラス</param>
        ''' <returns>非公開のクラス情報</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTypeOfNonPublic(Of TFriend)(ByVal nonPublicClassName As String) As Type
            Return GetTypeOfNonPublic(GetType(TFriend), nonPublicClassName)
        End Function

        ''' <summary>
        ''' 非公開のクラス情報を取得する
        ''' </summary>
        ''' <param name="friendType">非公開クラスのAssemblyにあるクラス情報（どれか一つ）</param>
        ''' <param name="nonPublicClassName">非公開クラス</param>
        ''' <returns>非公開のクラス情報</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTypeOfNonPublic(ByVal friendType As Type, ByVal nonPublicClassName As String) As Type
            Dim libAsm As Assembly = Assembly.GetAssembly(friendType)
            Dim result As Type = libAsm.GetType(nonPublicClassName)
            If result Is Nothing Then
                Throw New ArgumentException(String.Format("型 {0} のAssemblyに {1} はない", friendType.FullName, nonPublicClassName))
            End If
            Return result
        End Function

        ''' <summary>
        ''' インスタンス生成する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="parameters">コンストラクタ引数[]</param>
        ''' <returns>新インスタンス</returns>
        ''' <remarks></remarks>
        Public Shared Function NewInstance(Of T)(ByVal ParamArray parameters As Object()) As T
            Return DirectCast(NewInstance(GetType(T), parameters), T)
        End Function

        ''' <summary>
        ''' インスタンス生成する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="parameters">コンストラクタ引数[]</param>
        ''' <param name="types">コンストラクタ引数の厳密な型情報[]</param>
        ''' <returns>新インスタンス</returns>
        ''' <remarks></remarks>
        Public Shared Function NewInstance(Of T)(ByVal parameters As Object(), ByVal types As Type()) As T
            Return DirectCast(NewInstance(GetType(T), parameters, types), T)
        End Function

        ''' <summary>
        ''' インスタンス生成する
        ''' </summary>
        ''' <param name="aType">型情報</param>
        ''' <param name="parameters">コンストラクタ引数[]</param>
        ''' <returns>新インスタンス</returns>
        ''' <remarks></remarks>
        Public Shared Function NewInstance(ByVal aType As Type, ByVal ParamArray parameters As Object()) As Object
            Return NewInstance(aType, parameters, ConvTypesFromParams(parameters))
        End Function

        ''' <summary>
        ''' インスタンス生成する
        ''' </summary>
        ''' <param name="aType">型情報</param>
        ''' <param name="parameters">コンストラクタ引数[]</param>
        ''' <param name="types">コンストラクタ引数の厳密な型情報[]</param>
        ''' <returns>新インスタンス</returns>
        ''' <remarks></remarks>
        Public Shared Function NewInstance(ByVal aType As Type, ByVal parameters As Object(), ByVal types As Type()) As Object
            Dim constructor As ConstructorInfo = aType.GetConstructor(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance, Nothing, types, Nothing)
            If constructor Is Nothing Then
                Throw New ArgumentException(String.Format("型 {0} に.ctor({1}) は無い", aType.FullName, Join(ConvTypeNames(types), ",")))
            End If
            If constructor.IsPublic AndAlso (aType.IsPublic OrElse (aType.IsNested AndAlso aType.IsNestedPublic)) Then
                Throw New InvalidOperationException(String.Format("Publicクラス'{0}' でPublic .ctor({1}) なら正規にインスタンス生成すべき", aType.FullName, Join(ConvTypeNames(types), ",")))
            End If
            Return constructor.Invoke(parameters)
        End Function

        Private Shared Function ConvTypeNames(ByVal types As IEnumerable(Of Type)) As String()
            Dim typeNames As New List(Of String)
            For Each aType2 As Type In types
                typeNames.Add(aType2.FullName)
            Next
            Return typeNames.ToArray
        End Function

        Private Shared Function ConvTypesFromParams(ByVal parameters As IEnumerable(Of Object)) As Type()
            Dim types As New List(Of Type)
            For Each parameter As Object In parameters
                If parameter Is Nothing Then
                    Throw New InvalidOperationException("コンストラクタ引数の型を推測できない. 型情報を追加してください.")
                End If
                types.Add(parameter.GetType)
            Next
            Return types.ToArray
        End Function

        ''' <summary>
        ''' Field値を取得する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="instance">設定先のインスタンス</param>
        ''' <param name="fieldName">設定先のメンバー名</param>
        ''' <remarks></remarks>
        Public Shared Function GetField(Of T)(ByVal instance As T, ByVal fieldName As String) As Object
            Return GetField(DetectType(Of T)(instance), instance, fieldName)
        End Function

        ''' <summary>
        ''' Field値を取得する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="instance">設定先のインスタンス</param>
        ''' <param name="fieldName">設定先のメンバー名</param>
        ''' <remarks></remarks>
        Public Shared Function GetField(ByVal aType As Type, ByVal instance As Object, ByVal fieldName As String) As Object
            Return PerformGetField(aType, instance, fieldName, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.FlattenHierarchy)
        End Function

        ''' <summary>
        ''' Field値を取得する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="fieldName">設定先のメンバー名</param>
        ''' <remarks></remarks>
        Public Shared Function GetFieldOfShared(Of T)(ByVal fieldName As String) As Object
            Return GetFieldOfShared(GetType(T), fieldName)
        End Function

        ''' <summary>
        ''' Field値を取得する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="fieldName">設定先のメンバー名</param>
        ''' <remarks></remarks>
        Public Shared Function GetFieldOfShared(ByVal aType As Type, ByVal fieldName As String) As Object
            Return PerformGetField(aType, Nothing, fieldName, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.FlattenHierarchy)
        End Function

        ''' <summary>
        ''' Fieldに値を設定する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="instance">設定先のインスタンス ※Sharedの場合はnull</param>
        ''' <param name="fieldName">設定先のメンバー名</param>
        ''' <param name="aBindingFlags">Fieldの状態</param>
        ''' <remarks></remarks>
        Private Shared Function PerformGetField(ByVal aType As Type, ByVal instance As Object, ByVal fieldName As String, ByVal aBindingFlags As BindingFlags) As Object
            Dim field As FieldInfo = RecurGetField(aType, fieldName, aBindingFlags)
            If field Is Nothing Then
                Throw New ArgumentException(String.Format("型 {0} に{1}メンバー '{2}' は無い", aType.FullName, If((aBindingFlags And BindingFlags.Static) = BindingFlags.Static, "Shared", ""), fieldName))
            End If
            If field.IsPublic AndAlso (aType.IsPublic OrElse (aType.IsNested AndAlso aType.IsNestedPublic)) Then
                Throw New InvalidOperationException(String.Format("Publicクラス'{0}' でPublicメンバー '{1}' なら正規に値を取得すべき", aType.FullName, fieldName))
            End If
            Return field.GetValue(instance)
        End Function

        ''' <summary>
        ''' Fieldに値を設定する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="instance">設定先のインスタンス</param>
        ''' <param name="fieldName">設定先のメンバー名</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Shared Sub SetField(Of T)(ByVal instance As T, ByVal fieldName As String, ByVal value As Object)
            SetField(DetectType(instance), instance, fieldName, value)
        End Sub

        ''' <summary>
        ''' Fieldに値を設定する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="instance">設定先のインスタンス</param>
        ''' <param name="fieldName">設定先のメンバー名</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Shared Sub SetField(ByVal aType As Type, ByVal instance As Object, ByVal fieldName As String, ByVal value As Object)
            PerformSetField(aType, instance, fieldName, value, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.FlattenHierarchy)
        End Sub

        ''' <summary>
        ''' 静的Fieldに値を設定する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="sharedFieldName">設定先の静的メンバー名</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Shared Sub SetFieldOfShared(Of T)(ByVal sharedFieldName As String, ByVal value As Object)
            SetFieldOfShared(GetType(T), sharedFieldName, value)
        End Sub

        '''' <summary>
        '''' 非公開の静的Fieldに値を設定する
        '''' </summary>
        '''' <param name="aType">型</param>
        '''' <param name="sharedPropertyName">設定先の静的メンバー名</param>
        '''' <param name="value">値</param>
        '''' <remarks></remarks>
        'Public Shared Sub SetFieldOfShared(ByVal aType As Type, ByVal sharedPropertyName As String, ByVal value As Object)
        '    PerformSetField(aType, Nothing, sharedPropertyName, value, BindingFlags.NonPublic Or BindingFlags.Static)
        'End Sub

        ''' <summary>
        ''' 静的Fieldに値を設定する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="sharedFieldName">設定先の静的メンバー名</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Shared Sub SetFieldOfShared(ByVal aType As Type, ByVal sharedFieldName As String, ByVal value As Object)
            PerformSetField(aType, Nothing, sharedFieldName, value, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.FlattenHierarchy)
        End Sub

        ''' <summary>
        ''' Fieldに値を設定する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="instance">設定先のインスタンス ※Sharedの場合はnull</param>
        ''' <param name="fieldName">設定先のメンバー名</param>
        ''' <param name="value">値</param>
        ''' <param name="aBindingFlags">Fieldの状態</param>
        ''' <remarks></remarks>
        Private Shared Sub PerformSetField(ByVal aType As Type, ByVal instance As Object, ByVal fieldName As String, ByVal value As Object, ByVal aBindingFlags As BindingFlags)
            Dim field As FieldInfo = RecurGetField(aType, fieldName, aBindingFlags)
            If field Is Nothing Then
                Throw New ArgumentException(String.Format("型 {0} に{1}メンバー '{2}' は無い", aType.FullName, If((aBindingFlags And BindingFlags.Static) = BindingFlags.Static, "Shared", ""), fieldName))
            End If
            If field.IsPublic AndAlso (aType.IsPublic OrElse (aType.IsNested AndAlso aType.IsNestedPublic)) Then
                Throw New InvalidOperationException(String.Format("Publicクラス'{0}' でPublicメンバー '{1}' なら正規に値を設定すべき", aType.FullName, fieldName))
            End If
            field.SetValue(instance, value)
        End Sub

        Private Shared Function RecurGetField(ByVal aType As Type, ByVal fieldName As String, ByVal aBindingFlags As BindingFlags) As FieldInfo
            Dim field As FieldInfo = aType.GetField(fieldName, aBindingFlags)
            If field Is Nothing Then
                If TypeUtil.TypeObject IsNot aType.BaseType _
                        AndAlso (aBindingFlags And BindingFlags.FlattenHierarchy) = BindingFlags.FlattenHierarchy Then
                    ' FlattenHierarchy 指定があれば親クラスのFieldも参照するが、Privateスコープは参照できない
                    ' ここでは FlattenHierarchy 指定があればPrivateスコープでも参照できるよう救済する
                    Return RecurGetField(aType.BaseType, fieldName, aBindingFlags)
                End If
                Return Nothing
            End If
            Return field
        End Function

        Private Shared Function DetectType(Of T)(ByVal instance As T) As Type
            Dim tType As Type = GetType(T)
            Dim instanceType As Type = If(instance IsNot Nothing, instance.GetType, GetType(Object))
            If tType.IsAssignableFrom(instanceType) Then
                Return instanceType
            End If
            Return tType
        End Function

        ''' <summary>
        ''' Property値を取得する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="instance">取得先のインスタンス</param>
        ''' <param name="propertyName">取得するプロパティ名</param>
        ''' <remarks></remarks>
        Public Shared Function GetProperty(Of T)(ByVal instance As T, ByVal propertyName As String) As Object
            Return GetProperty(DetectType(Of T)(instance), instance, propertyName)
        End Function

        ''' <summary>
        ''' Property値を取得する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="instance">取得先のインスタンス</param>
        ''' <param name="propertyName">取得するプロパティ名</param>
        ''' <remarks></remarks>
        Public Shared Function GetProperty(ByVal aType As Type, ByVal instance As Object, ByVal propertyName As String) As Object
            Return PerformGetProperty(aType, instance, propertyName, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.FlattenHierarchy)
        End Function

        ''' <summary>
        ''' 静的Property値を取得する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="sharedPropertyName">取得する静的プロパティ名</param>
        ''' <remarks></remarks>
        Public Shared Function GetPropertyOfShared(Of T)(ByVal sharedPropertyName As String) As Object
            Return GetPropertyOfShared(GetType(T), sharedPropertyName)
        End Function

        ''' <summary>
        ''' 静的Property値を取得する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="sharedPropertyName">取得する静的プロパティ名</param>
        ''' <remarks></remarks>
        Public Shared Function GetPropertyOfShared(ByVal aType As Type, ByVal sharedPropertyName As String) As Object
            Return PerformGetProperty(aType, Nothing, sharedPropertyName, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.FlattenHierarchy)
        End Function

        ''' <summary>
        ''' Property値を取得する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="instance">取得先のインスタンス ※Sharedの場合はnull</param>
        ''' <param name="propertyName">取得するプロパティ名</param>
        ''' <param name="aBindingFlags">Propertyの状態</param>
        ''' <remarks></remarks>
        Private Shared Function PerformGetProperty(ByVal aType As Type, ByVal instance As Object, ByVal propertyName As String, ByVal aBindingFlags As BindingFlags) As Object
            Dim aPropertyInfo As PropertyInfo = RecurGetProperty(aType, propertyName, aBindingFlags Or BindingFlags.GetProperty)
            If aPropertyInfo Is Nothing Then
                Throw New ArgumentException(String.Format("型 {0} に{1}プロパティ '{2}' は無い", aType.FullName, If((aBindingFlags And BindingFlags.Static) = BindingFlags.Static, "Shared", ""), propertyName))
            End If
            Dim getMethod As MethodInfo = aPropertyInfo.GetGetMethod(True)
            If getMethod Is Nothing Then
                Throw New ArgumentException(String.Format("型 {0} のプロパティ '{1}' のGetPropertyがない", aType.FullName, propertyName))
            ElseIf getMethod.IsPublic AndAlso (aType.IsPublic OrElse (aType.IsNested AndAlso aType.IsNestedPublic)) Then
                Throw New InvalidOperationException(String.Format("Publicクラス'{0}' でPublicプロパティ '{1}' なら正規に値を取得すべき", aType.FullName, propertyName))
            End If
            Return getMethod.Invoke(instance, Nothing)
        End Function

        ''' <summary>
        ''' Propertyに値を設定する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="instance">設定先のインスタンス</param>
        ''' <param name="propertyName">設定先のプロパティ名</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Shared Sub SetProperty(Of T)(ByVal instance As T, ByVal propertyName As String, ByVal value As Object)
            PerformSetProperty(DetectType(Of T)(instance), instance, propertyName, value, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.FlattenHierarchy)
        End Sub

        ''' <summary>
        ''' 静的Propertyに値を設定する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="sharedPropertyName">設定先の静的プロパティ名</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Shared Sub SetPropertyOfShared(Of T)(ByVal sharedPropertyName As String, ByVal value As Object)
            SetPropertyOfShared(GetType(T), sharedPropertyName, value)
        End Sub

        ''' <summary>
        ''' 静的Propertyに値を設定する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="sharedPropertyName">設定先の静的プロパティ名</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Shared Sub SetPropertyOfShared(ByVal aType As Type, ByVal sharedPropertyName As String, ByVal value As Object)
            PerformSetProperty(aType, Nothing, sharedPropertyName, value, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.FlattenHierarchy)
        End Sub

        ''' <summary>
        ''' Propetyに値を設定する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="instance">設定先のインスタンス ※Sharedの場合はnull</param>
        ''' <param name="propertyName">設定先のプロパティ名</param>
        ''' <param name="value">値</param>
        ''' <param name="aBindingFlags">Propertyの状態</param>
        ''' <remarks></remarks>
        Private Shared Sub PerformSetProperty(ByVal aType As Type, ByVal instance As Object, ByVal propertyName As String, ByVal value As Object, ByVal aBindingFlags As BindingFlags)
            Dim aPropertyInfo As PropertyInfo = RecurGetProperty(aType, propertyName, aBindingFlags Or BindingFlags.SetProperty)
            If aPropertyInfo Is Nothing Then
                Throw New ArgumentException(String.Format("型 {0} に{1}プロパティ '{2}' は無い", aType.FullName, If((aBindingFlags And BindingFlags.Static) = BindingFlags.Static, "Shared", ""), propertyName))
            End If
            Dim setMethod As MethodInfo = aPropertyInfo.GetSetMethod(True)
            If setMethod Is Nothing Then
                Throw New ArgumentException(String.Format("型 {0} のプロパティ '{1}' のSetPropertyがない", aType.FullName, propertyName))
            ElseIf setMethod.IsPublic AndAlso (aType.IsPublic OrElse (aType.IsNested AndAlso aType.IsNestedPublic)) Then
                Throw New InvalidOperationException(String.Format("Publicクラス'{0}' でPublicプロパティ '{1}' なら正規に値を設定すべき", aType.FullName, propertyName))
            End If
            setMethod.Invoke(instance, New Object() {value})
        End Sub

        Private Shared Function RecurGetProperty(ByVal aType As Type, ByVal propertyName As String, ByVal aBindingFlags As BindingFlags) As PropertyInfo
            Dim aProperty As PropertyInfo = aType.GetProperty(propertyName, aBindingFlags)
            If aProperty Is Nothing Then
                If TypeUtil.TypeObject IsNot aType.BaseType _
                        AndAlso (aBindingFlags And BindingFlags.FlattenHierarchy) = BindingFlags.FlattenHierarchy Then
                    ' FlattenHierarchy 指定があれば親クラスのFieldも参照するが、Privateスコープは参照できない
                    ' ここでは FlattenHierarchy 指定があればPrivateスコープでも参照できるよう救済する
                    Return RecurGetProperty(aType.BaseType, propertyName, aBindingFlags)
                End If
                Return Nothing
            End If
            Return aProperty
        End Function

        ''' <summary>
        ''' メソッドを実行する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="instance">実行メソッドを持つインスタンス ※Sharedの場合はnull</param>
        ''' <param name="methodName">メソッド名</param>
        ''' <param name="parameters">引数[]</param>
        ''' <returns>メソッドの実行結果</returns>
        ''' <remarks></remarks>
        Public Shared Function InvokeMethod(Of T)(ByVal instance As T, ByVal methodName As String, ByVal ParamArray parameters As Object()) As Object
            Return InvokeMethod(DetectType(Of T)(instance), instance, methodName, parameters)
        End Function

        ''' <summary>
        ''' メソッドを実行する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="instance">実行メソッドを持つインスタンス ※Sharedの場合はnull</param>
        ''' <param name="methodName">メソッド名</param>
        ''' <param name="parameters">引数[]</param>
        ''' <param name="types">引数の型[] ※null可.但し予期せぬメソッド呼出の場合あり</param>
        ''' <returns>メソッドの実行結果</returns>
        ''' <remarks></remarks>
        Public Shared Function InvokeMethod(Of T)(ByVal instance As T, ByVal methodName As String, ByVal parameters As Object(), ByVal types As Type()) As Object
            Return InvokeMethod(DetectType(Of T)(instance), instance, methodName, parameters, types)
        End Function

        ''' <summary>
        ''' メソッドを実行する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="instance">実行メソッドを持つインスタンス ※Sharedの場合はnull</param>
        ''' <param name="methodName">メソッド名</param>
        ''' <param name="parameters">引数[]</param>
        ''' <returns>メソッドの実行結果</returns>
        ''' <remarks></remarks>
        Public Shared Function InvokeMethod(ByVal aType As Type, ByVal instance As Object, ByVal methodName As String, ByVal ParamArray parameters As Object()) As Object
            Return PerformInvokeMethod(aType, instance, methodName, parameters, Nothing, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
        End Function

        ''' <summary>
        ''' メソッドを実行する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="instance">実行メソッドを持つインスタンス ※Sharedの場合はnull</param>
        ''' <param name="methodName">メソッド名</param>
        ''' <param name="parameters">引数[]</param>
        ''' <param name="types">引数の型[] ※null可.但し予期せぬメソッド呼出の場合あり</param>
        ''' <returns>メソッドの実行結果</returns>
        ''' <remarks></remarks>
        Public Shared Function InvokeMethod(ByVal aType As Type, ByVal instance As Object, ByVal methodName As String, ByVal parameters As Object(), ByVal types As Type()) As Object
            Return PerformInvokeMethod(aType, instance, methodName, parameters, types, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
        End Function

        ''' <summary>
        ''' 静的メソッドを実行する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="methodName">メソッド名</param>
        ''' <param name="parameters">引数[]</param>
        ''' <returns>メソッドの実行結果</returns>
        ''' <remarks></remarks>
        Public Shared Function InvokeMethodOfShared(Of T)(ByVal methodName As String, ByVal ParamArray parameters As Object()) As Object
            Return InvokeMethodOfShared(GetType(T), methodName, parameters)
        End Function

        ''' <summary>
        ''' 静的メソッドを実行する
        ''' </summary>
        ''' <typeparam name="T">型</typeparam>
        ''' <param name="methodName">メソッド名</param>
        ''' <param name="parameters">引数[]</param>
        ''' <param name="types">引数の型[] ※null可.但し予期せぬメソッド呼出の場合あり</param>
        ''' <returns>メソッドの実行結果</returns>
        ''' <remarks></remarks>
        Public Shared Function InvokeMethodOfShared(Of T)(ByVal methodName As String, ByVal parameters As Object(), ByVal types As Type()) As Object
            Return InvokeMethodOfShared(GetType(T), methodName, parameters, types)
        End Function

        ''' <summary>
        ''' 静的メソッドを実行する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="methodName">メソッド名</param>
        ''' <param name="parameters">引数[]</param>
        ''' <returns>メソッドの実行結果</returns>
        ''' <remarks></remarks>
        Public Shared Function InvokeMethodOfShared(ByVal aType As Type, ByVal methodName As String, ByVal ParamArray parameters As Object()) As Object
            Return PerformInvokeMethod(aType, Nothing, methodName, parameters, Nothing, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static)
        End Function

        ''' <summary>
        ''' 静的メソッドを実行する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="methodName">メソッド名</param>
        ''' <param name="parameters">引数[]</param>
        ''' <param name="types">引数の型[] ※null可.但し予期せぬメソッド呼出の場合あり</param>
        ''' <returns>メソッドの実行結果</returns>
        ''' <remarks></remarks>
        Public Shared Function InvokeMethodOfShared(ByVal aType As Type, ByVal methodName As String, ByVal parameters As Object(), ByVal types As Type()) As Object
            Return PerformInvokeMethod(aType, Nothing, methodName, parameters, types, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static)
        End Function

        ''' <summary>
        ''' メソッドを実行する
        ''' </summary>
        ''' <param name="aType">型</param>
        ''' <param name="instance">実行メソッドを持つインスタンス ※Sharedの場合はnull</param>
        ''' <param name="methodName">メソッド名</param>
        ''' <param name="parameters">引数[]</param>
        ''' <param name="types">引数の型[] ※null可.但し予期せぬメソッド呼出の場合あり</param>
        ''' <param name="aBindingFlags">メソッドの状態</param>
        ''' <returns>メソッドの実行結果</returns>
        ''' <remarks></remarks>
        Private Shared Function PerformInvokeMethod(ByVal aType As Type, ByVal instance As Object, ByVal methodName As String, ByVal parameters As Object(), ByVal types As Type(), ByVal aBindingFlags As BindingFlags) As Object
            Dim method As MethodInfo
            If types Is Nothing Then
                method = aType.GetMethod(methodName, aBindingFlags)
            Else
                method = aType.GetMethod(methodName, aBindingFlags, Nothing, types, Nothing)
            End If
            If method Is Nothing Then
                Throw New ArgumentException(String.Format("型 {0} の{1}メソッド '{2}' はない", aType.FullName, If((aBindingFlags And BindingFlags.Static) = BindingFlags.Static, "Shared", ""), methodName))
            ElseIf method.IsPublic AndAlso (aType.IsPublic OrElse (aType.IsNested AndAlso aType.IsNestedPublic)) Then
                Throw New InvalidOperationException(String.Format("Publicクラス'{0}' でPublic{1}メソッド '{2}' なら正規に値を設定すべき", aType.FullName, If((aBindingFlags And BindingFlags.Static) = BindingFlags.Static, "Shared", ""), methodName))
            End If
            Return method.Invoke(instance, parameters)
        End Function

    End Class
End Namespace