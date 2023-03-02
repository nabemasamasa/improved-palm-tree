Public Class EnumUtil

    ''' <summary>
    ''' Enum名をEnum値にする
    ''' </summary>
    ''' <typeparam name="TEnum">Enum値の型</typeparam>
    ''' <param name="enumName">Enum名</param>
    ''' <param name="ignoresCase">大文字小文字同一視するならtrue</param>
    ''' <returns>Enum値</returns>
    ''' <remarks></remarks>
    Public Shared Function ParseByName(Of TEnum As Structure)(enumName As String, Optional ignoresCase As Boolean = False) As TEnum
        AssertEnum(Of TEnum)()
        If Not [Enum].IsDefined(GetType(TEnum), enumName) Then
            If ignoresCase Then
                For Each name As String In GetNames(Of TEnum)()
                    If name.Equals(enumName, StringComparison.OrdinalIgnoreCase) Then
                        Return ParseByName(Of TEnum)(name, ignoresCase:=False)
                    End If
                Next
            End If
            Throw New ArgumentException(String.Format("Enum {0} に'{1}'はありません", GetType(TEnum).Name, enumName))
        End If
        Return DirectCast([Enum].Parse(GetType(TEnum), enumName), TEnum)
    End Function

    ''' <summary>
    ''' 値をEnum値にする
    ''' </summary>
    ''' <typeparam name="TEnum">Enum値の型</typeparam>
    ''' <param name="enumValue">値</param>
    ''' <param name="ignoresDefine">未定義の値でもEnum値にするならtrue</param>
    ''' <returns>Enum値</returns>
    ''' <remarks></remarks>
    Public Shared Function ParseByValue(Of TEnum As Structure)(enumValue As Integer, Optional ignoresDefine As Boolean = False) As TEnum
        AssertEnum(Of TEnum)()
        If Not [Enum].IsDefined(GetType(TEnum), enumValue) AndAlso Not ignoresDefine Then
            Throw New ArgumentException(String.Format("Enum {0} に {1} はありません", GetType(TEnum).Name, enumValue))
        End If
        Return DirectCast([Enum].ToObject(GetType(TEnum), enumValue), TEnum)
    End Function

    ''' <summary>
    ''' 値をEnum値にする
    ''' </summary>
    ''' <typeparam name="TEnum">Enum値の型</typeparam>
    ''' <param name="enumValue">値</param>
    ''' <param name="ignoresDefine">未定義の値でもEnum値にするならtrue</param>
    ''' <returns>Enum値</returns>
    ''' <remarks></remarks>
    Public Shared Function ParseByNullableValue(Of TEnum As Structure)(enumValue As Integer?, Optional ignoresDefine As Boolean = False) As TEnum?
        If Not enumValue.HasValue Then
            Return Nothing
        End If
        Return ParseByValue(Of TEnum)(enumValue.Value, ignoresDefine)
    End Function

    ''' <summary>
    ''' Enum値の名前を取得する
    ''' </summary>
    ''' <typeparam name="TEnum">Enum型</typeparam>
    ''' <param name="enumValue">Enum値</param>
    ''' <returns>Enum名</returns>
    ''' <remarks></remarks>
    Public Shared Function GetName(Of TEnum As Structure)(enumValue As Object) As String
        AssertEnum(Of TEnum)()
        Return [Enum].GetName(GetType(TEnum), enumValue)
    End Function

    ''' <summary>
    ''' Enumに定義した名前をすべて取得する
    ''' </summary>
    ''' <typeparam name="TEnum">Enum型</typeparam>
    ''' <returns>Enum名[]</returns>
    ''' <remarks></remarks>
    Public Shared Function GetNames(Of TEnum As Structure)() As String()
        AssertEnum(Of TEnum)()
        Return [Enum].GetNames(GetType(TEnum))
    End Function

    ''' <summary>
    ''' Enumに定義した値をすべて取得する
    ''' </summary>
    ''' <typeparam name="TEnum">Enum型</typeparam>
    ''' <returns>Enum値[]</returns>
    ''' <remarks></remarks>
    Public Shared Function GetValues(Of TEnum As Structure)() As TEnum()
        AssertEnum(Of TEnum)()
        Return DirectCast([Enum].GetValues(GetType(TEnum)), TEnum())
    End Function

    ''' <summary>
    ''' 値が含まれるか？を返す
    ''' </summary>
    ''' <typeparam name="TEnum">Enum型</typeparam>
    ''' <param name="setEnum">集合したEnum値</param>
    ''' <param name="enumValue">Enum値</param>
    ''' <returns>判定結果</returns>
    ''' <remarks></remarks>
    Public Shared Function Contains(Of TEnum As Structure)(setEnum As TEnum, enumValue As TEnum) As Boolean
        AssertEnum(Of TEnum)()
        Dim setVal As Integer = ConvertGenericToInt(setEnum)
        Dim valueVal As Integer = ConvertGenericToInt(enumValue)
        Return (setVal And valueVal) = valueVal
    End Function

    Private Shared Sub AssertEnum(Of TEnum As Structure)()
        Dim type As Type = GetType(TEnum)
        If Not type.IsEnum Then
            Throw New ArgumentException(String.Format("型引数:{0}はEnum型ではない", type.Name))
        End If
    End Sub

    Private Shared Function ConvertGenericToInt(Of T)(value As T) As Integer
        Return CInt(DirectCast(value, Object))
    End Function

End Class
