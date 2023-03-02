Namespace Sys
    Public Class PrivateProfileHelper

        Private ReadOnly profile As PrivateProfile

        Public Sub New(ByVal profile As PrivateProfile)
            Me.profile = profile
        End Sub

        Private Const ARRAY_SEPARATOR As String = "/"
        Private Const EMPTY_ELEMENT_VALUE_OF_ARRAY As String = "__EMPTY__"
        Private Const PER_FACTOR As Decimal = 1000D

        ''' <summary>
        ''' 配列値をiniファイル保存用の文字列にして返す
        ''' </summary>
        ''' <param name="values">配列値</param>
        ''' <returns>iniファイル保存用の文字列</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvArrayToIniValue(ByVal values As String()) As String
            If values Is Nothing Then
                Return ""
            End If
            Dim values2 As New List(Of String)
            For Each val As String In values
                values2.Add(If(StringUtil.IsEmpty(val), EMPTY_ELEMENT_VALUE_OF_ARRAY, val))
            Next
            Return Join(values2.ToArray, ARRAY_SEPARATOR)
        End Function

        ''' <summary>
        ''' iniファイルの値を配列値にして返す
        ''' </summary>
        ''' <param name="iniValue">iniファイル値</param>
        ''' <returns>配列値</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvIniValueToArray(ByVal iniValue As String) As String()
            If StringUtil.IsEmpty(iniValue) Then
                Return New String() {}
            End If
            Dim result As New List(Of String)
            For Each val As String In Split(iniValue, ARRAY_SEPARATOR)
                result.Add(If(EMPTY_ELEMENT_VALUE_OF_ARRAY.Equals(val), "", val))
            Next
            Return result.ToArray
        End Function

        Private Const DICTIONARY_KEY_VALUE_SEPARATOR As String = "="
        Private Const DICTIONARY_SEPARATOR As String = ","

        ''' <summary>
        ''' Dictionary型のをiniファイル保存用の文字列にして返す
        ''' </summary>
        ''' <param name="values">Dictionary型の値</param>
        ''' <returns>iniファイル保存用の文字列</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvDictionaryToIniValue(ByVal values As Dictionary(Of String, String)) As String
            If values Is Nothing Then
                Return ""
            End If
            Dim result As New List(Of String)
            For Each pair As KeyValuePair(Of String, String) In values
                result.Add(pair.Key & DICTIONARY_KEY_VALUE_SEPARATOR & pair.Value)
            Next
            Return Join(result.ToArray, DICTIONARY_SEPARATOR)
        End Function

        ''' <summary>
        ''' iniファイルの値をDictionary型の値にして返す
        ''' </summary>
        ''' <param name="iniValue">iniファイル値</param>
        ''' <returns>Dictionary型の値</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvIniValueToDictionary(ByVal iniValue As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)
            If StringUtil.IsEmpty(iniValue) Then
                Return result
            End If
            For Each pair As String In Split(iniValue, DICTIONARY_SEPARATOR)
                Dim strings As String() = Split(pair, DICTIONARY_KEY_VALUE_SEPARATOR)
                If strings.Length = 0 Then
                    Continue For
                End If
                If strings.Length = 1 Then
                    result.Add(strings(0), String.Empty)
                End If
                result.Add(strings(0), strings(1))
            Next
            Return result
        End Function

        ''' <summary>
        ''' 値を設定する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Sub SetValue(ByVal section As String, ByVal key As String, ByVal value As Object)

            profile.Write(section, key, StringUtil.ToString(value))
        End Sub

        ''' <summary>
        ''' 配列値を設定する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <param name="values">配列値</param>
        ''' <remarks></remarks>
        Public Sub SetValues(ByVal section As String, ByVal key As String, ByVal values As Object())

            If values Is Nothing Then
                profile.Write(section, key, String.Empty)
                Return
            End If
            Dim iniValues As New List(Of String)
            For Each value As Object In values
                If value Is Nothing Then
                    iniValues.Add(String.Empty)
                    Continue For
                End If
                iniValues.Add(value.ToString)
            Next
            profile.Write(section, key, ConvArrayToIniValue(iniValues.ToArray))
        End Sub

        ''' <summary>
        ''' 値を設定する
        ''' </summary>
        ''' <typeparam name="T">設定値の型</typeparam>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Sub SetValue(Of T As Structure)(ByVal section As String, ByVal key As String, ByVal value As Nullable(Of T))

            If value Is Nothing Then
                profile.Write(section, key, String.Empty)
                Return
            End If
            Dim obj As Object = value.Value
            profile.Write(section, key, CStr(obj))
        End Sub

        ''' <summary>
        ''' 配列値を設定する
        ''' </summary>
        ''' <typeparam name="T">設定値の型</typeparam>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <param name="values">配列値</param>
        ''' <remarks></remarks>
        Public Sub SetValues(Of T As Structure)(ByVal section As String, ByVal key As String, ByVal values As ICollection(Of Nullable(Of T)))

            If values Is Nothing Then
                profile.Write(section, key, String.Empty)
                Return
            End If
            Dim iniValues As New List(Of String)
            For Each value As Nullable(Of T) In values
                If value Is Nothing Then
                    iniValues.Add(String.Empty)
                    Continue For
                End If
                iniValues.Add(value.ToString)
            Next
            profile.Write(section, key, ConvArrayToIniValue(iniValues.ToArray))
        End Sub

        ''' <summary>
        ''' パーセント値を設定する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <param name="value">パーセント値</param>
        ''' <remarks></remarks>
        Public Sub SetValueOfPercent(ByVal section As String, ByVal key As String, ByVal value As Decimal?)

            If value Is Nothing Then
                profile.Write(section, key, "")
                Return
            End If
            profile.Write(section, key, CStr(Decimal.Floor(value.Value * PER_FACTOR)))
        End Sub

        ''' <summary>
        ''' 値を取得する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <returns>取得した値</returns>
        ''' <remarks></remarks>
        Public Function GetValue(ByVal section As String, ByVal key As String) As String
            Return profile.Read(section, key)
        End Function

        ''' <summary>
        ''' 配列値を取得する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <returns>配列値</returns>
        ''' <remarks></remarks>
        Public Function GetValues(ByVal section As String, ByVal key As String) As String()
            Return ConvIniValueToArray(profile.Read(section, key))
        End Function

        ''' <summary>
        ''' Int値を取得する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Public Function GetValueAsInt(ByVal section As String, ByVal key As String) As Integer?
            Dim value As String = profile.Read(section, key)
            If Not IsNumeric(value) Then
                Return Nothing
            End If
            Return CInt(value)
        End Function

        ''' <summary>
        ''' Long値を取得する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Public Function GetValueAsLong(ByVal section As String, ByVal key As String) As Long?
            Dim value As String = profile.Read(section, key)
            If Not IsNumeric(value) Then
                Return Nothing
            End If
            Return CLng(value)
        End Function

        ''' <summary>
        ''' 配列のInt値を取得する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <returns>Int値の配列</returns>
        ''' <remarks></remarks>
        Public Function GetValuesAsInt(ByVal section As String, ByVal key As String) As Integer?()
            Dim iniValue As String = profile.Read(section, key)
            If iniValue Is Nothing Then
                Return New Integer?() {}
            End If
            Dim values As New List(Of Integer?)
            For Each value As String In ConvIniValueToArray(iniValue)
                If Not IsNumeric(value) Then
                    values.Add(Nothing)
                    Continue For
                End If
                values.Add(CInt(value))
            Next
            Return values.ToArray
        End Function

        ''' <summary>
        ''' Date値を取得する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Public Function GetValueAsDate(ByVal section As String, ByVal key As String) As DateTime?
            Dim value As String = profile.Read(section, key)
            If Not IsDate(value) Then
                Return Nothing
            End If
            Return CDate(value)
        End Function

        ''' <summary>
        ''' Boolean値を取得する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Public Function GetValueAsBool(ByVal section As String, ByVal key As String) As Boolean?

            Dim value As String = profile.Read(section, key)
            If Not EzUtil.IsBooleanValue(value) Then
                Return Nothing
            End If
            Return EzUtil.IsTrue(value)
        End Function

        ''' <summary>
        ''' パーセント値を取得する
        ''' </summary>
        ''' <param name="section">セクション</param>
        ''' <param name="key">キー</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Public Function GetValueAsPercent(ByVal section As String, ByVal key As String) As Decimal?

            Dim value As String = profile.Read(section, key)
            If Not IsNumeric(value) Then
                Return Nothing
            End If
            Return CDec(value) / PER_FACTOR
        End Function

    End Class
End Namespace