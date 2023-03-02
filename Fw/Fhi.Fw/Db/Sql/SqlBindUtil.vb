Imports System.Text
Imports System.Text.RegularExpressions

Namespace Db.Sql
    ''' <summary>
    ''' ADO.NETのバインド変数を担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SqlBindUtil
        ''' <summary>バインド変数名の先頭文字</summary>
        Private Const BIND_NAME_PREFIX As Char = "@"c
        ''' <summary>（内部的にIndexプロパティを表す）バインド変数名の区切り文字</summary>
        Private Const INTERNAL_BIND_NAME_INDEX_SEPARATOR As Char = "#"c
        ''' <summary>（内部的）バインド変数名の区切り文字</summary>
        Private Const INTERNAL_BIND_NAME_PROPERTY_SEPARATOR As Char = "$"c
        ''' <summary>Indexプロパティ名の開始文字</summary>
        Private Const INDEX_PROPERTY_OPEN As Char = "("c
        ''' <summary>Indexプロパティ名の終了文字</summary>
        Private Const INDEX_PROPERTY_CLOSE As String = ")"
        ''' <summary>プロパティ（アクセス）名の区切り文字</summary>
        Private Const PROPERTY_SEPARATOR As String = "."
        ''' <summary>プロパティ（アクセス）名の区切り文字</summary>
        Private Const PROPERTY_SEPARATOR_CHAR As Char = "."c

        ''' <summary>非公開コンストラクタ</summary>
        Private Sub New()
        End Sub

        ''' <summary>
        ''' プロパティ（アクセス）名を、（内部的な）バインド変数名にして返す
        ''' </summary>
        ''' <param name="propertyName">プロパティ（アクセス）名</param>
        ''' <returns>バインド変数名</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvPropertyNameToInternalBindName(ByVal propertyName As String) As String
            Return ConvPropertyNameToInternalBindName(propertyName, Nothing)
        End Function

        ''' <summary>
        ''' プロパティ（アクセス）名を、（内部的な）バインド変数名にして返す
        ''' </summary>
        ''' <param name="propertyName">プロパティ（アクセス）名</param>
        ''' <param name="index">index</param>
        ''' <returns>バインド変数名</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvPropertyNameToInternalBindName(ByVal propertyName As String, ByVal index As Integer?) As String
            Dim result As New StringBuilder
            If Not propertyName.StartsWith(BIND_NAME_PREFIX) Then
                result.Append(BIND_NAME_PREFIX)
            End If
            Dim propertyResult As New List(Of String)
            For Each part As String In Split(propertyName, PROPERTY_SEPARATOR)
                If 0 <= part.IndexOf(INDEX_PROPERTY_OPEN) And part.EndsWith(INDEX_PROPERTY_CLOSE) Then
                    propertyResult.Add(part.Replace(INDEX_PROPERTY_OPEN, INTERNAL_BIND_NAME_INDEX_SEPARATOR).Substring(0, part.Length - 1))
                Else
                    propertyResult.Add(part)
                End If
            Next
            result.Append(Join(propertyResult.ToArray, INTERNAL_BIND_NAME_PROPERTY_SEPARATOR))
            If index Is Nothing Then
                Return result.ToString
            End If
            Return result.Append(INTERNAL_BIND_NAME_INDEX_SEPARATOR).Append(index.Value).ToString()
        End Function

        ''' <summary>
        ''' ユーザー入力バインド変数名を内部的なバインド変数名にする
        ''' </summary>
        ''' <param name="userBindName">バインド変数</param>
        ''' <returns>内部的なバインド変数名</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvBindNameUserToInternal(ByVal userBindName As String) As String
            AssertBindNameStartPrefix(userBindName)
            Return ConvPropertyNameToInternalBindName(userBindName.Substring(1))
        End Function

        ''' <summary>
        ''' 内部的なバインド変数名を、ユーザー入力バインド変数名にする
        ''' </summary>
        ''' <param name="internalBindName">内部的なバインド変数名</param>
        ''' <returns>ユーザー入力バインド変数名</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvBindNameInternalToUser(ByVal internalBindName As String) As String
            Return BIND_NAME_PREFIX & ConvBindNameToPropertyName(internalBindName)
        End Function

        ''' <summary>
        ''' バインド変数名が@から始まっている事を保証する
        ''' </summary>
        ''' <param name="bindName">バインド変数名</param>
        ''' <remarks></remarks>
        Private Shared Sub AssertBindNameStartPrefix(ByVal bindName As String)

            If Not bindName.StartsWith(BIND_NAME_PREFIX) Then
                Throw New ArgumentException("バインド変数名は '" & BIND_NAME_PREFIX & "' から始まる", "bindName")
            End If
        End Sub

        ''' <summary>
        ''' Indexバインド変数名に使用する区切り文字かを返す
        ''' </summary>
        ''' <param name="c">判定文字</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Shared Function IsInternalBindNameIndexSeparator(ByVal c As Char) As Boolean
            Return INTERNAL_BIND_NAME_INDEX_SEPARATOR = c
        End Function

        ''' <summary>
        ''' （内部的な）バインド変数名の区切り文字かを返す
        ''' </summary>
        ''' <param name="c">判定文字</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Shared Function IsInternalBindNameNestSeparator(ByVal c As Char) As Boolean
            Return INTERNAL_BIND_NAME_PROPERTY_SEPARATOR = c
        End Function

        ''' <summary>
        ''' （内部的な）Indexバインド変数名かを返す
        ''' </summary>
        ''' <param name="bindName">バインド変数名</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Private Shared Function IsInternalIndexBindName(ByVal bindName As String) As Boolean
            Return 0 <= bindName.IndexOf(INTERNAL_BIND_NAME_INDEX_SEPARATOR)
        End Function

        ''' <summary>
        ''' （内部的な）Indexバインド変数名を、Indexプロパティ名にして返す
        ''' </summary>
        ''' <param name="indexBindName">Indexバインド変数名</param>
        ''' <returns>Indexプロパティ名</returns>
        ''' <remarks></remarks>
        Private Shared Function ConvInternalIndexBindNameToPropertyName(ByVal indexBindName As String) As String

            Return indexBindName.Replace(INTERNAL_BIND_NAME_INDEX_SEPARATOR, INDEX_PROPERTY_OPEN) & INDEX_PROPERTY_CLOSE
        End Function

        ''' <summary>
        ''' バインド変数名をプロパティ名にして返す（内部的なバインド変数名も含む）@Hoge#1  ->  Hoge(1)
        ''' </summary>
        ''' <param name="bindName">バインド変数名</param>
        ''' <returns>プロパティ名</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvBindNameToPropertyName(ByVal bindName As String) As String
            AssertBindNameStartPrefix(bindName)

            Dim propertyResult As New List(Of String)
            For Each part As String In Split(bindName.Substring(1), INTERNAL_BIND_NAME_PROPERTY_SEPARATOR)
                If IsInternalIndexBindName(part) Then
                    propertyResult.Add(ConvInternalIndexBindNameToPropertyName(part))
                Else
                    propertyResult.Add(part)
                End If
            Next
            Return Join(propertyResult.ToArray, PROPERTY_SEPARATOR)
        End Function

        ''' <summary>
        ''' プロパティ名の文字か？を返す
        ''' </summary>
        ''' <param name="c">文字</param>
        ''' <returns>プロパティ名ならtrue</returns>
        ''' <remarks></remarks>
        Public Shared Function IsPropertyNameChar(ByVal c As Char) As Boolean
            Return ("a"c <= c AndAlso c <= "z"c) OrElse ("A"c <= c AndAlso c <= "Z"c) OrElse ("0"c <= c AndAlso c <= "9"c) OrElse 0 <= "_".IndexOf(c)
        End Function

        ''' <summary>
        ''' 内部的なバインド変数名を表す文字か？を返す
        ''' </summary>
        ''' <param name="c">文字</param>
        ''' <returns>判定結果</returns>
        ''' <remarks></remarks>
        Public Shared Function IsInternalBindNameChar(ByVal c As Char) As Boolean

            Return IsPropertyNameChar(c) OrElse IsInternalBindNameNestSeparator(c) OrElse IsInternalBindNameIndexSeparator(c)
        End Function

        ''' <summary>
        ''' ユーザー入力バインド変数名を探索するRegexを作成する
        ''' </summary>
        ''' <returns>ユーザー入力バインド変数名を探索するRegex</returns>
        ''' <remarks></remarks>
        Public Shared Function NewUserBindNameRegex() As Regex
            Return New Regex("@[A-Z][A-Z0-9.]*(\([0-9]*\))?(\.[A-Z][A-Z0-9.]*(\([0-9]*\))?)*", RegexOptions.IgnoreCase)
        End Function

        ''' <summary>
        ''' ユーザー入力バインド変数名があれば内部的バインド変数名にする
        ''' </summary>
        ''' <param name="textValue">バインド変数名を含む可能性のある文字列</param>
        ''' <returns>倍部的バインド変数名にした文字列</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvInternalBindNameIfNecessary(ByVal textValue As String) As String
            Return ConvInternalBindNameIfNecessary(textValue, NewUserBindNameRegex)
        End Function

        ''' <summary>
        ''' ユーザー入力バインド変数名があれば内部的バインド変数名にする
        ''' </summary>
        ''' <param name="textValue">バインド変数名を含む可能性のある文字列</param>
        ''' <param name="userBindNameRegex">ユーザー入力バインド変数名を探索するRegex</param>
        ''' <returns>内部的バインド変数名にした文字列</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvInternalBindNameIfNecessary(ByVal textValue As String, ByVal userBindNameRegex As Regex) As String
            Dim bindNameMatch As Match = userBindNameRegex.Match(textValue)
            Dim result As New StringBuilder
            Dim start As Integer = 0
            While bindNameMatch.Success
                result.Append(textValue.Substring(start, bindNameMatch.Index - start))
                Dim matchStr As String = textValue.Substring(bindNameMatch.Index, bindNameMatch.Length)
                result.Append(SqlBindUtil.ConvBindNameUserToInternal(matchStr))
                start = bindNameMatch.Index + bindNameMatch.Length
                bindNameMatch = bindNameMatch.NextMatch()
            End While
            result.Append(textValue.Substring(start))
            Return result.ToString
        End Function
    End Class
End Namespace