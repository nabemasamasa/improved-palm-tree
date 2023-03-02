Imports System.Text
Imports TechTalk.SpecFlow

Namespace TestUtil
    ''' <summary>
    ''' SpecFlowのためのユーティリティクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SpecFlowUtil

        Private Shared ReadOnly FUNCTIONS As IFunction() = {New DateFunction, New ZenkakuFunction, New HankakuFunction, New GetFunction, New UserIdFunction, New ExpandFunction, New LeftFunction}

#Region "Nested classes..."
        Private Interface IFunction
            Function GetName() As String
            Function GetDefaultArg() As String
            Function Compute(ByVal arg As String) As String
        End Interface
        Private Class DateFunction : Implements IFunction

            Public Function GetName() As String Implements IFunction.GetName
                Return "@DATE"
            End Function

            Public Function GetDefaultArg() As String Implements IFunction.GetDefaultArg
                Return "yyyy/MM/dd"
            End Function

            Public Function Compute(ByVal arg As String) As String Implements IFunction.Compute
                If IsNumeric(arg) Then
                    Return DateTime.Now.AddDays(CInt(arg)).ToString(GetDefaultArg)
                End If
                Return DateTime.Now.ToString(If(String.IsNullOrEmpty(arg), GetDefaultArg, arg))
            End Function
        End Class
        Private Class ZenkakuFunction : Implements IFunction

            Public Function GetName() As String Implements IFunction.GetName
                Return "@ZENKAKU"
            End Function

            Public Function GetDefaultArg() As String Implements IFunction.GetDefaultArg
                Throw New ArgumentException("@Zenkakuに引数は必須")
            End Function

            Public Function Compute(ByVal arg As String) As String Implements IFunction.Compute
                Return StringUtil.ToZenkaku(arg)
            End Function
        End Class
        Private Class HankakuFunction : Implements IFunction

            Public Function GetName() As String Implements IFunction.GetName
                Return "@HANKAKU"
            End Function

            Public Function GetDefaultArg() As String Implements IFunction.GetDefaultArg
                Throw New ArgumentException("@Hankakuに引数は必須")
            End Function

            Public Function Compute(ByVal arg As String) As String Implements IFunction.Compute
                Return StringUtil.ToHankaku(arg)
            End Function
        End Class
        Private Class GetFunction : Implements IFunction

            Public Function GetName() As String Implements IFunction.GetName
                Return "@GET"
            End Function

            Public Function GetDefaultArg() As String Implements IFunction.GetDefaultArg
                Throw New ArgumentException("@Getに引数は必須")
            End Function

            Public Function Compute(ByVal arg As String) As String Implements IFunction.Compute
                Return GetCurrentScenarioValue(Of String)(arg)
            End Function
        End Class
        Private Class UserIdFunction : Implements IFunction

            Public Function GetName() As String Implements IFunction.GetName
                Return "@USERID"
            End Function

            Public Function GetDefaultArg() As String Implements IFunction.GetDefaultArg
                Return Nothing
            End Function

            Public Function Compute(ByVal arg As String) As String Implements IFunction.Compute
                Return Fhi.Fw.Sys.ComFunc.User.GetName
            End Function
        End Class
        Private Class ExpandFunction : Implements IFunction

            Public Function GetName() As String Implements IFunction.GetName
                Return "@EXPAND"
            End Function

            Public Function GetDefaultArg() As String Implements IFunction.GetDefaultArg
                Throw New ArgumentException("@EXPANDに引数は必須")
            End Function

            Public Function Compute(ByVal arg As String) As String Implements IFunction.Compute
                Return Environment.ExpandEnvironmentVariables(arg)
            End Function
        End Class
        Private Class LeftFunction : Implements IFunction

            Public Function GetName() As String Implements IFunction.GetName
                Return "@LEFT"
            End Function

            Public Function GetDefaultArg() As String Implements IFunction.GetDefaultArg
                Throw New ArgumentException("@Leftに引数は必須")
            End Function

            Public Function Compute(ByVal arg As String) As String Implements IFunction.Compute
                Dim args As String() = arg.Split(","c)
                If args.Length <> 2 OrElse Not IsNumeric(args(1)) Then
                    Throw New ArgumentException("引数は @Left(変数, 長さ) にすべき", "@Left(" & arg & ")")
                End If
                Dim value As String = GetCurrentScenarioValue(Of String)(args(0))
                Return Left(value, CInt(args(1)))
            End Function
        End Class

#End Region

        ''' <summary>
        ''' ScenarioContextを初期化する
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub InitializeScenarioContext(Optional ByVal scenarioTitle As String = "SpecFlowUtilScenarioTitle")
            TestRunnerManager.GetTestRunner.OnScenarioStart(New ScenarioInfo(scenarioTitle))
        End Sub

        ''' <summary>
        ''' FeatureContextを初期化する
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub InitializeFeatureContext(Optional ByVal featureTitle As String = "SpecFlowUtilFeatureTitle", Optional ByVal featuireDescription As String = "SpecFlowUtilFeatureDescription")
            TestRunnerManager.GetTestRunner.OnFeatureStart(New FeatureInfo(New Globalization.CultureInfo("en-US"), featureTitle, featuireDescription))
        End Sub

        ''' <summary>
        ''' ScenarioContextを破棄する
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub FinalizeScenarioContext()
            TestRunnerManager.GetTestRunner.OnScenarioEnd()
        End Sub

        ''' <summary>
        ''' FeautureContextを破棄する
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub FinalizeFeatureContext()
            TestRunnerManager.GetTestRunner.OnFeatureEnd()
        End Sub

        ''' <summary>
        ''' 現在のシナリオで保持した変数値を返す
        ''' </summary>
        ''' <typeparam name="T">保持した型</typeparam>
        ''' <param name="varName">変数名</param>
        ''' <returns>変数値</returns>
        ''' <remarks></remarks>
        Public Shared Function GetCurrentScenarioValue(Of T)(ByVal varName As String) As T
            If Not ScenarioContext.Current.ContainsKey(varName) Then
                Throw New ArgumentException(String.Format("変数 ""{0}"" はみつからない", varName))
            End If
            Return ScenarioContext.Current.Get(Of T)(varName)
        End Function

        ''' <summary>
        ''' 現在のシナリオに値を保持する
        ''' </summary>
        ''' <param name="varName">変数名</param>
        ''' <param name="value">値</param>
        ''' <remarks></remarks>
        Public Shared Sub SetCurrentScenarioValue(ByVal varName As String, ByVal value As Object)
            If ScenarioContext.Current.ContainsKey(varName) Then
                ScenarioContext.Current.Remove(varName)
            End If
            ScenarioContext.Current.Add(varName, value)
        End Sub

        ''' <summary>
        ''' 関数を含め評価する
        ''' </summary>
        ''' <param name="value">文字列値</param>
        ''' <returns>結果を反映した値</returns>
        ''' <remarks></remarks>
        Public Shared Function EvaluateValue(ByVal value As String) As String
            Const OPEN_BRACKET As String = "("
            Const FUNCTION_MARK As Char = "@"c
            Dim index As Integer = value.IndexOf(FUNCTION_MARK)
            If index < 0 Then
                Return value
            End If

            Dim result As New StringBuilder(value.Substring(0, index))
            Dim cutValueFromAt As String = value.Substring(index)
            Dim aFunction As IFunction = GetIFunction(cutValueFromAt)

            Dim functionArg As String
            Dim appendStr As String
            Dim functionName As String = aFunction.GetName
            If functionName.Length < cutValueFromAt.Length AndAlso OPEN_BRACKET.Equals(cutValueFromAt.Substring(functionName.Length, 1)) Then
                Dim cutValueFromOpenBracket As String = cutValueFromAt.Substring(functionName.Length)
                Dim pairBracketIndex As Integer = GetIndexOfPairBracket(cutValueFromOpenBracket)
                functionArg = EvaluateValue(cutValueFromOpenBracket.Substring(OPEN_BRACKET.Length, pairBracketIndex - OPEN_BRACKET.Length))
                appendStr = cutValueFromOpenBracket.Substring(pairBracketIndex + 1)
            Else
                functionArg = aFunction.GetDefaultArg
                appendStr = cutValueFromAt.Substring(functionName.Length)
            End If
            result.Append(aFunction.Compute(functionArg)).Append(EvaluateValue(appendStr))
            Return result.ToString
        End Function

        ''' <summary>
        ''' 内部関数に一致する IFunction を返す
        ''' </summary>
        ''' <param name="value">内部関数を使用している文字列値</param>
        ''' <returns>一致するIFunction</returns>
        ''' <remarks></remarks>
        Private Shared Function GetIFunction(ByVal value As String) As IFunction
            For Each aFunction As IFunction In FUNCTIONS
                If value.StartsWith(aFunction.GetName, StringComparison.OrdinalIgnoreCase) Then
                    Return aFunction
                End If
            Next
            Throw New ArgumentException("一致するFunctionが見つからない " & value)
        End Function

        ''' <summary>
        ''' 最初の開き括弧に対応する閉じ括弧の位置indexを返す
        ''' </summary>
        ''' <param name="value">文字列値</param>
        ''' <returns>閉じ括弧の位置index</returns>
        ''' <remarks></remarks>
        Private Shared Function GetIndexOfPairBracket(ByVal value As String) As Integer
            Const OPEN_BRACKET As Char = "("c
            Const CLOSE_BRACKET As Char = ")"c
            Dim chars As Char() = value.ToCharArray
            Dim count As Integer = 0
            For i As Integer = 0 To chars.Length - 1
                Dim c As Char = chars(i)
                If c = OPEN_BRACKET Then
                    count += 1
                ElseIf c = CLOSE_BRACKET Then
                    count -= 1
                End If
                If 0 < count Then
                    Continue For
                End If
                Return i
            Next
            Throw New ArgumentException(value & "に対応する括弧が見つからない")
        End Function

        ''' <summary>
        ''' SpecFlowのTable構造を二段階配列にする
        ''' </summary>
        ''' <param name="aTable">Table構造の値</param>
        ''' <returns>二段階配列</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvTableToTwoJaggedArray(ByVal aTable As Table) As String()()
            Dim titles As String() = aTable.Header.ToArray
            Return aTable.Rows.Select(Function(row) Enumerable.Range(0, titles.Length).Select(Function(i) EvaluateValue(row(titles(i)))).ToArray).ToArray
        End Function

        ''' <summary>
        ''' 「ハンドルが無効」の例外か？
        ''' </summary>
        ''' <param name="ex">Win32エラーの例外</param>
        ''' <returns>エラーコードがハンドル無効ならTRUE</returns>
        ''' <remarks></remarks>
        Public Shared Function IsErrorInvalidHandle(ByVal ex As ComponentModel.Win32Exception) As Boolean
            Const ERROR_INVALID_HANDLE As Integer = 6
            Return ex.NativeErrorCode = ERROR_INVALID_HANDLE
        End Function

    End Class
End Namespace