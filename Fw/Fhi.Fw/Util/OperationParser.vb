Imports System.Text

Namespace Util
    ''' <summary>
    ''' 操作文字列の解析を担うクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class OperationParser
#Region "Nested classes..."
        ''' <summary>
        ''' 解析内容
        ''' </summary>
        ''' <remarks></remarks>
        Public Class Operation
            ''' <summary>親</summary>
            Public Parent As Operation
            ''' <summary>操作</summary>
            Public Member As String
            ''' <summary>操作の引数</summary>
            Public Args As Operation()
        End Class
        Private Class InternalParseResult
            Public EndIndex As Integer
            Public Operation As Operation
        End Class
#End Region

        Private Shared Function RecurParse(ByVal operationChars As Char(), ByVal startIndex As Integer, ByVal parent As Operation) As InternalParseResult
            Const OPEN_BRACKET As Char = "("c
            Const CLOSE_BRACKET As Char = ")"c
            Const OPERATION_SEPARATOR As Char = "."c
            Const PARAMETER_SEPARATOR As Char = ","c
            Dim result As New InternalParseResult With {.EndIndex = operationChars.Length - 1, .Operation = New Operation With {.Parent = parent}}
            Dim args As New List(Of Operation)
            Dim requresArgs As Boolean = False
            Dim sb As New StringBuilder
            For i As Integer = startIndex To operationChars.Length - 1
                Select Case operationChars(i)
                    Case OPERATION_SEPARATOR
                        If StringUtil.IsEmpty(result.Operation.Member) Then
                            Dim member As String = sb.ToString.Trim
                            If member.Length <= 0 Then
                                Throw New ArgumentException("ピリオドの位置が不正です")
                            End If
                            result.Operation.Member = member
                        End If
                        If requresArgs Then
                            result.Operation.Args = args.ToArray
                        End If
                        Return RecurParse(operationChars, i + 1, result.Operation)

                    Case OPEN_BRACKET
                        Dim member As String = sb.ToString.Trim
                        If member.Length <= 0 Then
                            Throw New ArgumentException("Member無しの括弧始まりは処理しない", "operationChars")
                        End If
                        result.Operation.Member = member
                        requresArgs = True
                        sb.Length = 0
                        Do
                            Dim aParseResult As InternalParseResult = RecurParse(operationChars, i + 1, Nothing)
                            If StringUtil.IsNotEmpty(aParseResult.Operation.Member) Then
                                args.Add(aParseResult.Operation)
                            End If
                            i = aParseResult.EndIndex
                        Loop While operationChars(i) = PARAMETER_SEPARATOR

                    Case CLOSE_BRACKET
                        Dim member As String = sb.ToString.Trim
                        If 0 < member.Length Then
                            result.Operation.Member = member
                            result.EndIndex = i
                            Return result
                        End If
                        sb.Length = 0

                    Case PARAMETER_SEPARATOR
                        result.EndIndex = i
                        Exit For

                    Case Else
                        sb.Append(operationChars(i))
                End Select
            Next
            If 0 < sb.Length Then
                result.Operation.Member = sb.ToString.Trim
            End If
            If requresArgs Then
                result.Operation.Args = args.ToArray
            End If
            Return result
        End Function

        ''' <summary>
        ''' 操作を解析する
        ''' </summary>
        ''' <param name="operation">操作</param>
        ''' <returns>解析結果</returns>
        ''' <remarks></remarks>
        Public Shared Function Parse(ByVal operation As String) As Operation
            AssertBraketPair(operation)
            Return RecurParse(operation.ToCharArray, startIndex:=0, parent:=Nothing).Operation
        End Function

        Private Shared Sub AssertBraketPair(ByVal operation As String)
            Dim chars As Char() = operation.ToCharArray
            Dim isString As Boolean = False
            Dim countOfOpenBrakets As Integer = 0
            For Each c As Char In chars
                Select Case c
                    Case "("c
                        If isString Then
                            Continue For
                        End If
                        countOfOpenBrakets += 1
                    Case ")"c
                        If isString Then
                            Continue For
                        End If
                        countOfOpenBrakets -= 1
                        If countOfOpenBrakets < 0 Then
                            Throw New ArgumentException("括弧を開いて無いのに閉じている", "operation")
                        End If
                    Case """"c
                        isString = Not isString
                End Select
            Next
            If 0 < countOfOpenBrakets Then
                Throw New ArgumentException("括弧を開きっぱなし", "operation")
            End If
        End Sub

    End Class
End Namespace