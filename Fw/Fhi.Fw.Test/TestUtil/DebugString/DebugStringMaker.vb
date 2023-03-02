Imports System.Text
Imports Fhi.Fw.Domain
Imports Fhi.Fw.Util
Imports System.Reflection

Namespace TestUtil.DebugString
    ''' <summary>
    ''' 検証文字列作成を担うクラス
    ''' </summary>
    ''' <typeparam name="T">Voの型</typeparam>
    ''' <remarks></remarks>
    Public Class DebugStringMaker(Of T) : Implements DebugStringMakerWildcard

#Region "Nested classes..."
        Private Class Length
            Public [String] As Integer = 0
            Public [Integer] As Integer = 0
            Public [Decimal] As Integer = 0
        End Class
        Private Class ForSideBySide
            Public Property Values As T()
        End Class
#End Region
        Private ReadOnly builder As DebugStringRuleBuilder(Of T)
        Private ReadOnly aConfigure As DebugStringRuleBuilder(Of T).Configure

        ''' <summary>数値型でも左揃えにする場合、true</summary>
        Public AlignsLeftTheNumeric As Boolean = False

        ''' <summary>小数点以下の表示桁数</summary>
        Public DecimalLength As Integer = 5

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="aConfigure"> 検証文字列の列設定</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal aConfigure As DebugStringRuleBuilder(Of T).Configure)
            Me.aConfigure = aConfigure
            builder = New DebugStringRuleBuilder(Of T)(aConfigure)
        End Sub

        ''' <summary>
        ''' 検証用値情報のタイトル一覧を作成する
        ''' </summary>
        ''' <param name="parentTitle">親タイトル</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MakeTitles(Optional parentTitle As String = Nothing) As String() Implements DebugStringMakerWildcard.MakeTitles
            Return builder.MakeTitles(parentTitle)
        End Function

        ''' <summary>
        ''' 検証用値情報を作成して格納する
        ''' </summary>
        ''' <param name="record"></param>
        ''' <remarks></remarks>
        Public Sub StoreAfterMaking(ByVal record As Object) Implements DebugStringMakerWildcard.StoreAfterMaking
            builder.StoreAfterMaking(DirectCast(record, T))
        End Sub

        ''' <summary>
        ''' 格納情報から値情報を構築して格納情報をクリアするCallbackを返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCallbackThatBuildValuesAndClearStored() As Func(Of String()()) Implements DebugStringMakerWildcard.GetCallbackThatBuildValuesAndClearStored
            Return builder.GetCallbackThatBuildValuesAndClearStored
        End Function

        ''' <summary>
        ''' Empty値の値情報作成Callbackを返す
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCallbackThatMakeEmptyValues() As Func(Of String()) Implements DebugStringMakerWildcard.GetCallbackThatMakeEmptyValues
            Return builder.GetCallbackThatMakeEmptyValues
        End Function

        ''' <summary>
        ''' 横並びで検証用文字列を作成する
        ''' </summary>
        ''' <param name="records">値情報[]</param>
        ''' <returns>検証用文字列</returns>
        ''' <remarks></remarks>
        Public Function MakeSideBySide(ByVal records As IEnumerable(Of T)) As String
            Return MakeSideBySide(records.ToArray)
        End Function
        ''' <summary>
        ''' 横並びで検証用文字列を作成する
        ''' </summary>
        ''' <param name="records">値情報[]</param>
        ''' <returns>検証用文字列</returns>
        ''' <remarks></remarks>
        Public Function MakeSideBySide(ByVal ParamArray records As T()) As String
            Dim maker As New DebugStringMaker(Of ForSideBySide)(
                Function(defineBy As IDebugStringRuleBinder, instance As ForSideBySide) _
                    defineBy.JoinFromSideBySide(instance.Values, aConfigure))
            maker.builder.SuppressesParentColumnName = True
            Return maker.MakeString(New ForSideBySide With {.Values = records})
        End Function

        ''' <summary>
        ''' 検証用文字列を作成する
        ''' </summary>
        ''' <param name="records">値情報[]</param>
        ''' <returns>検証用文字列</returns>
        ''' <remarks></remarks>
        Public Function MakeString(ByVal ParamArray records As T()) As String
            Return MakeString(DirectCast(records, IEnumerable(Of T)))
        End Function

        ''' <summary>
        ''' 検証用文字列を作成する
        ''' </summary>
        ''' <param name="records">値情報[]</param>
        ''' <returns>検証用文字列</returns>
        ''' <remarks></remarks>
        Public Function MakeString(ByVal records As IEnumerable(Of T)) As String
            Dim matrix As New List(Of String())
            For Each record As T In records
                builder.StoreAfterMaking(record)
            Next
            matrix.AddRange(builder.BuildValuesWithStored)
            matrix.Insert(0, builder.MakeTitles)
            Dim maxLengthesByColumn As New List(Of Length)
            For Each values As String() In matrix
                For i As Integer = 0 To values.Count - 1
                    If maxLengthesByColumn.Count <= i Then
                        maxLengthesByColumn.Add(New Length)
                    End If
                    Dim value As String = values(i)
                    If Not AlignsLeftTheNumeric AndAlso IsNumeric(value) Then
                        Dim aDec As Decimal = CDec(value)
                        Dim decimalSize As Integer = EzUtil.GetDecimalSize(aDec)
                        maxLengthesByColumn(i).Decimal = Math.Max(maxLengthesByColumn(i).Decimal, decimalSize)
                        maxLengthesByColumn(i).Integer = Math.Max(maxLengthesByColumn(i).Integer, StringUtil.GetLengthByte(Decimal.Floor(aDec).ToString))
                    Else
                        maxLengthesByColumn(i).String = Math.Max(maxLengthesByColumn(i).String, StringUtil.GetLengthByte(value))
                    End If
                Next
            Next
            Const SPACE As Char = " "c
            Const SPACE_LENGTH As Integer = 1
            Const CRLF_LENGTH As Integer = 2
            Const DECIMAL_POINT As Char = "."c
            Const DECIMAL_POINT_LENGTH As Integer = 1
            Dim isContents As Boolean = False
            Dim result As New StringBuilder
            For Each values As String() In matrix
                For i As Integer = 0 To values.Count - 1
                    Dim maxDecimalLength As Integer = Math.Min(maxLengthesByColumn(i).Decimal, Me.DecimalLength)
                    Dim maxLength As Integer = Math.Max(maxLengthesByColumn(i).String, maxLengthesByColumn(i).Integer + If(0 < maxDecimalLength, maxDecimalLength + DECIMAL_POINT_LENGTH, 0))
                    If Not AlignsLeftTheNumeric AndAlso IsNumeric(values(i)) AndAlso isContents Then
                        Dim value As Decimal = CDec(values(i))
                        Dim decimalSize As Integer = EzUtil.GetDecimalSize(value)
                        Dim lengthOfInteger As Integer = maxLength - If(0 < maxDecimalLength, maxDecimalLength + DECIMAL_POINT_LENGTH, 0)
                        result.Append(StringUtil.MakeFixedString(Decimal.Floor(value).ToString, lengthOfInteger, alignsRight:=True))
                        If 0 < maxDecimalLength Then
                            If 0 < decimalSize Then
                                result.Append(DECIMAL_POINT).Append(StringUtil.MakeFixedString(StringUtil.OmitIfLengthOver(Right(value.ToString, decimalSize), maxDecimalLength, "…"), maxDecimalLength))
                            Else
                                result.Append(SPACE, maxDecimalLength + DECIMAL_POINT_LENGTH)
                            End If
                        End If
                    Else
                        result.Append(StringUtil.MakeFixedString(values(i), maxLength))
                    End If
                    result.Append(SPACE)
                Next
                If SPACE_LENGTH < result.Length Then
                    result.Remove(result.Length - SPACE_LENGTH, SPACE_LENGTH)
                End If
                result.Append(vbCrLf)
                isContents = True
            Next
            If CRLF_LENGTH < result.Length Then
                result.Remove(result.Length - CRLF_LENGTH, CRLF_LENGTH)
            End If
            Return result.ToString
        End Function

    End Class
    Friend Class DebugStringMaker
        ''' <summary>
        ''' 検証用文字列にする
        ''' </summary>
        ''' <param name="value">値</param>
        ''' <returns>検証用文字列</returns>
        ''' <remarks></remarks>
        Friend Overloads Shared Function ConvDebugValue(ByVal value As Object) As String
            If value Is Nothing Then
                Return "null"
            End If
            Dim str As String = TryCast(value, String)
            If str IsNot Nothing Then
                Return ToStringValue(str)
            End If
            Dim aType As Type = TypeUtil.GetTypeIfNullable(value.GetType)
            If aType Is GetType(DateTime) Then
                Return ConvDebugValue(StringUtil.ToDateTimeString(DirectCast(value, DateTime)))
            ElseIf GetType(ValueObject).IsAssignableFrom(aType) Then
                If GetType(PrimitiveValueObject).IsAssignableFrom(aType) Then
                    Return ConvDebugValue(DirectCast(value, PrimitiveValueObject).Value)
                End If
                Const ATTRR As BindingFlags = Reflection.BindingFlags.Public Or Reflection.BindingFlags.Instance
                Dim fieldInfos As FieldInfo() = aType.GetFields(ATTRR)
                Dim propertyInfos As PropertyInfo() = aType.GetProperties(ATTRR)
                If fieldInfos.Length = 1 AndAlso propertyInfos.Length = 0 Then
                    Return ConvDebugValue(fieldInfos(0).GetValue(value))
                ElseIf fieldInfos.Length = 0 AndAlso propertyInfos.Length = 1 Then
                    Return ConvDebugValue(propertyInfos(0).GetValue(value, Nothing))
                ElseIf fieldInfos.Length = 0 AndAlso propertyInfos.Length = 0 Then
                    Dim propertyInfoOfValue As PropertyInfo = VoPropertyMarker.GetPropertyInfoOfValueByType(aType)
                    If propertyInfoOfValue IsNot Nothing AndAlso propertyInfoOfValue.CanRead Then
                        Return ConvDebugValue(propertyInfoOfValue.GetValue(value, Nothing))
                    End If
                    Dim fieldInfoOfValue As FieldInfo = VoPropertyMarker.GetFieldInfoOfValueByType(aType)
                    If fieldInfoOfValue IsNot Nothing Then
                        Return ConvDebugValue(fieldInfoOfValue.GetValue(value))
                    End If
                End If
                Return ToStringValue(value.ToString())
            End If
            Return StringUtil.ToString(value)
        End Function

        Private Shared Function ToStringValue(value As String) As String
            Const SQ As Char = "'"c
            Dim sb As New StringBuilder
            Return sb.Append(SQ).Append(value.ToString().Replace(vbCrLf, "\n")).Append(SQ).ToString()
        End Function
    End Class
End Namespace